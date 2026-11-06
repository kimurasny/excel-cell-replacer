"""Excel文字列置換・複製CLIツール。

CSVで定義した置換ルールに従い、Excelファイル（.xlsx / .xlsm）の指定セル内の
文字列を置換して、別名のExcelファイルとして保存する。

使用例:
    python excel_replace.py --input template.xlsx --csv replace.csv
"""

from __future__ import annotations

import argparse
import csv
import os
import re
import sys
import tempfile
import zipfile
from collections.abc import Sequence
from dataclasses import dataclass
from pathlib import Path
from typing import NoReturn, Union

from openpyxl import Workbook, load_workbook
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE, Cell, MergedCell
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.utils.cell import column_index_from_string
from openpyxl.worksheet.worksheet import Worksheet

EXIT_SUCCESS = 0
EXIT_INPUT_ERROR = 1
EXIT_RUNTIME_ERROR = 2

SUPPORTED_EXTENSIONS = (".xlsx", ".xlsm")
CSV_COLUMNS = ("sheet", "cell", "before", "after")
CSV_ENCODING = "utf-8-sig"

MAX_ROW = 1_048_576
MAX_COLUMN = 16_384
MAX_CELL_TEXT_LENGTH = 32_767
CELL_ADDRESS_RE = re.compile(r"^([A-Za-z]{1,3})([1-9][0-9]{0,6})$")

CellText = Union[str, CellRichText]

# openpyxl で読み込み・保存すると失われる要素の検出パターン（パーツ名の正規表現, 内容の正規表現, 説明）
UNSUPPORTED_FEATURES: tuple[tuple[re.Pattern[str], re.Pattern[bytes] | None, str], ...] = (
    (
        re.compile(r"^xl/drawings/drawing\d+\.xml$"),
        re.compile(rb"<(\w+:)?(sp|cxnSp|grpSp)[\s>]"),
        "図形・テキストボックス",
    ),
    (
        re.compile(r"^xl/worksheets/sheet\d+\.xml$"),
        re.compile(rb"<(\w+:)?controls[\s>]"),
        "フォームコントロール・ActiveXコントロール",
    ),
    (
        re.compile(r"^xl/worksheets/sheet\d+\.xml$"),
        re.compile(rb"<(\w+:)?extLst[\s>]"),
        "シートの拡張情報（スパークライン、スライサー、新しい形式の条件付き書式・入力規則など）",
    ),
    (re.compile(r"^customUI(14)?/"), None, "リボンのカスタマイズ（customUI）"),
)


@dataclass(frozen=True)
class ReplaceRule:
    """CSVの1行分の置換ルール。"""

    line_no: int
    sheet: str
    cell: str
    before: str
    after: str


@dataclass(frozen=True)
class RuleError:
    """入力内容・置換ルールのエラー。line_no が None の場合は行に依存しないエラー。"""

    message: str
    line_no: int | None = None

    def format(self) -> str:
        if self.line_no is None:
            return self.message
        return f"[{self.line_no}行目] {self.message}"


class ExecutionError(Exception):
    """ファイルの読み込み・保存など、実行時のエラー。"""


class InputError(Exception):
    """入力内容・置換ルールのエラー（複数件をまとめて保持する）。"""

    def __init__(self, errors: list[RuleError]) -> None:
        super().__init__(f"{len(errors)} 件のエラー")
        self.errors = errors


class _ArgumentParser(argparse.ArgumentParser):
    """引数エラー時の終了コードを 1（入力内容のエラー）にするための ArgumentParser。"""

    def error(self, message: str) -> NoReturn:
        self.print_usage(sys.stderr)
        self.exit(EXIT_INPUT_ERROR, f"ERROR: コマンドライン引数が不正です。{message}\n")


def build_parser() -> argparse.ArgumentParser:
    parser = _ArgumentParser(
        prog="excel_replace.py",
        description="CSVの定義に従ってExcelの指定セルの文字列を置換し、別名で保存します。",
    )
    parser.add_argument("--input", required=True, help="複製元Excelファイル（.xlsx / .xlsm）")
    parser.add_argument("--csv", required=True, help="置換定義CSVファイル（UTF-8）")
    parser.add_argument(
        "--output",
        help="出力Excelファイルのパス（省略時はCSVと同じディレクトリに「CSVのベース名＋入力Excelの拡張子」で出力）",
    )
    return parser


def resolve_output_path(input_path: Path, csv_path: Path, output: str | None) -> Path:
    """出力先パスを決定する。未指定の場合はCSVのベース名と入力Excelの拡張子を使う。"""
    if output:
        return Path(output)
    return csv_path.with_name(csv_path.stem + input_path.suffix)


def validate_paths(input_path: Path, csv_path: Path, output_path: Path) -> list[RuleError]:
    """入力Excel・CSV・出力先のパスを検証し、エラーの一覧を返す。"""
    errors: list[RuleError] = []
    input_ext = input_path.suffix.lower()

    if input_ext not in SUPPORTED_EXTENSIONS:
        errors.append(
            RuleError(
                f"入力Excelファイルの拡張子「{input_path.suffix}」は対象外です"
                f"（対象: {' / '.join(SUPPORTED_EXTENSIONS)}）: {input_path}"
            )
        )
    elif not input_path.exists():
        errors.append(RuleError(f"入力Excelファイルが存在しません: {input_path}"))
    elif not input_path.is_file():
        errors.append(RuleError(f"入力Excelファイルがファイルではありません: {input_path}"))

    if not csv_path.exists():
        errors.append(RuleError(f"CSVファイルが存在しません: {csv_path}"))
    elif not csv_path.is_file():
        errors.append(RuleError(f"CSVファイルがファイルではありません: {csv_path}"))

    if input_ext in SUPPORTED_EXTENSIONS and output_path.suffix.lower() != input_ext:
        errors.append(
            RuleError(f"出力Excelファイルの拡張子は入力Excelと同じ「{input_path.suffix}」にしてください: {output_path}")
        )
    if output_path.exists():
        errors.append(RuleError(f"出力ファイルが既に存在します: {output_path}"))
    elif not output_path.parent.is_dir():
        errors.append(RuleError(f"出力先ディレクトリが存在しません: {output_path.parent}"))

    return errors


def parse_cell_address(address: str) -> tuple[int, int] | None:
    """A1形式のセル番地を (行, 列) に変換する。不正な場合は None を返す。"""
    match = CELL_ADDRESS_RE.match(address)
    if not match:
        return None
    column = column_index_from_string(match.group(1).upper())
    row = int(match.group(2))
    if column > MAX_COLUMN or row > MAX_ROW:
        return None
    return row, column


def read_rules(csv_path: Path) -> tuple[list[ReplaceRule], list[RuleError]]:
    """置換定義CSVを読み込み、置換ルールと行単位のエラーを返す。

    ヘッダー不正・置換ルールなしなど、CSV全体に関わるエラーの場合は InputError を送出する。
    """
    try:
        with csv_path.open("r", encoding=CSV_ENCODING, newline="") as f:
            reader = csv.reader(f, strict=True)
            rows: list[tuple[int, list[str]]] = []
            try:
                for row in reader:
                    rows.append((reader.line_num, row))
            except csv.Error as e:
                raise InputError([RuleError(f"CSVの形式が不正です（{e}）。", reader.line_num)]) from e
    except UnicodeDecodeError as e:
        raise InputError([RuleError(f"CSVファイルをUTF-8として読み込めません（{e.start}バイト目）: {csv_path}")]) from e
    except OSError as e:
        raise ExecutionError(f"CSVファイルの読み込みに失敗しました: {csv_path}（{e}）") from e
    return _parse_rules(rows)


def _parse_rules(rows: Sequence[tuple[int, list[str]]]) -> tuple[list[ReplaceRule], list[RuleError]]:
    if not rows:
        raise InputError([RuleError("CSVファイルが空です。ヘッダー行が必要です。")])
    header_line_no, header = rows[0]
    header_names = [name.strip() for name in header]
    if sorted(header_names) != sorted(CSV_COLUMNS):
        raise InputError(
            [
                RuleError(
                    f"CSVのヘッダーが不正です。「{','.join(CSV_COLUMNS)}」の4列が必要です"
                    f"（実際: 「{','.join(header)}」）。",
                    line_no=header_line_no,
                )
            ]
        )
    index = {name: header_names.index(name) for name in CSV_COLUMNS}

    rules: list[ReplaceRule] = []
    errors: list[RuleError] = []
    for line_no, row in rows[1:]:
        if all(value == "" for value in row):
            continue  # 空行は無視する
        if len(row) != len(CSV_COLUMNS):
            errors.append(RuleError(f"列数が不正です（{len(CSV_COLUMNS)}列必要ですが、{len(row)}列です）。", line_no))
            continue
        rule = ReplaceRule(
            line_no=line_no,
            sheet=row[index["sheet"]],
            cell=row[index["cell"]],
            before=row[index["before"]],
            after=row[index["after"]],
        )
        row_errors = _validate_rule_fields(rule)
        if row_errors:
            errors.extend(row_errors)
        else:
            rules.append(rule)

    if not rules and not errors:
        raise InputError([RuleError("CSVに置換ルールが存在しません。")])
    return rules, errors


def _validate_rule_fields(rule: ReplaceRule) -> list[RuleError]:
    """Excelを参照せずに検証できる項目（必須項目・セル番地）を検証する。"""
    errors: list[RuleError] = []
    missing = [name for name in ("sheet", "cell", "before") if getattr(rule, name) == ""]
    if missing:
        errors.append(RuleError(f"必須項目が未設定です: {', '.join(missing)}", rule.line_no))
    if rule.cell != "" and parse_cell_address(rule.cell) is None:
        errors.append(
            RuleError(
                f"シート「{rule.sheet}」セル番地「{rule.cell}」が不正です（A1形式で指定してください）。",
                rule.line_no,
            )
        )
    return errors


def detect_unsupported_features(input_path: Path) -> list[str]:
    """openpyxl で保持できず、出力ファイルで失われる可能性が高い要素を検出する。"""
    found: list[str] = []
    try:
        with zipfile.ZipFile(input_path) as archive:
            names = archive.namelist()
            for name_re, content_re, description in UNSUPPORTED_FEATURES:
                if description in found:
                    continue
                for name in names:
                    if name_re.match(name) and (content_re is None or content_re.search(archive.read(name))):
                        found.append(description)
                        break
    except (OSError, zipfile.BadZipFile):
        return []  # 読み込みエラーは load_excel で報告する
    return found


def load_excel(input_path: Path) -> Workbook:
    """Excelファイルを読み込む。.xlsm の場合はマクロを保持する。"""
    keep_vba = input_path.suffix.lower() == ".xlsm"
    try:
        return load_workbook(input_path, keep_vba=keep_vba, rich_text=True)
    except Exception as e:  # openpyxl は破損ファイルで様々な例外を送出するためまとめて扱う
        raise ExecutionError(f"Excelファイルの読み込みに失敗しました: {input_path}（{type(e).__name__}: {e}）") from e


def _get_target_cell(ws: Worksheet, rule: ReplaceRule) -> Cell | MergedCell:
    row, column = parse_cell_address(rule.cell)  # type: ignore[misc]  # 検証済み
    return ws.cell(row=row, column=column)


def _check_cell_type(cell: Cell | MergedCell, rule: ReplaceRule) -> str | None:
    """セルが置換対象（文字列セル）かを判定し、対象外の場合はエラー内容を返す。"""
    if isinstance(cell, MergedCell):
        return "は結合セルの左上以外のセルのため置換できません。"
    if cell.data_type == "f":
        return "は数式セルのため置換対象外です。"
    if cell.value is None:
        return "は空のため置換できません（文字列型ではありません）。"
    if cell.data_type == "e":
        return "はエラー値のため置換できません（文字列型ではありません）。"
    if cell.data_type != "s" or not isinstance(cell.value, (str, CellRichText)):
        return f"は文字列型ではありません（{_describe_type(cell.value)}）。"
    return None


def _describe_type(value: object) -> str:
    if isinstance(value, bool):
        return "真偽値"
    if isinstance(value, (int, float)):
        return "数値"
    if hasattr(value, "isoformat"):
        return "日付・時刻"
    return type(value).__name__


def replace_text(value: CellText, before: str, after: str) -> CellText:
    """セル値内の before をすべて after に置換する。

    リッチテキストの場合、すべての出現箇所が1つの書式ブロック内に収まっていれば
    ブロック単位で置換して書式を維持する。書式ブロックをまたぐ場合はプレーン文字列にする。
    """
    if isinstance(value, str):
        return value.replace(before, after)

    expected = str(value).replace(before, after)
    blocks: list[str | TextBlock] = []
    for block in value:
        if isinstance(block, TextBlock):
            blocks.append(TextBlock(block.font, block.text.replace(before, after)))
        else:
            blocks.append(block.replace(before, after))
    replaced = CellRichText([b for b in blocks if (b.text if isinstance(b, TextBlock) else b)])
    if str(replaced) == expected:
        return replaced
    return expected


def plan_replacements(
    wb: Workbook, rules: Sequence[ReplaceRule]
) -> tuple[dict[tuple[str, str], CellText], list[ReplaceRule], list[RuleError]]:
    """ワークブックを変更せずに全ルールを検証し、置換後のセル値を計算する。

    同一セルへの複数ルールは、CSVの記載順に置換した結果に対して次のルールを検証する。

    Returns:
        (セルごとの置換後の値, 成功したルール, エラー)
    """
    new_values: dict[tuple[str, str], CellText] = {}
    succeeded: list[ReplaceRule] = []
    errors: list[RuleError] = []
    sheetnames = set(wb.sheetnames)

    for rule in rules:
        if rule.sheet not in sheetnames:
            errors.append(RuleError(f"シート「{rule.sheet}」が存在しません。", rule.line_no))
            continue
        ws = wb[rule.sheet]
        location = f"シート「{rule.sheet}」セル「{rule.cell}」"
        if not isinstance(ws, Worksheet):
            errors.append(RuleError(f"シート「{rule.sheet}」はグラフシートのため置換できません。", rule.line_no))
            continue

        cell = _get_target_cell(ws, rule)
        key = (rule.sheet, cell.coordinate)
        if key in new_values:
            current = new_values[key]
        else:
            type_error = _check_cell_type(cell, rule)
            if type_error:
                errors.append(RuleError(location + type_error, rule.line_no))
                continue
            current = cell.value

        if rule.before not in str(current):
            errors.append(RuleError(f"{location}に置換前文字列「{rule.before}」が存在しません。", rule.line_no))
            continue

        replaced = replace_text(current, rule.before, rule.after)
        text = str(replaced)
        if ILLEGAL_CHARACTERS_RE.search(text):
            message = f"{location}の置換後文字列にExcelで使用できない制御文字が含まれています。"
            errors.append(RuleError(message, rule.line_no))
            continue
        if len(text) > MAX_CELL_TEXT_LENGTH:
            errors.append(
                RuleError(
                    f"{location}の置換後文字列がExcelの上限（{MAX_CELL_TEXT_LENGTH}文字）を超えます。",
                    rule.line_no,
                )
            )
            continue

        new_values[key] = replaced
        succeeded.append(rule)

    return new_values, succeeded, errors


def apply_replacements(wb: Workbook, new_values: dict[tuple[str, str], CellText]) -> None:
    """計算済みの置換後の値をワークブックに反映する。"""
    for (sheet, coordinate), value in new_values.items():
        cell = wb[sheet][coordinate]
        cell.value = value
        # 「=」で始まる文字列やエラー値と同じ文字列が数式・エラー値として扱われないようにする
        cell.data_type = "s"


def save_workbook_safely(wb: Workbook, output_path: Path) -> None:
    """一時ファイルに保存してから出力先へ移動し、失敗時に不完全なファイルを残さない。"""
    fd, tmp_name = tempfile.mkstemp(prefix=f".{output_path.stem}.", suffix=".tmp", dir=output_path.parent)
    os.close(fd)
    tmp_path = Path(tmp_name)
    try:
        wb.save(tmp_path)
        _apply_default_permissions(tmp_path)
        if output_path.exists():
            raise ExecutionError(f"出力ファイルが既に存在します: {output_path}")
        os.replace(tmp_path, output_path)
    except ExecutionError:
        raise
    except Exception as e:
        raise ExecutionError(f"Excelファイルの保存に失敗しました: {output_path}（{type(e).__name__}: {e}）") from e
    finally:
        if tmp_path.exists():
            tmp_path.unlink()


def _apply_default_permissions(path: Path) -> None:
    """mkstemp は所有者のみ読み書き可能なファイルを作るため、通常のファイル作成と同じ権限に戻す。"""
    umask = os.umask(0)
    os.umask(umask)
    os.chmod(path, 0o666 & ~umask)


def print_errors(errors: Sequence[RuleError]) -> None:
    lines = ["ERROR: 置換処理に失敗しました。", ""]
    lines += [error.format() for error in errors]
    lines += ["", f"エラー件数: {len(errors)}", "出力ファイルは作成されませんでした。"]
    print("\n".join(lines), file=sys.stderr)


def print_runtime_error(message: str) -> None:
    print(f"ERROR: {message}\n出力ファイルは作成されませんでした。", file=sys.stderr)


def run(input_path: Path, csv_path: Path, output_path: Path) -> int:
    """置換処理を実行し、終了コードを返す。"""
    errors = validate_paths(input_path, csv_path, output_path)

    rules: list[ReplaceRule] = []
    if csv_path.is_file():
        try:
            rules, rule_errors = read_rules(csv_path)
            errors += rule_errors
        except InputError as e:
            errors += e.errors

    # パス・CSVの形式エラーがある場合はExcelを読み込まずに終了する
    if any(error.line_no is None for error in errors) or not rules:
        print_errors(errors)
        return EXIT_INPUT_ERROR

    wb = load_excel(input_path)
    new_values, succeeded, plan_errors = plan_replacements(wb, rules)
    errors += plan_errors
    if errors:
        print_errors(sorted(errors, key=lambda error: error.line_no or 0))
        return EXIT_INPUT_ERROR

    apply_replacements(wb, new_values)
    save_workbook_safely(wb, output_path)

    unsupported = detect_unsupported_features(input_path)
    if unsupported:
        print("警告: openpyxlの制約により、以下の要素は出力Excelで失われている可能性があります。")
        for description in unsupported:
            print(f"  - {description}")
        print()

    print("置換成功:")
    for rule in succeeded:
        print(f"  [{rule.line_no}行目] {rule.sheet}!{rule.cell.upper()}")
    print()
    print(f"置換件数: {len(succeeded)}")
    print("処理結果: SUCCESS")
    return EXIT_SUCCESS


def main(argv: Sequence[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    input_path = Path(args.input)
    csv_path = Path(args.csv)
    output_path = resolve_output_path(input_path, csv_path, args.output)

    print("Excel文字列置換処理を開始します。")
    print()
    print(f"入力Excel : {input_path}")
    print(f"置換CSV   : {csv_path}")
    print(f"出力Excel : {output_path}")
    print()
    sys.stdout.flush()

    try:
        return run(input_path, csv_path, output_path)
    except ExecutionError as e:
        print_runtime_error(str(e))
        return EXIT_RUNTIME_ERROR


def _configure_console() -> None:
    """Windowsのコンソール等で表示できない文字があっても異常終了しないようにする。"""
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(errors="replace")


if __name__ == "__main__":
    _configure_console()
    sys.exit(main())
