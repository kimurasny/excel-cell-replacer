"""excel_replace.py の自動テスト。"""

from __future__ import annotations

import hashlib
import io
import re
import subprocess
import sys
import zipfile
from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.comments import Comment
from openpyxl.drawing.image import Image
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation
from PIL import Image as PILImage

import excel_replace

ROOT = Path(__file__).resolve().parent.parent
VBA_BYTES = b"dummy-vba-project-\x00\x01\x02"


# ---------------------------------------------------------------------------
# テスト用ファイルの作成
# ---------------------------------------------------------------------------


def build_template(path: Path) -> Path:
    """書式・数式・画像などを含むテンプレートExcelを作成する。"""
    wb = Workbook()
    cover = wb.active
    cover.title = "表紙"
    cover["B2"] = "株式会社AAA 御中"
    cover["B2"].font = Font(name="メイリオ", size=14, bold=True, color="FF0000")
    cover["B2"].fill = PatternFill("solid", fgColor="FFFF00")
    thin = Side(style="thin", color="000000")
    cover["B2"].border = Border(left=thin, right=thin, top=thin, bottom=thin)
    cover["B2"].alignment = Alignment(horizontal="center", wrap_text=True)
    cover["B3"] = "2025年度 計画書（2025年度版）"
    cover["B4"] = "AAA-AAA-AAA"
    cover["B5"] = "不要な文字列を削除する"
    cover["B6"] = "大文字ABC"
    cover["C2"] = "置換対象外 株式会社AAA"
    cover["C2"].font = Font(italic=True, color="0000FF")
    cover["D2"] = 12345
    cover["D2"].number_format = "#,##0"
    cover["D3"] = "=D2*2"
    cover["D4"] = "=SUM(D2:D3)"
    cover.merge_cells("E2:G3")
    cover["E2"] = "結合セル AAA"
    cover.row_dimensions[2].height = 40
    cover.column_dimensions["B"].width = 30
    cover.freeze_panes = "A2"
    cover.page_setup.orientation = "landscape"
    cover.page_setup.paperSize = cover.PAPERSIZE_A4
    cover.page_margins.left = 1.5
    cover.print_area = "A1:G20"
    cover.oddHeader.center.text = "ヘッダー"
    cover["C5"].comment = Comment("コメントです", "tester")

    dv = DataValidation(type="list", formula1='"はい,いいえ"', allow_blank=True)
    cover.add_data_validation(dv)
    dv.add("H2")

    img_buf = io.BytesIO()
    PILImage.new("RGB", (20, 10), (255, 0, 0)).save(img_buf, format="PNG")
    img_buf.seek(0)
    cover.add_image(Image(img_buf), "J2")

    settings = wb.create_sheet("設定")
    settings["C5"] = "環境: 開発環境"
    settings["D10"] = "旧システムから旧システムへ移行"
    settings["A1"] = "=表紙!D2"

    wb.save(path)
    return path


def make_xlsm(xlsx_path: Path, xlsm_path: Path) -> Path:
    """xlsx にダミーの vbaProject.bin を追加し、マクロ有効ブック（.xlsm）を作成する。"""
    with zipfile.ZipFile(xlsx_path) as src, zipfile.ZipFile(xlsm_path, "w", zipfile.ZIP_DEFLATED) as dst:
        for item in src.infolist():
            data = src.read(item.filename)
            if item.filename == "[Content_Types].xml":
                text = data.decode("utf-8")
                text = text.replace(
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
                    "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
                )
                text = text.replace(
                    "<Default ",
                    '<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/><Default ',
                    1,
                )
                data = text.encode("utf-8")
            elif item.filename == "xl/_rels/workbook.xml.rels":
                text = data.decode("utf-8").replace(
                    "</Relationships>",
                    '<Relationship Id="rIdVBA" '
                    'Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" '
                    'Target="vbaProject.bin"/></Relationships>',
                )
                data = text.encode("utf-8")
            dst.writestr(item, data)
        dst.writestr("xl/vbaProject.bin", VBA_BYTES)
    return xlsm_path


def write_csv(path: Path, content: str, encoding: str = "utf-8") -> Path:
    path.write_bytes(content.encode(encoding))
    return path


def csv_text(*rows: str) -> str:
    return "\n".join(("sheet,cell,before,after", *rows)) + "\n"


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


@pytest.fixture
def template(tmp_path: Path) -> Path:
    return build_template(tmp_path / "template.xlsx")


def run_cli(*args: str | Path) -> int:
    return excel_replace.main([str(a) for a in args])


# ---------------------------------------------------------------------------
# 正常系
# ---------------------------------------------------------------------------


def test_replace_xlsx(tmp_path: Path, template: Path, capsys: pytest.CaptureFixture[str]) -> None:
    csv_path = write_csv(
        tmp_path / "customer_A.csv",
        csv_text(
            "表紙,B2,株式会社AAA,株式会社BBB",
            "表紙,B3,2025年度,2026年度",
            "設定,C5,開発環境,本番環境",
            "設定,D10,旧システム,新システム",
        ),
    )
    assert run_cli("--input", template, "--csv", csv_path) == 0

    output = tmp_path / "customer_A.xlsx"
    wb = load_workbook(output)
    assert wb["表紙"]["B2"].value == "株式会社BBB 御中"
    assert wb["表紙"]["B3"].value == "2026年度 計画書（2026年度版）"
    assert wb["設定"]["C5"].value == "環境: 本番環境"
    assert wb["設定"]["D10"].value == "新システムから新システムへ移行"

    out = capsys.readouterr().out
    assert "Excel文字列置換処理を開始します。" in out
    assert f"入力Excel : {template}" in out
    assert f"置換CSV   : {csv_path}" in out
    assert f"出力Excel : {output}" in out
    assert "  [2行目] 表紙!B2\n  [3行目] 表紙!B3\n  [4行目] 設定!C5\n  [5行目] 設定!D10\n" in out
    assert "置換件数: 4" in out
    assert "処理結果: SUCCESS" in out


def test_replace_xlsm_keeps_macro(tmp_path: Path, template: Path) -> None:
    xlsm = make_xlsm(template, tmp_path / "macro_template.xlsm")
    csv_path = write_csv(tmp_path / "customer_A.csv", csv_text("表紙,B2,株式会社AAA,株式会社BBB"))

    assert run_cli("--input", xlsm, "--csv", csv_path) == 0

    output = tmp_path / "customer_A.xlsm"
    assert output.exists()
    with zipfile.ZipFile(output) as zf:
        assert zf.read("xl/vbaProject.bin") == VBA_BYTES
        content_types = zf.read("[Content_Types].xml").decode("utf-8")
        assert "application/vnd.ms-excel.sheet.macroEnabled.main+xml" in content_types
        assert "vbaProject" in zf.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    wb = load_workbook(output, keep_vba=True)
    assert wb["表紙"]["B2"].value == "株式会社BBB 御中"


def test_replace_multiple_sheets(tmp_path: Path, template: Path) -> None:
    csv_path = write_csv(tmp_path / "multi.csv", csv_text("設定,C5,開発,検証", "表紙,B6,ABC,XYZ"))
    assert run_cli("--input", template, "--csv", csv_path) == 0
    wb = load_workbook(tmp_path / "multi.xlsx")
    assert wb["設定"]["C5"].value == "環境: 検証環境"
    assert wb["表紙"]["B6"].value == "大文字XYZ"


def test_multiple_rules_same_cell_applied_in_order(tmp_path: Path, template: Path) -> None:
    csv_path = write_csv(
        tmp_path / "same_cell.csv",
        csv_text(
            "表紙,B2,株式会社AAA,株式会社BBB",
            "表紙,B2,BBB 御中,BBB 様",  # 1行目の置換結果に対して検証・置換される
            "表紙,b2,株式会社,(株)",  # 小文字のセル番地も同じセルとして扱う
        ),
    )
    assert run_cli("--input", template, "--csv", csv_path) == 0
    assert load_workbook(tmp_path / "same_cell.xlsx")["表紙"]["B2"].value == "(株)BBB 様"


def test_partial_match_replace(tmp_path: Path, template: Path) -> None:
    csv_path = write_csv(tmp_path / "partial.csv", csv_text("表紙,B2,AAA,CCC"))
    assert run_cli("--input", template, "--csv", csv_path) == 0
    assert load_workbook(tmp_path / "partial.xlsx")["表紙"]["B2"].value == "株式会社CCC 御中"


def test_replace_all_occurrences_in_cell(tmp_path: Path, template: Path) -> None:
    csv_path = write_csv(tmp_path / "all.csv", csv_text("表紙,B4,AAA,Z"))
    assert run_cli("--input", template, "--csv", csv_path) == 0
    assert load_workbook(tmp_path / "all.xlsx")["表紙"]["B4"].value == "Z-Z-Z"


def test_replace_with_empty_string(tmp_path: Path, template: Path) -> None:
    csv_path = write_csv(tmp_path / "delete.csv", csv_text("表紙,B5,不要な,", "表紙,B2, 御中,"))
    assert run_cli("--input", template, "--csv", csv_path) == 0
    ws = load_workbook(tmp_path / "delete.xlsx")["表紙"]
    assert ws["B5"].value == "文字列を削除する"
    assert ws["B2"].value == "株式会社AAA"


def test_replace_japanese_text_with_bom_csv(tmp_path: Path, template: Path) -> None:
    csv_path = write_csv(
        tmp_path / "日本語.csv",
        csv_text('表紙,B3,"2025年度 計画書","令和8年度　事業計画書, 改訂版"'),
        encoding="utf-8-sig",
    )
    assert run_cli("--input", template, "--csv", csv_path) == 0
    value = load_workbook(tmp_path / "日本語.xlsx")["表紙"]["B3"].value
    assert value == "令和8年度　事業計画書, 改訂版（2025年度版）"


def test_whitespace_is_not_trimmed_and_case_sensitive(tmp_path: Path, template: Path) -> None:
    csv_path = write_csv(tmp_path / "space.csv", csv_text("表紙,B2,AAA ,AAA_", "表紙,B6,abc,xyz"))
    assert run_cli("--input", template, "--csv", csv_path) == 1  # 大文字・小文字を区別するため abc は存在しない

    csv_path = write_csv(tmp_path / "space2.csv", csv_text("表紙,B2,AAA ,AAA_"))
    assert run_cli("--input", template, "--csv", csv_path) == 0
    assert load_workbook(tmp_path / "space2.xlsx")["表紙"]["B2"].value == "株式会社AAA_御中"


def test_default_output_uses_csv_name_and_directory(tmp_path: Path, template: Path) -> None:
    csv_dir = tmp_path / "csv_dir"
    csv_dir.mkdir()
    csv_path = write_csv(csv_dir / "customer_A.csv", csv_text("表紙,B2,AAA,BBB"))
    assert run_cli("--input", template, "--csv", csv_path) == 0
    assert (csv_dir / "customer_A.xlsx").exists()
    assert not (tmp_path / "customer_A.xlsx").exists()


def test_output_option(tmp_path: Path, template: Path) -> None:
    csv_path = write_csv(tmp_path / "customer_A.csv", csv_text("表紙,B2,AAA,BBB"))
    out_dir = tmp_path / "out"
    out_dir.mkdir()
    output = out_dir / "納品用.xlsx"
    assert run_cli("--input", template, "--csv", csv_path, "--output", output) == 0
    assert load_workbook(output)["表紙"]["B2"].value == "株式会社BBB 御中"
    assert not (tmp_path / "customer_A.xlsx").exists()


def test_formats_and_other_contents_are_preserved(tmp_path: Path, template: Path) -> None:
    original_hash = sha256(template)
    csv_path = write_csv(tmp_path / "keep.csv", csv_text("表紙,B2,AAA,BBB"))
    assert run_cli("--input", template, "--csv", csv_path) == 0
    assert sha256(template) == original_hash  # 元ファイルは変更しない

    src = load_workbook(template)["表紙"]
    wb = load_workbook(tmp_path / "keep.xlsx")
    ws = wb["表紙"]
    assert wb.sheetnames == ["表紙", "設定"]

    # 置換対象セル自体の書式も維持される（StyleProxy 同士は == で比較できないため repr で比較する）
    for attr in ("font", "fill", "border", "alignment"):
        assert repr(getattr(ws["B2"], attr)) == repr(getattr(src["B2"], attr))
    # 置換対象外のセル・書式
    assert ws["C2"].value == "置換対象外 株式会社AAA"
    assert repr(ws["C2"].font) == repr(src["C2"].font)
    assert ws["D2"].value == 12345
    assert ws["D2"].number_format == "#,##0"
    assert ws["D3"].value == "=D2*2"
    assert ws["D4"].value == "=SUM(D2:D3)"
    assert wb["設定"]["A1"].value == "=表紙!D2"
    assert [str(r) for r in ws.merged_cells.ranges] == ["E2:G3"]
    assert ws.row_dimensions[2].height == 40
    assert ws.column_dimensions["B"].width == 30
    assert ws.freeze_panes == "A2"
    assert ws.page_setup.orientation == "landscape"
    assert int(ws.page_setup.paperSize) == int(ws.PAPERSIZE_A4)
    assert ws.page_margins.left == 1.5
    assert ws.print_area == "'表紙'!$A$1:$G$20"
    assert ws.oddHeader.center.text == "ヘッダー"
    assert ws["C5"].comment.text == "コメントです"
    assert len(ws.data_validations.dataValidation) == 1
    assert len(ws._images) == 1


def test_rich_text_formatting_is_kept(tmp_path: Path) -> None:
    template = tmp_path / "rich.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "表紙"
    red = InlineFont(color="FF0000", b=True)
    ws["A1"] = CellRichText("宛先: ", TextBlock(red, "株式会社AAA"), " 御中")
    ws["A2"] = CellRichText("株式会社", TextBlock(red, "AAA"))
    ws["A3"] = CellRichText("書式あり", TextBlock(red, "保持"))
    wb.save(template)

    # A2 は置換前文字列が書式ブロックをまたぐため、プレーン文字列として置換される
    csv_path = write_csv(tmp_path / "rich_out.csv", csv_text("表紙,A1,AAA,BBB", "表紙,A2,株式会社AAA,株式会社CCC"))
    assert run_cli("--input", template, "--csv", csv_path) == 0

    ws = load_workbook(tmp_path / "rich_out.xlsx", rich_text=True)["表紙"]
    a1 = ws["A1"].value
    assert isinstance(a1, CellRichText)
    assert str(a1) == "宛先: 株式会社BBB 御中"
    assert any(isinstance(b, TextBlock) and b.text == "株式会社BBB" and b.font.b for b in a1)
    assert str(ws["A2"].value) == "株式会社CCC"
    assert isinstance(ws["A3"].value, CellRichText)  # 置換対象外セルのリッチテキストも維持される


def test_result_starting_with_equal_stays_string(tmp_path: Path, template: Path) -> None:
    csv_path = write_csv(tmp_path / "equal.csv", csv_text("表紙,B6,大文字ABC,=1+1", "表紙,B4,AAA-AAA-AAA,#N/A"))
    assert run_cli("--input", template, "--csv", csv_path) == 0
    ws = load_workbook(tmp_path / "equal.xlsx")["表紙"]
    assert ws["B6"].value == "=1+1"
    assert ws["B6"].data_type == "s"
    assert ws["B4"].value == "#N/A"
    assert ws["B4"].data_type == "s"


def test_no_warning_for_supported_contents(tmp_path: Path, template: Path, capsys: pytest.CaptureFixture[str]) -> None:
    csv_path = write_csv(tmp_path / "plain.csv", csv_text("表紙,B2,AAA,BBB"))
    assert run_cli("--input", template, "--csv", csv_path) == 0
    assert "警告" not in capsys.readouterr().out  # 画像・コメント・入力規則などは警告対象外


def test_detect_unsupported_features(tmp_path: Path) -> None:
    path = tmp_path / "unsupported.xlsx"
    with zipfile.ZipFile(path, "w") as zf:
        zf.writestr("xl/drawings/drawing1.xml", "<xdr:wsDr><xdr:twoCellAnchor><xdr:sp><xdr:nvSpPr/></xdr:sp>")
        zf.writestr("xl/worksheets/sheet1.xml", '<worksheet><controls><control r:id="rId1"/></controls>')
        zf.writestr("xl/worksheets/sheet2.xml", "<worksheet><extLst><ext/></extLst></worksheet>")
        zf.writestr("customUI/customUI14.xml", "<customUI/>")
    assert excel_replace.detect_unsupported_features(path) == [
        "図形・テキストボックス",
        "フォームコントロール・ActiveXコントロール",
        "シートの拡張情報（スパークライン、スライサー、新しい形式の条件付き書式・入力規則など）",
        "リボンのカスタマイズ（customUI）",
    ]


def test_cli_subprocess(tmp_path: Path, template: Path) -> None:
    csv_path = write_csv(tmp_path / "customer_A.csv", csv_text("表紙,B2,AAA,BBB"))
    result = subprocess.run(
        [sys.executable, str(ROOT / "excel_replace.py"), "--input", str(template), "--csv", str(csv_path)],
        capture_output=True,
        text=True,
        encoding="utf-8",
        env={"PYTHONIOENCODING": "utf-8", "PATH": ""},
    )
    assert result.returncode == 0, result.stderr
    assert "処理結果: SUCCESS" in result.stdout
    assert (tmp_path / "customer_A.xlsx").exists()


def test_example_csv_is_valid(tmp_path: Path, template: Path) -> None:
    csv_path = tmp_path / "replace.csv"
    csv_path.write_bytes((ROOT / "examples" / "replace.csv").read_bytes())
    assert run_cli("--input", template, "--csv", csv_path) == 0


# ---------------------------------------------------------------------------
# 異常系
# ---------------------------------------------------------------------------


def assert_failed(
    capsys: pytest.CaptureFixture[str], output: Path, *messages: str, error_count: int | None = None
) -> str:
    err = capsys.readouterr().err
    assert err.startswith("ERROR: 置換処理に失敗しました。")
    for message in messages:
        assert message in err
    if error_count is not None:
        assert f"エラー件数: {error_count}" in err
    assert "出力ファイルは作成されませんでした。" in err
    assert not output.exists()
    assert not list(output.parent.glob("*.tmp"))
    return err


def test_sheet_not_found(tmp_path: Path, template: Path, capsys: pytest.CaptureFixture[str]) -> None:
    csv_path = write_csv(tmp_path / "ng.csv", csv_text("存在しない,B2,AAA,BBB"))
    assert run_cli("--input", template, "--csv", csv_path) == 1
    assert_failed(capsys, tmp_path / "ng.xlsx", "[2行目] シート「存在しない」が存在しません。", error_count=1)


def test_sheet_name_must_match_exactly(tmp_path: Path, template: Path, capsys: pytest.CaptureFixture[str]) -> None:
    csv_path = write_csv(tmp_path / "ng.csv", csv_text("表紙 ,B2,AAA,BBB"))
    assert run_cli("--input", template, "--csv", csv_path) == 1
    assert_failed(capsys, tmp_path / "ng.xlsx", "シート「表紙 」が存在しません。")


def test_before_text_not_found(tmp_path: Path, template: Path, capsys: pytest.CaptureFixture[str]) -> None:
    csv_path = write_csv(tmp_path / "ng.csv", csv_text("表紙,B2,株式会社ZZZ,BBB"))
    assert run_cli("--input", template, "--csv", csv_path) == 1
    assert_failed(
        capsys, tmp_path / "ng.xlsx", "[2行目] シート「表紙」セル「B2」に置換前文字列「株式会社ZZZ」が存在しません。"
    )


def test_before_text_removed_by_previous_rule(
    tmp_path: Path, template: Path, capsys: pytest.CaptureFixture[str]
) -> None:
    csv_path = write_csv(tmp_path / "ng.csv", csv_text("表紙,B2,株式会社AAA,株式会社BBB", "表紙,B2,株式会社AAA,X"))
    assert run_cli("--input", template, "--csv", csv_path) == 1
    err = assert_failed(capsys, tmp_path / "ng.xlsx", "[3行目] シート「表紙」セル「B2」に置換前文字列", error_count=1)
    assert "[2行目]" not in err


@pytest.mark.parametrize("address", ["B", "2", "1B", "B0", "XFE1", "A1048577", " B2", "B2:C3", "$B$2", "表紙!B2"])
def test_invalid_cell_address(tmp_path: Path, template: Path, capsys: pytest.CaptureFixture[str], address: str) -> None:
    csv_path = write_csv(tmp_path / "ng.csv", csv_text(f'表紙,"{address}",AAA,BBB'))
    assert run_cli("--input", template, "--csv", csv_path) == 1
    assert_failed(capsys, tmp_path / "ng.xlsx", f"[2行目] シート「表紙」セル番地「{address}」が不正です")


def test_empty_before(tmp_path: Path, template: Path, capsys: pytest.CaptureFixture[str]) -> None:
    csv_path = write_csv(tmp_path / "ng.csv", csv_text("表紙,B2,,BBB"))
    assert run_cli("--input", template, "--csv", csv_path) == 1
    assert_failed(capsys, tmp_path / "ng.xlsx", "[2行目] 必須項目が未設定です: before")


def test_missing_required_fields(tmp_path: Path, template: Path, capsys: pytest.CaptureFixture[str]) -> None:
    csv_path = write_csv(tmp_path / "ng.csv", csv_text(",,,BBB"))
    assert run_cli("--input", template, "--csv", csv_path) == 1
    assert_failed(capsys, tmp_path / "ng.xlsx", "[2行目] 必須項目が未設定です: sheet, cell, before")


def test_numeric_cell(tmp_path: Path, template: Path, capsys: pytest.CaptureFixture[str]) -> None:
    csv_path = write_csv(tmp_path / "ng.csv", csv_text("表紙,D2,123,456"))
    assert run_cli("--input", template, "--csv", csv_path) == 1
    assert_failed(capsys, tmp_path / "ng.xlsx", "[2行目] シート「表紙」セル「D2」は文字列型ではありません（数値）。")


def test_formula_cell(tmp_path: Path, template: Path, capsys: pytest.CaptureFixture[str]) -> None:
    csv_path = write_csv(tmp_path / "ng.csv", csv_text("表紙,D4,SUM,AVERAGE"))
    assert run_cli("--input", template, "--csv", csv_path) == 1
    assert_failed(capsys, tmp_path / "ng.xlsx", "[2行目] シート「表紙」セル「D4」は数式セルのため置換対象外です。")


def test_empty_and_merged_cells(tmp_path: Path, template: Path, capsys: pytest.CaptureFixture[str]) -> None:
    csv_path = write_csv(tmp_path / "ng.csv", csv_text("表紙,Z99,AAA,BBB", "表紙,F2,AAA,BBB"))
    assert run_cli("--input", template, "--csv", csv_path) == 1
    assert_failed(
        capsys,
        tmp_path / "ng.xlsx",
        "[2行目] シート「表紙」セル「Z99」は空のため置換できません",
        "[3行目] シート「表紙」セル「F2」は結合セルの左上以外のセルのため置換できません。",
        error_count=2,
    )


@pytest.mark.parametrize(
    ("content", "message"),
    [
        ("sheet,cell,before\n表紙,B2,AAA\n", "CSVのヘッダーが不正です"),
        ("シート,セル,置換前,置換後\n表紙,B2,AAA,BBB\n", "CSVのヘッダーが不正です"),
        ("sheet,cell,before,after,memo\n表紙,B2,AAA,BBB,x\n", "CSVのヘッダーが不正です"),
        ("", "CSVファイルが空です"),
        ("sheet,cell,before,after\n", "CSVに置換ルールが存在しません。"),
        ("sheet,cell,before,after\n\n,,,\n", "CSVに置換ルールが存在しません。"),
        ("sheet,cell,before,after\n表紙,B2,AAA\n", "[2行目] 列数が不正です（4列必要ですが、3列です）。"),
        ('sheet,cell,before,after\n表紙,B2,"AA"A,BBB\n', "[2行目] CSVの形式が不正です"),
    ],
)
def test_invalid_csv(
    tmp_path: Path, template: Path, capsys: pytest.CaptureFixture[str], content: str, message: str
) -> None:
    csv_path = write_csv(tmp_path / "ng.csv", content)
    assert run_cli("--input", template, "--csv", csv_path) == 1
    assert_failed(capsys, tmp_path / "ng.xlsx", message)


def test_csv_not_utf8(tmp_path: Path, template: Path, capsys: pytest.CaptureFixture[str]) -> None:
    csv_path = write_csv(tmp_path / "ng.csv", csv_text("表紙,B2,AAA,BBB"), encoding="cp932")
    assert run_cli("--input", template, "--csv", csv_path) == 1
    assert_failed(capsys, tmp_path / "ng.xlsx", "CSVファイルをUTF-8として読み込めません")


def test_header_order_can_differ(tmp_path: Path, template: Path) -> None:
    csv_path = write_csv(tmp_path / "order.csv", "before,after,sheet,cell\nAAA,BBB,表紙,B2\n")
    assert run_cli("--input", template, "--csv", csv_path) == 0
    assert load_workbook(tmp_path / "order.xlsx")["表紙"]["B2"].value == "株式会社BBB 御中"


def test_output_file_already_exists(tmp_path: Path, template: Path, capsys: pytest.CaptureFixture[str]) -> None:
    csv_path = write_csv(tmp_path / "dup.csv", csv_text("表紙,B2,AAA,BBB"))
    output = tmp_path / "dup.xlsx"
    output.write_bytes(b"existing")
    assert run_cli("--input", template, "--csv", csv_path) == 1
    err = capsys.readouterr().err
    assert f"出力ファイルが既に存在します: {output}" in err
    assert output.read_bytes() == b"existing"


def test_output_same_as_input_is_rejected(tmp_path: Path, template: Path) -> None:
    original_hash = sha256(template)
    csv_path = write_csv(tmp_path / "self.csv", csv_text("表紙,B2,AAA,BBB"))
    assert run_cli("--input", template, "--csv", csv_path, "--output", template) == 1
    assert sha256(template) == original_hash


def test_multiple_errors_reported_together(tmp_path: Path, template: Path, capsys: pytest.CaptureFixture[str]) -> None:
    csv_path = write_csv(
        tmp_path / "multi_ng.csv",
        csv_text(
            "表紙,B2,株式会社AAA,株式会社BBB",  # 正常
            "設定x,C5,開発環境,本番環境",
            "表紙,B10,株式会社AAA,株式会社BBB",
            "表紙,ZZZZ1,AAA,BBB",
            "表紙,D3,D2,D5",
            "設定,C5,,本番",
        ),
    )
    assert run_cli("--input", template, "--csv", csv_path) == 1
    err = assert_failed(
        capsys,
        tmp_path / "multi_ng.xlsx",
        "[3行目] シート「設定x」が存在しません。",
        "[4行目] シート「表紙」セル「B10」は空のため置換できません",
        "[5行目] シート「表紙」セル番地「ZZZZ1」が不正です",
        "[6行目] シート「表紙」セル「D3」は数式セルのため置換対象外です。",
        "[7行目] 必須項目が未設定です: before",
        error_count=5,
    )
    # CSVの行番号順に表示される
    line_numbers = [int(n) for n in re.findall(r"^\[(\d+)行目\]", err, flags=re.MULTILINE)]
    assert line_numbers == [3, 4, 5, 6, 7]


def test_error_message_format(tmp_path: Path, template: Path, capsys: pytest.CaptureFixture[str]) -> None:
    csv_path = write_csv(
        tmp_path / "fmt.csv",
        csv_text("表紙,B2,AAA,BBB", "設定,C5,存在しない文字,X", "表紙,B3,株式会社AAA,株式会社BBB"),
    )
    assert run_cli("--input", template, "--csv", csv_path) == 1
    assert capsys.readouterr().err == (
        "ERROR: 置換処理に失敗しました。\n"
        "\n"
        "[3行目] シート「設定」セル「C5」に置換前文字列「存在しない文字」が存在しません。\n"
        "[4行目] シート「表紙」セル「B3」に置換前文字列「株式会社AAA」が存在しません。\n"
        "\n"
        "エラー件数: 2\n"
        "出力ファイルは作成されませんでした。\n"
    )


def test_input_and_csv_not_found(tmp_path: Path, capsys: pytest.CaptureFixture[str]) -> None:
    missing_input = tmp_path / "missing.xlsx"
    missing_csv = tmp_path / "missing.csv"
    assert run_cli("--input", missing_input, "--csv", missing_csv) == 1
    assert_failed(
        capsys,
        tmp_path / "missing.xlsx",
        f"入力Excelファイルが存在しません: {missing_input}",
        f"CSVファイルが存在しません: {missing_csv}",
        error_count=2,
    )


@pytest.mark.parametrize("name", ["template.xls", "template.csv", "template"])
def test_unsupported_extension(tmp_path: Path, capsys: pytest.CaptureFixture[str], name: str) -> None:
    input_path = tmp_path / name
    input_path.write_bytes(b"dummy")
    csv_path = write_csv(tmp_path / "ng.csv", csv_text("表紙,B2,AAA,BBB"))
    assert run_cli("--input", input_path, "--csv", csv_path) == 1
    assert "は対象外です" in capsys.readouterr().err


def test_output_extension_must_match_input(tmp_path: Path, template: Path, capsys: pytest.CaptureFixture[str]) -> None:
    csv_path = write_csv(tmp_path / "ng.csv", csv_text("表紙,B2,AAA,BBB"))
    output = tmp_path / "out.xlsm"
    assert run_cli("--input", template, "--csv", csv_path, "--output", output) == 1
    assert_failed(capsys, output, "出力Excelファイルの拡張子は入力Excelと同じ「.xlsx」にしてください")


def test_output_directory_not_found(tmp_path: Path, template: Path, capsys: pytest.CaptureFixture[str]) -> None:
    csv_path = write_csv(tmp_path / "ng.csv", csv_text("表紙,B2,AAA,BBB"))
    output = tmp_path / "no_dir" / "out.xlsx"
    assert run_cli("--input", template, "--csv", csv_path, "--output", output) == 1
    assert "出力先ディレクトリが存在しません" in capsys.readouterr().err


def test_broken_excel_returns_exit_code_2(tmp_path: Path, capsys: pytest.CaptureFixture[str]) -> None:
    broken = tmp_path / "broken.xlsx"
    broken.write_bytes(b"this is not a zip file")
    csv_path = write_csv(tmp_path / "ng.csv", csv_text("表紙,B2,AAA,BBB"))
    assert run_cli("--input", broken, "--csv", csv_path) == 2
    err = capsys.readouterr().err
    assert "ERROR: Excelファイルの読み込みに失敗しました" in err
    assert not (tmp_path / "ng.xlsx").exists()


def test_save_failure_leaves_no_file(
    tmp_path: Path, template: Path, capsys: pytest.CaptureFixture[str], monkeypatch: pytest.MonkeyPatch
) -> None:
    def broken_save(self: Workbook, filename: str) -> None:
        Path(filename).write_bytes(b"partial")
        raise OSError("ディスクがいっぱいです")

    monkeypatch.setattr(Workbook, "save", broken_save)
    csv_path = write_csv(tmp_path / "save_ng.csv", csv_text("表紙,B2,AAA,BBB"))
    assert run_cli("--input", template, "--csv", csv_path) == 2
    err = capsys.readouterr().err
    assert "ERROR: Excelファイルの保存に失敗しました" in err
    assert "処理結果: SUCCESS" not in capsys.readouterr().out
    assert not (tmp_path / "save_ng.xlsx").exists()
    assert not list(tmp_path.glob("*.tmp"))


def test_missing_arguments_returns_exit_code_1(capsys: pytest.CaptureFixture[str]) -> None:
    with pytest.raises(SystemExit) as exc:
        excel_replace.main(["--input", "template.xlsx"])
    assert exc.value.code == 1
    assert "コマンドライン引数が不正です" in capsys.readouterr().err
