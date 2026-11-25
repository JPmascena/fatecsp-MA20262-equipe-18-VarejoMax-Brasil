"""Gera VarejoMax_05_APIs_Web.xlsx via Python stdlib (OpenXML) a partir dos CSVs coletados.

Abas geradas:
1. Resumo_APIs (sumario executivo e metricas reais das 3 APIs)
2. Cambio_PTAX (608 linhas)
3. Feriados_BrasilAPI (22 linhas)
4. Clima_24_Cidades (14.592 linhas)
"""

from __future__ import annotations

import csv
import datetime as dt
import statistics
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

BASE_DIR = Path(__file__).resolve().parents[1]
DADOS_DIR = BASE_DIR / "dados_api"
OUT_XLSX = BASE_DIR / "VarejoMax_05_APIs_Web.xlsx"


def col_letter(idx: int) -> str:
    """Converte indice 1-based para letra de coluna Excel (1 -> A, 27 -> AA)."""
    res = []
    while idx > 0:
        idx, rem = divmod(idx - 1, 26)
        res.append(chr(65 + rem))
    return "".join(reversed(res))


def clean_xml_text(text: str) -> str:
    """Remove caracteres de controle invalidos em XML 1.0 e escapa entidades."""
    cleaned = "".join(ch for ch in text if ch in ("\t", "\n", "\r") or ord(ch) >= 0x20)
    return escape(cleaned)


class SharedStrings:
    """Tabela de strings compartilhadas para compactar o arquivo OpenXML."""

    def __init__(self) -> None:
        self._map: dict[str, int] = {}
        self._items: list[str] = []
        self.total_count = 0

    def add(self, text: str) -> int:
        self.total_count += 1
        idx = self._map.get(text)
        if idx is None:
            idx = len(self._items)
            self._map[text] = idx
            self._items.append(text)
        return idx

    def to_xml(self) -> str:
        parts = [
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
            f'<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="{self.total_count}" uniqueCount="{len(self._items)}">',
        ]
        for s in self._items:
            preserve = ' xml:space="preserve"' if s.startswith(" ") or s.endswith(" ") else ""
            parts.append(f"<si><t{preserve}>{clean_xml_text(s)}</t></si>")
        parts.append("</sst>")
        return "".join(parts)


def build_sheet_xml(
    rows: list[list[tuple[str, object, int]]],
    col_widths: list[float],
    sst: SharedStrings,
    freeze_row: int = 1,
    autofilter_end_col: int | None = None,
) -> str:
    """Constroi XML de worksheet a partir de linhas tipadas [(tipo, valor, style_idx), ...].

    tipo: 's' (string), 'n' (number)
    """
    max_row = len(rows)
    max_col = max((len(r) for r in rows), default=1)
    dim_ref = f"A1:{col_letter(max_col)}{max_row}"

    parts = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        f'<dimension ref="{dim_ref}"/>',
        "<sheetViews>",
        '<sheetView workbookViewId="0">',
    ]
    if freeze_row > 0:
        parts.append(
            f'<pane ySplit="{freeze_row}" topLeftCell="A{freeze_row + 1}" activePane="bottomLeft" state="frozen"/>'
        )
    parts.extend(["</sheetView>", "</sheetViews>", '<sheetFormatPr defaultRowHeight="18"/>', "<cols>"])
    for i, w in enumerate(col_widths, start=1):
        parts.append(f'<col min="{i}" max="{i}" width="{w:.2f}" customWidth="1"/>')
    parts.append("</cols><sheetData>")

    for r_idx, row in enumerate(rows, start=1):
        row_attrs = f'r="{r_idx}"'
        if r_idx == 1 and freeze_row == 1:
            row_attrs += ' ht="22" customHeight="1"'
        parts.append(f"<row {row_attrs}>")
        for c_idx, (ctype, val, style_idx) in enumerate(row, start=1):
            ref = f"{col_letter(c_idx)}{r_idx}"
            if ctype == "s":
                s_idx = sst.add(str(val))
                parts.append(f'<c r="{ref}" t="s" s="{style_idx}"><v>{s_idx}</v></c>')
            else:
                parts.append(f'<c r="{ref}" s="{style_idx}"><v>{val}</v></c>')
        parts.append("</row>")

    parts.append("</sheetData>")
    if autofilter_end_col is not None and max_row > 1:
        parts.append(f'<autoFilter ref="A1:{col_letter(autofilter_end_col)}{max_row}"/>')
    parts.append("</worksheet>")
    return "".join(parts)


def styles_xml() -> str:
    """Estilos OpenXML:
    0 = normal
    1 = titulo negrito 13pt
    2 = cabecalho tabela (fundo azul marinho #1F4E78, texto branco negrito)
    3 = numero 4 casas (0.0000)
    4 = numero 1 casa (0.0)
    5 = inteiro (#,##0)
    6 = texto negrito
    """
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="3">
    <numFmt numFmtId="164" formatCode="0.0000"/>
    <numFmt numFmtId="165" formatCode="0.0"/>
    <numFmt numFmtId="166" formatCode="#,##0"/>
  </numFmts>
  <fonts count="4">
    <font><sz val="10"/><name val="Calibri"/></font>
    <font><b/><sz val="13"/><color rgb="FF1F4E78"/><name val="Calibri"/></font>
    <font><b/><sz val="10"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>
    <font><b/><sz val="10"/><name val="Calibri"/></font>
  </fonts>
  <fills count="3">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF1F4E78"/></patternFill></fill>
  </fills>
  <borders count="2">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border>
      <left style="thin"><color rgb="FFD9D9D9"/></left>
      <right style="thin"><color rgb="FFD9D9D9"/></right>
      <top style="thin"><color rgb="FFD9D9D9"/></top>
      <bottom style="thin"><color rgb="FFD9D9D9"/></bottom>
    </border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="7">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/>
    <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>
    <xf numFmtId="0" fontId="2" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>
    <xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/>
    <xf numFmtId="165" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/>
    <xf numFmtId="166" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/>
    <xf numFmtId="0" fontId="3" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1"/>
  </cellXfs>
  <cellStyles count="1">
    <cellStyle name="Normal" xfId="0" builtinId="0"/>
  </cellStyles>
</styleSheet>"""


def load_csv(path: Path) -> tuple[list[str], list[list[str]]]:
    with path.open("r", encoding="utf-8", newline="") as f:
        reader = csv.reader(f, delimiter=";")
        header = next(reader)
        data = list(reader)
    return header, data


def build_workbook() -> None:
    ptax_hdr, ptax_rows_raw = load_csv(DADOS_DIR / "cambio_ptax_2025_2026.csv")
    fer_hdr, fer_rows_raw = load_csv(DADOS_DIR / "feriados_nacionais_2025_2026.csv")
    clima_hdr, clima_rows_raw = load_csv(DADOS_DIR / "clima_24_cidades_2025_2026.csv")

    sst = SharedStrings()

    # Estatisticas para a aba Resumo_APIs
    vendas = [float(r[2]) for r in ptax_rows_raw]
    compras = [float(r[1]) for r in ptax_rows_raw]
    oficiais = sum(1 for r in ptax_rows_raw if r[3] == "Oficial BCB")
    carry = sum(1 for r in ptax_rows_raw if r[3] == "Carry-forward")

    tmaxs = [float(r[2]) for r in clima_rows_raw]
    tmins = [float(r[3]) for r in clima_rows_raw]
    precs = [float(r[4]) for r in clima_rows_raw]

    resumo_rows: list[list[tuple[str, object, int]]] = [
        [("s", "VarejoMax Brasil — Consolidação de APIs Públicas REST (2025-01-01 a 2026-08-31)", 1)],
        [("s", f"Gerado em: {dt.datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | Padrão decimal: ponto (.) | Encoding: UTF-8", 0)],
        [],
        [
            ("s", "API / Fonte", 2),
            ("s", "Endpoint Base", 2),
            ("s", "Aba / CSV Destino", 2),
            ("s", "Qtd. Registros", 2),
            ("s", "Métricas Reais Coletadas", 2),
        ],
        [
            ("s", "1. Banco Central do Brasil (Olinda PTAX)", 6),
            ("s", "https://olinda.bcb.gov.br/olinda/servico/PTAX/versao/v1/odata/CotacaoDolarPeriodo", 0),
            ("s", "Cambio_PTAX / cambio_ptax_2025_2026.csv", 0),
            ("n", len(ptax_rows_raw), 5),
            (
                "s",
                f"Oficial BCB: {oficiais} | Carry-forward: {carry} | Venda min: {min(vendas):.4f} / max: {max(vendas):.4f} / media: {statistics.mean(vendas):.4f} | Compra media: {statistics.mean(compras):.4f}",
                0,
            ),
        ],
        [
            ("s", "2. BrasilAPI Feriados Nacionais", 6),
            ("s", "https://brasilapi.com.br/api/feriados/v1/{ano}", 0),
            ("s", "Feriados_BrasilAPI / feriados_nacionais_2025_2026.csv", 0),
            ("n", len(fer_rows_raw), 5),
            ("s", f"Total no recorte 2025-01-01 a 2026-08-31: {len(fer_rows_raw)} feriados nacionais (2025: 14 | 2026 jan-ago: 8)", 0),
        ],
        [
            ("s", "3. Open-Meteo Archive API", 6),
            ("s", "https://archive-api.open-meteo.com/v1/archive", 0),
            ("s", "Clima_24_Cidades / clima_24_cidades_2025_2026.csv", 0),
            ("n", len(clima_rows_raw), 5),
            (
                "s",
                f"24 cidades x 608 dias | Temp Max: {min(tmaxs):.1f} a {max(tmaxs):.1f} C (media {statistics.mean(tmaxs):.1f}) | Temp Min: {min(tmins):.1f} a {max(tmins):.1f} C | Precip max: {max(precs):.1f} mm (media {statistics.mean(precs):.1f} mm)",
                0,
            ),
        ],
    ]

    # Aba Cambio_PTAX
    ptax_sheet_rows: list[list[tuple[str, object, int]]] = [
        [("s", h, 2) for h in ptax_hdr]
    ]
    for r in ptax_rows_raw:
        ptax_sheet_rows.append(
            [
                ("s", r[0], 0),
                ("n", float(r[1]), 3),
                ("n", float(r[2]), 3),
                ("s", r[3], 0),
                ("s", r[4], 0),
            ]
        )

    # Aba Feriados_BrasilAPI
    fer_sheet_rows: list[list[tuple[str, object, int]]] = [
        [("s", h, 2) for h in fer_hdr]
    ]
    for r in fer_rows_raw:
        fer_sheet_rows.append([("s", r[0], 0), ("s", r[1], 0), ("s", r[2], 0)])

    # Aba Clima_24_Cidades
    clima_sheet_rows: list[list[tuple[str, object, int]]] = [
        [("s", h, 2) for h in clima_hdr]
    ]
    for r in clima_rows_raw:
        clima_sheet_rows.append(
            [
                ("s", r[0], 0),
                ("s", r[1], 0),
                ("n", float(r[2]), 4),
                ("n", float(r[3]), 4),
                ("n", float(r[4]), 4),
                ("s", r[5], 0),
            ]
        )

    sheet1_xml = build_sheet_xml(resumo_rows, [38, 62, 42, 16, 85], sst, freeze_row=0, autofilter_end_col=None)
    sheet2_xml = build_sheet_xml(ptax_sheet_rows, [14, 16, 16, 18, 24], sst, freeze_row=1, autofilter_end_col=5)
    sheet3_xml = build_sheet_xml(fer_sheet_rows, [14, 38, 14], sst, freeze_row=1, autofilter_end_col=3)
    sheet4_xml = build_sheet_xml(clima_sheet_rows, [20, 14, 12, 12, 14, 28], sst, freeze_row=1, autofilter_end_col=6)

    content_types = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet3.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet4.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
</Types>"""

    root_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"""

    workbook_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Resumo_APIs" sheetId="1" r:id="rId1"/>
    <sheet name="Cambio_PTAX" sheetId="2" r:id="rId2"/>
    <sheet name="Feriados_BrasilAPI" sheetId="3" r:id="rId3"/>
    <sheet name="Clima_24_Cidades" sheetId="4" r:id="rId4"/>
  </sheets>
</workbook>"""

    workbook_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet3.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet4.xml"/>
  <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>"""

    with zipfile.ZipFile(OUT_XLSX, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types.encode("utf-8"))
        zf.writestr("_rels/.rels", root_rels.encode("utf-8"))
        zf.writestr("xl/workbook.xml", workbook_xml.encode("utf-8"))
        zf.writestr("xl/_rels/workbook.xml.rels", workbook_rels.encode("utf-8"))
        zf.writestr("xl/styles.xml", styles_xml().encode("utf-8"))
        zf.writestr("xl/sharedStrings.xml", sst.to_xml().encode("utf-8"))
        zf.writestr("xl/worksheets/sheet1.xml", sheet1_xml.encode("utf-8"))
        zf.writestr("xl/worksheets/sheet2.xml", sheet2_xml.encode("utf-8"))
        zf.writestr("xl/worksheets/sheet3.xml", sheet3_xml.encode("utf-8"))
        zf.writestr("xl/worksheets/sheet4.xml", sheet4_xml.encode("utf-8"))

    # Validacao de integridade ZIP e XML
    with zipfile.ZipFile(OUT_XLSX, "r") as zf:
        bad = zf.testzip()
        if bad is not None:
            raise RuntimeError(f"Arquivo corrompido no ZIP: {bad}")
        for name in zf.namelist():
            if name.endswith(".xml") or name.endswith(".rels"):
                ET.fromstring(zf.read(name))

    print(f"Workbook gerado e validado com sucesso: {OUT_XLSX} ({OUT_XLSX.stat().st_size:,} bytes)")


if __name__ == "__main__":
    build_workbook()
