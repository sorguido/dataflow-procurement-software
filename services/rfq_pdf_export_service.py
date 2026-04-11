"""Servizio export RFQ in PDF A4 multipagina."""

from datetime import datetime
from html import escape
from typing import Dict, List, Optional, Tuple

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import Image, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

from database_manager import DatabaseManager
from utils.format_utils import format_quantity_display
from utils.i18n_utils import normalize_rfq_type, tr, translate_rfq_type


def _safe_text(value) -> str:
    if value is None:
        return ""
    return str(value).strip()


def _fmt_date(db_date: Optional[str]) -> str:
    if not db_date:
        return "-"
    try:
        return datetime.strptime(db_date, "%Y-%m-%d").strftime("%d/%m/%Y")
    except Exception:
        return db_date


def _build_logo(logo_path: Optional[str], max_width: float, max_height: float) -> Optional[Image]:
    if not logo_path:
        return None

    try:
        logo = Image(logo_path)
        ratio = min(max_width / float(logo.imageWidth), max_height / float(logo.imageHeight), 1.0)
        logo.drawWidth = float(logo.imageWidth) * ratio
        logo.drawHeight = float(logo.imageHeight) * ratio
        logo.hAlign = "LEFT"
        return logo
    except Exception:
        return None


def _build_paragraph_styles() -> Dict[str, ParagraphStyle]:
    styles = getSampleStyleSheet()
    return {
        "title": ParagraphStyle(
            "RfqPdfTitle",
            parent=styles["Heading1"],
            fontName="Helvetica-Bold",
            fontSize=16,
            leading=19,
            alignment=0,
            textColor=colors.HexColor("#1f2937"),
            spaceAfter=6,
        ),
        "meta": ParagraphStyle(
            "RfqPdfMeta",
            parent=styles["Normal"],
            fontName="Helvetica",
            fontSize=10,
            leading=13,
            alignment=0,
            textColor=colors.HexColor("#111827"),
        ),
        "body": ParagraphStyle(
            "RfqPdfBody",
            parent=styles["Normal"],
            fontName="Helvetica",
            fontSize=10,
            leading=14,
            textColor=colors.HexColor("#111827"),
        ),
        "table_header": ParagraphStyle(
            "RfqPdfTableHeader",
            parent=styles["Normal"],
            fontName="Helvetica-Bold",
            fontSize=9,
            leading=11,
            textColor=colors.white,
            alignment=1,
        ),
        "table_cell": ParagraphStyle(
            "RfqPdfTableCell",
            parent=styles["Normal"],
            fontName="Helvetica",
            fontSize=8.8,
            leading=11,
            textColor=colors.HexColor("#111827"),
        ),
        "table_cell_center": ParagraphStyle(
            "RfqPdfTableCellCenter",
            parent=styles["Normal"],
            fontName="Helvetica",
            fontSize=8.8,
            leading=11,
            textColor=colors.HexColor("#111827"),
            alignment=1,
        ),
    }


def _build_table(
    details_rows: List[Tuple],
    is_conto_lavoro: bool,
    styles: Dict[str, ParagraphStyle],
    usable_width: float,
) -> Table:
    if is_conto_lavoro:
        headers = [
            tr("Item #"),
            tr("Code"),
            tr("Attachment"),
            tr("Description"),
            tr("Qty"),
            tr("Raw Code"),
            tr("Raw Attachment"),
            tr("Material for Processing"),
        ]
        width_weights = [0.055, 0.1, 0.13, 0.28, 0.08, 0.1, 0.13, 0.125]
    else:
        headers = [
            tr("Item #"),
            tr("Code"),
            tr("Attachment"),
            tr("Description"),
            tr("Qty"),
        ]
        width_weights = [0.07, 0.13, 0.17, 0.47, 0.16]

    col_widths = [usable_width * weight for weight in width_weights]

    header_row = [Paragraph(escape(h), styles["table_header"]) for h in headers]
    table_data = [header_row]

    for idx, row in enumerate(details_rows, start=1):
        _, code, attachment, description, qty, raw_code, raw_attachment, material_for_processing = row
        base_cells = [
            Paragraph(escape(str(idx)), styles["table_cell_center"]),
            Paragraph(escape(_safe_text(code)), styles["table_cell"]),
            Paragraph(escape(_safe_text(attachment)), styles["table_cell"]),
            Paragraph(escape(_safe_text(description)), styles["table_cell"]),
            Paragraph(escape(_safe_text(format_quantity_display(qty))), styles["table_cell_center"]),
        ]

        if is_conto_lavoro:
            base_cells.extend(
                [
                    Paragraph(escape(_safe_text(raw_code)), styles["table_cell"]),
                    Paragraph(escape(_safe_text(raw_attachment)), styles["table_cell"]),
                    Paragraph(escape(_safe_text(material_for_processing)), styles["table_cell"]),
                ]
            )

        table_data.append(base_cells)

    table = Table(table_data, colWidths=col_widths, repeatRows=1, hAlign="LEFT")
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1f4e78")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("ALIGN", (0, 0), (-1, 0), "CENTER"),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LINEBELOW", (0, 0), (-1, 0), 0.8, colors.HexColor("#d1d5db")),
                ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#d1d5db")),
                ("LEFTPADDING", (0, 0), (-1, -1), 4),
                ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f8fafc")]),
            ]
        )
    )
    return table


def _load_rfq_dataset(db_path: str, request_id: int, read_only: bool = False) -> dict:
    with DatabaseManager(db_path, read_only=read_only) as db_manager:
        request_data = db_manager.get_richiesta_full_data(request_id)
        if not request_data:
            raise ValueError(tr("Dettagli RdO non trovati."))

        issue_date_db, expiry_date_db, _reference, rfq_type = request_data
        details_rows = db_manager.get_dettagli_by_richiesta(request_id)

    return {
        "request_id": request_id,
        "issue_date": _fmt_date(issue_date_db),
        "expiry_date": _fmt_date(expiry_date_db),
        "rfq_type": normalize_rfq_type(rfq_type),
        "details_rows": details_rows,
    }


def export_rfq_pdf(
    db_path: str,
    request_id: int,
    output_path: str,
    logo_path: Optional[str] = None,
    read_only: bool = False,
) -> dict:
    """Genera il PDF RFQ e ritorna metadata di export (warning inclusi)."""
    dataset = _load_rfq_dataset(db_path=db_path, request_id=request_id, read_only=read_only)
    styles = _build_paragraph_styles()
    warnings: List[str] = []

    doc = SimpleDocTemplate(
        output_path,
        pagesize=A4,
        leftMargin=2 * cm,
        rightMargin=2 * cm,
        topMargin=2 * cm,
        bottomMargin=2 * cm,
        title=tr("Richiesta di Offerta") + f" {request_id}",
    )

    story = []

    logo = _build_logo(logo_path=logo_path, max_width=5.0 * cm, max_height=2.4 * cm)
    if logo_path and logo is None:
        warnings.append(tr("Logo configurato non valido: export eseguito senza logo."))

    title_text = tr("Richiesta di Offerta")
    rfq_type_label = tr("Tipo RdO")
    rfq_type_value = translate_rfq_type(dataset["rfq_type"])
    meta_html = "<br/>".join(
        [
            f"<b>{escape(tr('Numero RdO'))}:</b> {escape(str(dataset['request_id']))}",
            f"<b>{escape(rfq_type_label)}:</b> {escape(rfq_type_value)}",
            f"<b>{escape(tr('Issue Date'))}:</b> {escape(dataset['issue_date'])}",
            f"<b>{escape(tr('Expiry Date'))}:</b> {escape(dataset['expiry_date'])}",
        ]
    )
    if logo is not None:
        story.append(logo)
        story.append(Spacer(1, 0.35 * cm))
    story.append(Paragraph(escape(title_text), styles["title"]))
    story.append(Paragraph(meta_html, styles["meta"]))
    story.append(Spacer(1, 0.5 * cm))

    story.append(
        Paragraph(
            escape(tr("Gentile Fornitore, con la presente sono a richiedere la Vs. migliore quotazione per il seguente materiale:")),
            styles["body"],
        )
    )
    story.append(Spacer(1, 0.35 * cm))

    is_conto_lavoro = dataset["rfq_type"] == "Conto lavoro"
    table = _build_table(
        details_rows=dataset["details_rows"],
        is_conto_lavoro=is_conto_lavoro,
        styles=styles,
        usable_width=doc.width,
    )
    story.append(table)
    story.append(Spacer(1, 0.45 * cm))

    story.append(
        Paragraph(
            escape(tr("In attesa di un Vs. gentile riscontro, porgo cordiali saluti.")),
            styles["body"],
        )
    )

    doc.build(story)
    return {
        "output_path": output_path,
        "warnings": warnings,
        "rfq_type": dataset["rfq_type"],
    }
