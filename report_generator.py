"""Python 3.8 Word report generator using COM automation.

This module mirrors the MATLAB script in the repository and keeps external
dependencies to a minimum by relying solely on ``pywin32`` (``win32com``).
It can build a Word report with a cover page, section headings, paragraphs,
bullet lists, tables, figures, and placeholder replacement.
"""
from __future__ import annotations

import os
import tempfile
from typing import Any, Dict, Iterable, List, Optional
import base64
import html
import importlib.util


# Word constant values (avoids importing win32com constants module)
WD_ALIGN_PARAGRAPH_LEFT = 0
WD_ALIGN_PARAGRAPH_CENTER = 1
WD_LINE_SPACE_MULTIPLE = 5
WD_PAGE_BREAK = 7
WD_SECTION_BREAK_CONTINUOUS = 3
WD_HEADER_FOOTER_PRIMARY = 1
WD_ALIGN_PAGE_NUMBER_CENTER = 1
WD_AUTO_FIT_CONTENT = 2
WD_COLLAPSE_END = 0
TABLE_BORDER_STYLE = "1px solid #999"


def _get_dispatch():
    spec = importlib.util.find_spec("win32com.client")
    if spec is None:
        raise RuntimeError(
            "pywin32 (win32com) is required to control Microsoft Word via COM"
        )

    from win32com.client import Dispatch  # type: ignore

    return Dispatch


def generate_word_report(
    output_path: str,
    report_title: str,
    sections: Iterable[Dict[str, Any]],
    options: Optional[Dict[str, Any]] = None,
) -> None:
    """Generate a Word report with COM automation.

    Parameters mirror the MATLAB implementation:
    * ``output_path``: destination ``.doc``/``.docx`` file path.
    * ``report_title``: main title on the cover page.
    * ``sections``: iterable of dictionaries describing content blocks.
    * ``options``: optional dictionary controlling layout, metadata, and
      placeholder replacement.
    """

    options = _merge_options(options)

    Dispatch = _get_dispatch()
    word = Dispatch("Word.Application")
    word.Visible = False
    doc = None
    try:
        doc = _create_document(word, options)
        _set_document_properties(doc, options)
        selection = word.Selection
        _configure_page_setup(word, options)

        _add_cover_page(selection, report_title, options)
        _add_body_content(selection, list(sections), options)
        _add_footer_and_numbers(doc, options)
        _replace_placeholders(doc, options)

        doc.SaveAs(os.path.abspath(output_path))
    finally:
        if doc is not None:
            doc.Close()
        word.Quit()


# ---------------------------------------------------------------------------
# Document helpers
# ---------------------------------------------------------------------------

def _merge_options(options: Optional[Dict[str, Any]]) -> Dict[str, Any]:
    opts: Dict[str, Any] = options.copy() if options else {}

    opts.setdefault("AddPageNums", True)
    opts.setdefault("HeadingFont", {"Name": "Arial", "Size": 16})
    opts.setdefault("BodyFont", {"Name": "Arial", "Size": 11})
    opts.setdefault("LineSpacing", 1.15)
    opts.setdefault("SpaceBefore", 6)
    opts.setdefault("SpaceAfter", 6)
    opts.setdefault("Margins", {})
    opts.setdefault("TableStyle", "Table Grid")

    margins = opts["Margins"]
    margins.setdefault("Top", 72)
    margins.setdefault("Bottom", 72)
    margins.setdefault("Left", 72)
    margins.setdefault("Right", 72)
    return opts


def _create_document(word: Any, options: Dict[str, Any]) -> Any:
    template = options.get("Template")
    docs = word.Documents
    if template and os.path.exists(template):
        return docs.Open(template)
    return docs.Add()


def _set_document_properties(doc: Any, options: Dict[str, Any]) -> None:
    if options.get("Author"):
        doc.SetProperty("Author", options["Author"])
    if options.get("Company"):
        doc.SetProperty("Company", options["Company"])


def _configure_page_setup(word: Any, options: Dict[str, Any]) -> None:
    doc = word.ActiveDocument
    page_setup = doc.PageSetup
    margins = options["Margins"]
    page_setup.TopMargin = margins["Top"]
    page_setup.BottomMargin = margins["Bottom"]
    page_setup.LeftMargin = margins["Left"]
    page_setup.RightMargin = margins["Right"]


def _apply_paragraph_formatting(selection: Any, options: Dict[str, Any], alignment: Optional[int] = None) -> None:
    para_format = selection.ParagraphFormat
    if alignment is not None:
        para_format.Alignment = alignment
    para_format.LineSpacingRule = WD_LINE_SPACE_MULTIPLE
    para_format.LineSpacing = options["LineSpacing"] * 12
    para_format.SpaceBefore = options["SpaceBefore"]
    para_format.SpaceAfter = options["SpaceAfter"]


def _add_cover_page(selection: Any, report_title: str, options: Dict[str, Any]) -> None:
    selection.WholeStory()
    selection.Delete()

    _apply_paragraph_formatting(selection, options, WD_ALIGN_PARAGRAPH_CENTER)
    selection.Font.Name = options["HeadingFont"]["Name"]
    selection.Font.Size = options["HeadingFont"]["Size"]
    selection.Font.Bold = True
    selection.TypeText(report_title)
    selection.TypeParagraph()

    selection.Font.Bold = False
    selection.Font.Name = options["BodyFont"]["Name"]
    selection.Font.Size = options["BodyFont"]["Size"]
    if options.get("Author"):
        selection.TypeText(f"Author: {options['Author']}")
        selection.TypeParagraph()
    if options.get("Company"):
        selection.TypeText(f"Company: {options['Company']}")
        selection.TypeParagraph()
    selection.InsertBreak(WD_PAGE_BREAK)


def _add_body_content(selection: Any, sections: List[Dict[str, Any]], options: Dict[str, Any]) -> None:
    for section in sections:
        title = section.get("Title")
        if title:
            _apply_paragraph_formatting(selection, options, WD_ALIGN_PARAGRAPH_LEFT)
            selection.Font.Name = options["HeadingFont"]["Name"]
            selection.Font.Size = options["HeadingFont"]["Size"]
            selection.Font.Bold = True
            selection.TypeText(str(title))
            selection.TypeParagraph()

        selection.Font.Bold = False
        selection.Font.Name = options["BodyFont"]["Name"]
        selection.Font.Size = options["BodyFont"]["Size"]
        _apply_paragraph_formatting(selection, options, WD_ALIGN_PARAGRAPH_LEFT)

        for paragraph in section.get("Paragraphs", []) or []:
            selection.TypeText(str(paragraph))
            selection.TypeParagraph()

        for bullet in section.get("Bullets", []) or []:
            selection.TypeText(f"• {bullet}")
            selection.TypeParagraph()
        if section.get("Bullets"):
            selection.TypeParagraph()

        for table in section.get("Tables", []) or []:
            _add_table(selection, table, options)
            selection.TypeParagraph()

        figures = section.get("Figures")
        if figures:
            _add_figures(selection, figures)

        selection.InsertBreak(WD_SECTION_BREAK_CONTINUOUS)


def _add_table(selection: Any, table_def: Dict[str, Any], options: Dict[str, Any]) -> None:
    rows_data = table_def.get("Rows")
    if not rows_data:
        return

    rows = len(rows_data)
    cols = len(rows_data[0]) if rows_data else 0
    header = table_def.get("Header") or []
    if header:
        rows += 1

    word_table = selection.Tables.Add(selection.Range, rows, cols)
    word_table.Style = options["TableStyle"]
    word_table.Range.Font.Name = options["BodyFont"]["Name"]
    word_table.Range.Font.Size = options["BodyFont"]["Size"]

    current_row = 1
    if header:
        for col, value in enumerate(header, start=1):
            cell = word_table.Cell(current_row, col)
            cell.Range.Text = str(value)
            cell.Range.Font.Bold = True
        current_row += 1

    for row_values in rows_data:
        for col, value in enumerate(row_values, start=1):
            cell = word_table.Cell(current_row, col)
            cell.Range.Text = str(value)
        current_row += 1

    word_table.AutoFitBehavior(WD_AUTO_FIT_CONTENT)
    range_after = word_table.Range
    range_after.Collapse(WD_COLLAPSE_END)
    range_after.Select()


def _add_figures(selection: Any, figures: Iterable[Dict[str, Any]]) -> None:
    normalized, temp_files = _normalize_figures(figures)
    try:
        row_indices = [fig.get("RowIndex", 1) or 1 for fig in normalized]
        for row in sorted(set(row_indices)):
            row_figures = [fig for fig, idx in zip(normalized, row_indices) if idx == row]
            _add_figure_row(selection, row_figures)
    finally:
        _delete_temp_files(temp_files)


def _add_figure_row(selection: Any, figure_row: List[Dict[str, Any]]) -> None:
    if not figure_row:
        return

    table = selection.Tables.Add(selection.Range, 1, len(figure_row))
    table.Borders.Enable = False

    for col, fig in enumerate(figure_row, start=1):
        cell = table.Cell(1, col)
        cell_range = cell.Range
        inline_shapes = cell_range.InlineShapes
        inline_shapes.AddPicture(fig["Path"], False, True)

        cell_range.Collapse(WD_COLLAPSE_END)
        caption_text = f" {fig['Caption']}" if fig.get("Caption") else ""
        selection = cell_range.Document.Application.Selection
        cell_range.Select()
        selection.InsertCaption("Figure", caption_text)

    range_after = table.Range
    range_after.Collapse(WD_COLLAPSE_END)
    range_after.Select()
    selection.TypeParagraph()


def _add_footer_and_numbers(doc: Any, options: Dict[str, Any]) -> None:
    sections = doc.Sections
    for index in range(1, sections.Count + 1):
        section = sections.Item(index)
        primary_footer = section.Footers.Item(WD_HEADER_FOOTER_PRIMARY)
        rng = primary_footer.Range
        footer_text = options.get("FooterText")
        if footer_text:
            rng.Text = footer_text
        if options.get("AddPageNums", True):
            primary_footer.PageNumbers.Add(WD_ALIGN_PAGE_NUMBER_CENTER)


def _replace_placeholders(doc: Any, options: Dict[str, Any]) -> None:
    placeholders = options.get("Placeholders")
    if not placeholders:
        return

    for name, payload in placeholders.items():
        token = f"{{{{{name}}}}}"
        search_range = doc.Content
        find_obj = search_range.Find
        find_obj.Forward = True
        find_obj.Format = False

        while find_obj.Execute(token, False, False, False, False, False, True, 1, False, "", False):
            if isinstance(payload, str):
                search_range.Text = payload
            elif isinstance(payload, dict):
                if payload.get("Rows"):
                    _add_table_at_range(search_range, payload, options)
                elif payload.get("Path") or len(payload) > 1:
                    _add_figures_at_range(search_range, payload)
                else:
                    search_range.Text = ""
            else:
                search_range.Text = ""

            start = search_range.End
            doc_content = doc.Content
            search_range = doc.Range(Start=start, End=doc_content.End)
            find_obj = search_range.Find
            find_obj.Forward = True
            find_obj.Format = False


def _add_table_at_range(range_obj: Any, table_def: Dict[str, Any], options: Dict[str, Any]) -> None:
    range_obj.Text = ""
    range_obj.Collapse(WD_COLLAPSE_END)

    rows_data = table_def.get("Rows")
    if not rows_data:
        return

    rows = len(rows_data)
    cols = len(rows_data[0]) if rows_data else 0
    header = table_def.get("Header") or []
    if header:
        rows += 1

    word_table = range_obj.Tables.Add(range_obj, rows, cols)
    word_table.Style = options["TableStyle"]
    word_table.Range.Font.Name = options["BodyFont"]["Name"]
    word_table.Range.Font.Size = options["BodyFont"]["Size"]

    current_row = 1
    if header:
        for col, value in enumerate(header, start=1):
            cell = word_table.Cell(current_row, col)
            cell.Range.Text = str(value)
            cell.Range.Font.Bold = True
        current_row += 1

    for row_values in rows_data:
        for col, value in enumerate(row_values, start=1):
            cell = word_table.Cell(current_row, col)
            cell.Range.Text = str(value)
        current_row += 1

    word_table.AutoFitBehavior(WD_AUTO_FIT_CONTENT)
    range_after = word_table.Range
    range_after.Collapse(WD_COLLAPSE_END)
    range_after.Select()


def _add_figures_at_range(range_obj: Any, figures: Iterable[Dict[str, Any]]) -> None:
    normalized, temp_files = _normalize_figures(figures)
    try:
        range_obj.Text = ""
        range_obj.Collapse(WD_COLLAPSE_END)

        row_indices = [fig.get("RowIndex", 1) or 1 for fig in normalized]
        for row in sorted(set(row_indices)):
            row_figures = [fig for fig, idx in zip(normalized, row_indices) if idx == row]
            word_table = range_obj.Tables.Add(range_obj, 1, len(row_figures))
            word_table.Borders.Enable = False

            for col, fig in enumerate(row_figures, start=1):
                cell = word_table.Cell(1, col)
                cell_range = cell.Range
                inline_shapes = cell_range.InlineShapes
                inline_shapes.AddPicture(fig["Path"], False, True)

                cell_range.Collapse(WD_COLLAPSE_END)
                caption_text = f" {fig['Caption']}" if fig.get("Caption") else ""
                selection = cell_range.Document.Application.Selection
                cell_range.Select()
                selection.InsertCaption("Figure", caption_text)

            range_obj = word_table.Range
            range_obj.Collapse(WD_COLLAPSE_END)
    finally:
        _delete_temp_files(temp_files)


def _normalize_figures(figures: Iterable[Dict[str, Any]]) -> (List[Dict[str, Any]], List[str]):
    normalized: List[Dict[str, Any]] = []
    temp_files: List[str] = []
    for figure in figures:
        normalized_fig, temp_files = _ensure_figure_path(dict(figure), temp_files)
        normalized.append(normalized_fig)
    return normalized, temp_files


def _ensure_figure_path(fig: Dict[str, Any], temp_files: List[str]) -> (Dict[str, Any], List[str]):
    if fig.get("Path") and os.path.isfile(str(fig["Path"])):
        fig["Path"] = os.path.abspath(str(fig["Path"]))
        return fig, temp_files

    # Support matplotlib figure objects if present without importing matplotlib globally.
    candidate = fig.get("FigureHandle") or fig.get("Path")
    if candidate is not None and _looks_like_matplotlib(candidate):
        temp_path = os.path.abspath(tempfile.mktemp(suffix=".png"))
        candidate.savefig(temp_path)
        fig["Path"] = temp_path
        temp_files.append(temp_path)
        return fig, temp_files

    raise ValueError("Figure entries must provide an existing file path or matplotlib figure handle")


def _looks_like_matplotlib(obj: Any) -> bool:
    return hasattr(obj, "savefig")


def _delete_temp_files(temp_files: List[str]) -> None:
    for path in temp_files:
        try:
            os.remove(path)
        except FileNotFoundError:
            pass


def generate_html_report(
    output_path: str,
    report_title: str,
    sections: Iterable[Dict[str, Any]],
    options: Optional[Dict[str, Any]] = None,
) -> None:
    """Generate an HTML report using the same schema as ``generate_word_report``.

    The function mirrors the section/option structure used for the Word export
    but produces a UTF-8 encoded ``.html`` file. Images are normalized to
    absolute paths or embedded as base64 data URIs when ``EmbedImages`` is set
    (globally or per-figure with ``Embed``).
    """

    options = _merge_options(options)
    sections = list(sections)

    html_parts: List[str] = [
        "<!DOCTYPE html>",
        "<html lang=\"zh-CN\">",
        "<head>",
        "<meta http-equiv=\"X-UA-Compatible\" content=\"IE=edge\">",
        "<meta charset=\"UTF-8\">",
        f"<title>{html.escape(report_title)}</title>",
        "</head>",
    ]

    root_style = _build_root_style(options)
    html_parts.append(f"<body style=\"{root_style}\">")

    html_parts.append(_render_cover_page(report_title, options))
    for section in sections:
        html_parts.append(_render_section(section, options))

    final_html = _replace_placeholders_html("".join(html_parts) + "</body></html>", options)

    with open(output_path, "w", encoding="utf-8") as fp:
        fp.write(final_html)


def _build_root_style(options: Dict[str, Any]) -> str:
    body_font = options["BodyFont"]
    styles = [
        f"font-family:{body_font['Name']}, sans-serif",
        f"font-size:{body_font['Size']}pt",
        f"line-height:{options['LineSpacing']}",
        "margin:16px",
        "-ms-text-size-adjust:100%",
    ]
    return "; ".join(styles)


def _render_cover_page(report_title: str, options: Dict[str, Any]) -> str:
    heading_font = options["HeadingFont"]
    body_font = options["BodyFont"]
    parts = ["<section style=\"text-align:center; page-break-after:always;\">"]
    parts.append(
        f"<h1 style=\"font-family:{heading_font['Name']}, sans-serif; font-size:{heading_font['Size']}pt; margin:24px 0;\">{html.escape(report_title)}"  # noqa: E501
        "</h1>"
    )

    if options.get("Author"):
        parts.append(
            f"<p style=\"font-family:{body_font['Name']}, sans-serif; font-size:{body_font['Size']}pt;\">"
            f"作者：{html.escape(str(options['Author']))}</p>"
        )
    if options.get("Company"):
        parts.append(
            f"<p style=\"font-family:{body_font['Name']}, sans-serif; font-size:{body_font['Size']}pt;\">"
            f"单位：{html.escape(str(options['Company']))}</p>"
        )

    parts.append("</section>")
    return "".join(parts)


def _render_section(section: Dict[str, Any], options: Dict[str, Any]) -> str:
    heading_font = options["HeadingFont"]
    body_font = options["BodyFont"]
    html_chunks: List[str] = ["<section style=\"margin-bottom:24px;\">"]

    title = section.get("Title")
    if title:
        html_chunks.append(
            f"<h2 style=\"font-family:{heading_font['Name']}, sans-serif; font-size:{heading_font['Size']}pt; margin:12px 0;\">{html.escape(str(title))}</h2>"
        )

    for paragraph in section.get("Paragraphs", []) or []:
        text = _escape_or_placeholder(paragraph, options)
        html_chunks.append(
            f"<p style=\"font-family:{body_font['Name']}, sans-serif; font-size:{body_font['Size']}pt; margin:8px 0;\">{text}</p>"
        )

    bullets = section.get("Bullets", []) or []
    if bullets:
        html_chunks.append(
            f"<ul style=\"font-family:{body_font['Name']}, sans-serif; font-size:{body_font['Size']}pt; margin:8px 0 16px 20px;\">"
        )
        for bullet in bullets:
            html_chunks.append(f"<li>{_escape_or_placeholder(bullet, options)}</li>")
        html_chunks.append("</ul>")

    for table in section.get("Tables", []) or []:
        html_chunks.append(_render_table(table, options))

    figures = section.get("Figures")
    if figures:
        html_chunks.append(_render_figures(figures, options))

    html_chunks.append("</section>")
    return "".join(html_chunks)


def _render_table(table_def: Dict[str, Any], options: Dict[str, Any]) -> str:
    rows_data = table_def.get("Rows")
    if not rows_data:
        return ""

    body_font = options["BodyFont"]
    header = table_def.get("Header") or []
    cell_style = (
        f"border: {TABLE_BORDER_STYLE}; padding:6px; "
        f"font-family:{body_font['Name']}, sans-serif; font-size:{body_font['Size']}pt;"
    )
    table_style = f"border-collapse:collapse; width:100%; margin:12px 0; border: {TABLE_BORDER_STYLE};"

    html_table: List[str] = [f"<table style=\"{table_style}\">"]
    if header:
        html_table.append("<thead><tr>")
        for value in header:
            html_table.append(f"<th style=\"{cell_style}; font-weight:bold;\">{html.escape(str(value))}</th>")
        html_table.append("</tr></thead>")

    html_table.append("<tbody>")
    for row_values in rows_data:
        html_table.append("<tr>")
        for value in row_values:
            html_table.append(f"<td style=\"{cell_style}\">{html.escape(str(value))}</td>")
        html_table.append("</tr>")
    html_table.append("</tbody></table>")
    return "".join(html_table)


def _render_figures(figures: Iterable[Dict[str, Any]], options: Dict[str, Any]) -> str:
    normalized, temp_files = _normalize_figures(figures)
    parts: List[str] = []
    try:
        row_indices = [fig.get("RowIndex", 1) or 1 for fig in normalized]
        for row in sorted(set(row_indices)):
            row_figures = [fig for fig, idx in zip(normalized, row_indices) if idx == row]
            parts.append(_render_figure_row(row_figures, options))
    finally:
        _delete_temp_files(temp_files)
    return "".join(parts)


def _render_figure_row(figure_row: List[Dict[str, Any]], options: Dict[str, Any]) -> str:
    if not figure_row:
        return ""

    cells: List[str] = ["<tr>"]
    for fig in figure_row:
        src = _figure_src(fig, options)
        caption = html.escape(str(fig.get("Caption", ""))) if fig.get("Caption") else ""
        img_tag = f"<img src=\"{src}\" alt=\"{caption}\" style=\"max-width:100%; height:auto; display:block; margin:auto;\">"
        caption_tag = f"<div style=\"text-align:center; margin-top:6px;\">{caption}</div>" if caption else ""
        cell_style = f"padding:8px; text-align:center; border:{TABLE_BORDER_STYLE};"
        cells.append(f"<td style=\"{cell_style}\">{img_tag}{caption_tag}</td>")
    cells.append("</tr>")

    table_style = f"border-collapse:collapse; width:100%; margin:12px 0; border: {TABLE_BORDER_STYLE};"
    return f"<table style=\"{table_style}\"><tbody>{''.join(cells)}</tbody></table>"


def _figure_src(fig: Dict[str, Any], options: Dict[str, Any]) -> str:
    embed = fig.get("Embed", options.get("EmbedImages", False))
    path = fig.get("Path")
    if not path:
        return ""
    if embed:
        with open(path, "rb") as fp:
            encoded = base64.b64encode(fp.read()).decode("ascii")
        mime = "image/png" if path.lower().endswith(".png") else "image/jpeg"
        return f"data:{mime};base64,{encoded}"
    return path


def _escape_or_placeholder(value: Any, options: Dict[str, Any]) -> str:
    text = str(value)
    placeholders = options.get("Placeholders") or {}
    if text.startswith("{{") and text.endswith("}}") and text[2:-2] in placeholders:
        return _render_placeholder(text[2:-2], placeholders[text[2:-2]], options)
    return html.escape(text)


def _render_placeholder(name: str, payload: Any, options: Dict[str, Any]) -> str:
    if isinstance(payload, str):
        return html.escape(payload)
    if isinstance(payload, dict):
        if payload.get("Rows"):
            return _render_table(payload, options)
        if payload.get("Path") or len(payload) > 1:
            return _render_figures([payload], options)
    return ""


def _replace_placeholders_html(content: str, options: Dict[str, Any]) -> str:
    placeholders = options.get("Placeholders") or {}
    for name, payload in placeholders.items():
        token = f"{{{{{name}}}}}"
        rendered = _render_placeholder(name, payload, options)
        content = content.replace(token, rendered)
    return content


__all__ = ["generate_word_report", "generate_html_report"]
