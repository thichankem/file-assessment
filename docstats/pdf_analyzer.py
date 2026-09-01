"""PDF analysis: detect tables, cells, merged cells, nested tables and tables
continued across pages.

A PDF does not store table structure the way Word does - it is only strokes and
glyphs, so tables are recovered from the page geometry (PyMuPDF
``find_tables``) and from there:

* **merged cells** - compare each cell against the line grid: a cell wider than
  one column is a horizontal merge, taller than one row a vertical one.
* **nested tables** - a table fully contained in another table's box.
* **split tables** - the bottom fragment of page N and the top fragment of page
  N+1 are joined into ONE table when they have the same column count and the
  same vertical rule positions.
* **repeated header** - whether the first row of the continuation matches the
  first row of the previous fragment.
"""

from __future__ import annotations

import io
import os
import re
from contextlib import redirect_stdout
from dataclasses import dataclass, field

from .analyzer import (
    RE_FIG_CAPTION,
    RE_TBL_CAPTION,
    DocResult,
    FigureInfo,
    TableInfo,
    _finalize,
)

# --- thresholds for detecting a table continued on the next page ----------
#: the end of a table must sit within this share of the page bottom
BOTTOM_ZONE = 0.14
#: the continuation must start within this share of the page top
TOP_ZONE = 0.20
#: tolerance when comparing vertical rule positions (points)
COL_TOL = 3.0
#: tolerance when clustering coordinates into a row/column grid
GRID_TOL = 2.0

#: ignore images / drawings smaller than this (points)
MIN_FIG_SIDE = 14
MIN_FIG_AREA = 900

#: 1 point = 1/72 inch
PT_TO_CM = 2.54 / 72


@dataclass
class _Part:
    """One fragment of a table on a single page."""

    page: int
    bbox: tuple
    rows: int
    cols: int
    cells: int
    merged_h: int
    merged_v: int
    col_bounds: list
    first_row: str
    preview: str
    parent: int | None = None  # index of the parent _Part (nested table)
    used: bool = False  # already chained into another table
    children: list = field(default_factory=list)


# ---------------------------------------------------------------------------
# Geometry helpers
# ---------------------------------------------------------------------------


def _cluster(values: list[float], tol: float = GRID_TOL) -> list[float]:
    """Merge coordinates that are close together into a single grid line."""
    out: list[float] = []
    for v in sorted(values):
        if not out or v - out[-1] > tol:
            out.append(v)
        else:
            out[-1] = (out[-1] + v) / 2
    return out


def _span(lo: float, hi: float, bounds: list[float], tol: float = GRID_TOL) -> int:
    """How many slots of the ``bounds`` grid the span [lo, hi] covers."""
    if len(bounds) < 2:
        return 1
    i0 = min(range(len(bounds)), key=lambda i: abs(bounds[i] - lo))
    i1 = min(range(len(bounds)), key=lambda i: abs(bounds[i] - hi))
    return max(1, i1 - i0)


def _contains(outer: tuple, inner: tuple, tol: float = 2.0) -> bool:
    return (
        outer[0] - tol <= inner[0]
        and outer[1] - tol <= inner[1]
        and outer[2] + tol >= inner[2]
        and outer[3] + tol >= inner[3]
        and (outer[2] - outer[0]) * (outer[3] - outer[1])
        > (inner[2] - inner[0]) * (inner[3] - inner[1])
    )


def _overlap_frac(a: tuple, b: tuple) -> float:
    """Share of ``a`` that is covered by ``b``."""
    w = min(a[2], b[2]) - max(a[0], b[0])
    h = min(a[3], b[3]) - max(a[1], b[1])
    if w <= 0 or h <= 0:
        return 0.0
    area = (a[2] - a[0]) * (a[3] - a[1])
    return (w * h) / area if area else 0.0


# ---------------------------------------------------------------------------
# Reading the tables of one page
# ---------------------------------------------------------------------------


def _find_tables(page, borderless: bool):
    """Call find_tables, swallowing the pymupdf_layout hint printed to stdout."""
    buf = io.StringIO()
    with redirect_stdout(buf):
        try:
            tables = list(page.find_tables().tables)
        except Exception:
            tables = []
        if borderless and not tables:
            try:
                tables = list(page.find_tables(strategy="text").tables)
            except Exception:
                tables = []
    return tables


def _part_from_table(page_no: int, table) -> _Part:
    boxes = [tuple(c) for c in table.cells if c]
    if not boxes:
        boxes = [tuple(table.bbox)]
    col_bounds = _cluster([b[0] for b in boxes] + [b[2] for b in boxes])
    row_bounds = _cluster([b[1] for b in boxes] + [b[3] for b in boxes])

    merged_h = merged_v = 0
    for b in boxes:
        if _span(b[0], b[2], col_bounds) > 1:
            merged_h += 1
        if _span(b[1], b[3], row_bounds) > 1:
            merged_v += 1

    try:
        data = table.extract()
    except Exception:
        data = []
    first_row = ""
    preview = ""
    if data:
        cells = [" ".join(str(c or "").split()) for c in data[0]]
        first_row = "|".join(cells)
        preview = " | ".join(c[:24] for c in cells if c)[:160]

    return _Part(
        page=page_no,
        bbox=tuple(table.bbox),
        rows=table.row_count,
        cols=table.col_count,
        cells=len(boxes),
        merged_h=merged_h,
        merged_v=merged_v,
        col_bounds=col_bounds,
        first_row=first_row,
        preview=preview,
    )


def _link_nested(parts: list[_Part]) -> None:
    """Link tables that sit inside another table on the same page."""
    by_page: dict[int, list[int]] = {}
    for i, p in enumerate(parts):
        by_page.setdefault(p.page, []).append(i)
    for indexes in by_page.values():
        for i in indexes:
            best = None
            best_area = None
            for j in indexes:
                if i == j or not _contains(parts[j].bbox, parts[i].bbox):
                    continue
                area = (parts[j].bbox[2] - parts[j].bbox[0]) * (
                    parts[j].bbox[3] - parts[j].bbox[1]
                )
                if best_area is None or area < best_area:
                    best, best_area = j, area
            parts[i].parent = best
        for i in indexes:
            if parts[i].parent is not None:
                parts[parts[i].parent].children.append(i)


def _cols_match(a: _Part, b: _Part) -> bool:
    if a.cols != b.cols or len(a.col_bounds) != len(b.col_bounds):
        return False
    return all(abs(x - y) <= COL_TOL for x, y in zip(a.col_bounds, b.col_bounds))


def _chain_pages(parts: list[_Part], page_heights: dict[int, float]) -> list[list[int]]:
    """Chain top-level fragments at the bottom of page N to the top of page N+1."""
    tops = [i for i, p in enumerate(parts) if p.parent is None]
    by_page: dict[int, list[int]] = {}
    for i in tops:
        by_page.setdefault(parts[i].page, []).append(i)

    chains: list[list[int]] = []
    for i in tops:
        if parts[i].used:
            continue
        chain = [i]
        parts[i].used = True
        current = i
        while True:
            cur = parts[current]
            height = page_heights.get(cur.page, 0)
            if not height or cur.bbox[3] < height * (1 - BOTTOM_ZONE):
                break  # does not end near the page bottom -> table is complete
            nxt_height = page_heights.get(cur.page + 1)
            if not nxt_height:
                break
            follower = None
            for j in by_page.get(cur.page + 1, []):
                cand = parts[j]
                if cand.used or cand.bbox[1] > nxt_height * TOP_ZONE:
                    continue
                if _cols_match(cur, cand):
                    follower = j
                    break
            if follower is None:
                break
            parts[follower].used = True
            chain.append(follower)
            current = follower
        chains.append(chain)
    return chains


# ---------------------------------------------------------------------------
# Figures
# ---------------------------------------------------------------------------


def _figures_of_page(
    page, table_boxes: list[tuple], images: list[dict], text_len: int
) -> list[tuple[str, tuple]]:
    found: list[tuple[str, tuple]] = []
    page_area = page.rect.get_area() or 1.0

    for info in images:
        box = tuple(info.get("bbox", (0, 0, 0, 0)))
        if box[2] - box[0] < MIN_FIG_SIDE or box[3] - box[1] < MIN_FIG_SIDE:
            continue
        area = (box[2] - box[0]) * (box[3] - box[1])
        if area > 0.85 * page_area and text_len < 15:
            continue  # full-page scan: that is the page itself, not a figure
        found.append(("Image", box))

    try:
        clusters = page.cluster_drawings()
    except Exception:
        clusters = []
    for cluster in clusters:
        box = tuple(cluster)
        w, h = box[2] - box[0], box[3] - box[1]
        if w < MIN_FIG_SIDE or h < MIN_FIG_SIDE or w * h < MIN_FIG_AREA:
            continue
        if w * h > 0.92 * page_area:
            continue  # a border around the whole page
        if any(_overlap_frac(box, t) > 0.6 for t in table_boxes):
            continue  # table ruling lines, not a figure
        if any(_overlap_frac(box, other) > 0.6 for _k, other in found):
            continue  # a frame around an image already counted
        found.append(("Vector drawing", box))

    return found


# ---------------------------------------------------------------------------
# Whole-file analysis
# ---------------------------------------------------------------------------


def analyze_pdf(path: str, threshold: int, borderless: bool = False) -> DocResult:
    import pymupdf

    res = DocResult(path=path, name=os.path.basename(path), file_type="PDF")
    try:
        res.size = os.path.getsize(path)
    except OSError:
        pass

    doc = pymupdf.open(path)
    try:
        if doc.needs_pass:
            res.error = "The PDF is password protected and cannot be read."
            return res

        res.pages = doc.page_count
        parts: list[_Part] = []
        page_heights: dict[int, float] = {}
        page_tables: dict[int, list[tuple]] = {}
        figures: list[FigureInfo] = []

        for pno in range(doc.page_count):
            page = doc[pno]
            page_no = pno + 1
            page_heights[page_no] = page.rect.height

            tables = _find_tables(page, borderless)
            boxes = [tuple(t.bbox) for t in tables]
            page_tables[page_no] = boxes
            for table in tables:
                parts.append(_part_from_table(page_no, table))

            text = page.get_text().strip()
            try:
                images = page.get_image_info() or []
            except Exception:
                images = []

            page_area = page.rect.get_area() or 1.0
            for kind, box in _figures_of_page(page, boxes, images, len(text)):
                w_pt, h_pt = box[2] - box[0], box[3] - box[1]
                figures.append(
                    FigureInfo(
                        file=path,
                        file_name=res.name,
                        index=len(figures) + 1,
                        kind=kind,
                        page=page_no,
                        in_table=any(_overlap_frac(box, t) > 0.6 for t in boxes),
                        width_cm=w_pt * PT_TO_CM,
                        height_cm=h_pt * PT_TO_CM,
                        area_ratio=(w_pt * h_pt) / page_area,
                    )
                )
            for line in text.splitlines():
                line = line.strip()
                if RE_FIG_CAPTION.match(line):
                    res.fig_captions += 1
                elif RE_TBL_CAPTION.match(line):
                    res.tbl_captions += 1

            if len(text) < 15 and any(
                _overlap_frac(tuple(page.rect), tuple(i.get("bbox", (0, 0, 0, 0)))) > 0.6
                for i in images
            ):
                res.scanned_pages += 1

        res.figures = figures
        _link_nested(parts)
        chains = _chain_pages(parts, page_heights)
        res.tables = _build_tables(path, res.name, parts, chains, threshold)

        if res.scanned_pages:
            res.warnings.append(
                f"{res.scanned_pages}/{res.pages} pages are scanned images - "
                "no tables can be detected on them (OCR needed)."
            )
        if not res.tables and res.pages and not borderless:
            res.warnings.append(
                "No tables detected. If the tables have no ruling lines, enable "
                "the 'PDF: detect borderless tables' option."
            )
    finally:
        doc.close()
    return res


def _build_tables(
    path: str,
    name: str,
    parts: list[_Part],
    chains: list[list[int]],
    threshold: int,
) -> list[TableInfo]:
    out: list[TableInfo] = []
    chains.sort(key=lambda c: (parts[c[0]].page, parts[c[0]].bbox[1], parts[c[0]].bbox[0]))

    for index, chain in enumerate(chains, start=1):
        head = parts[chain[0]]
        label = f"Table {index}"
        info = TableInfo(
            file=path,
            file_name=name,
            index=index,
            label=label,
            level=1,
            page_start=head.page,
            page_end=parts[chain[-1]].page,
            grid_cols=head.cols,
            preview=head.preview,
        )
        repeat_header = len(chain) > 1
        rows = cells = 0
        for pos, part_i in enumerate(chain):
            part = parts[part_i]
            part_rows, part_cells = part.rows, part.cells
            if pos > 0:
                same_header = bool(part.first_row) and part.first_row == head.first_row
                repeat_header = repeat_header and same_header
                if same_header:
                    # a repeated header row is not new data
                    part_rows -= 1
                    part_cells -= min(part.cols, part.cells)
                info.split_rows.append(rows)
            rows += part_rows
            cells += part_cells
            info.merged_h += part.merged_h
            info.merged_v += part.merged_v

        info.rows = rows
        info.cells = cells
        info.pages_spanned = info.page_end - info.page_start + 1
        info.is_split = len(chain) > 1
        info.repeat_header = repeat_header and info.is_split
        _finalize(info, threshold)
        out.append(info)

        _add_children(path, name, parts, chain, info, index, threshold, out)

    return out


def _add_children(
    path: str,
    name: str,
    parts: list[_Part],
    chain: list[int],
    parent_info: TableInfo,
    index: int,
    threshold: int,
    out: list[TableInfo],
) -> None:
    """Add the nested tables of one table, recursively."""
    kids = [k for part_i in chain for k in parts[part_i].children]
    kids.sort(key=lambda i: (parts[i].page, parts[i].bbox[1], parts[i].bbox[0]))
    parent_info.nested_direct = len(kids)

    for n, kid_i in enumerate(kids, start=1):
        kid = parts[kid_i]
        label = f"{parent_info.label}.{n}"
        child = TableInfo(
            file=path,
            file_name=name,
            index=index,
            label=label,
            level=parent_info.level + 1,
            parent_label=parent_info.label,
            is_nested=True,
            page_start=kid.page,
            page_end=kid.page,
            pages_spanned=1,
            rows=kid.rows,
            grid_cols=kid.cols,
            cells=kid.cells,
            merged_h=kid.merged_h,
            merged_v=kid.merged_v,
            preview=f"[inside {parent_info.label}] {kid.preview}",
        )
        _finalize(child, threshold)
        out.append(child)
        _add_children(path, name, parts, [kid_i], child, index, threshold, out)
        parent_info.nested_total += 1 + child.nested_total
