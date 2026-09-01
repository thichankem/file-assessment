"""docstats - table and figure statistics for Word and PDF documents."""

from .analyzer import (
    DEFAULT_COMPLEX_THRESHOLD,
    DEFAULT_FIGURE_MIN_AREA,
    DocResult,
    FigureInfo,
    ScanResult,
    TableInfo,
    analyze_folder,
    analyze_file,
    find_documents,
    open_document,
    open_in_word,
    recompute_complexity,
    recompute_figures,
)

__all__ = [
    "DEFAULT_COMPLEX_THRESHOLD",
    "DEFAULT_FIGURE_MIN_AREA",
    "DocResult",
    "FigureInfo",
    "ScanResult",
    "TableInfo",
    "analyze_folder",
    "analyze_file",
    "find_documents",
    "open_document",
    "open_in_word",
    "recompute_complexity",
    "recompute_figures",
]
