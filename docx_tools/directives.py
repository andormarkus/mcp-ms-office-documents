"""Document-level directive parsing.

The renderer already has a *per-block* comment-directive mechanism
(``COMMENT_DIRECTIVE_PATTERN`` in ``patterns.py``): one or more ``<!-- key:
value -->`` lines immediately above a block (a table, paragraph, list, …)
attach options to that single block only (``borderless``, ``widths``,
``style``, …). This module adds a second, *document-level* tier: a handful of
keys — recognised only when they appear in the leading run of directive lines
at the very top of the markdown content — configure the document as a whole
(page size, margins, default font, multi-column layout, smart quotes) rather
than the block that follows them.

Recognised document-level directives are removed from the content before it
reaches the block parser so they are never mistaken for (harmless but
meaningless) per-block directives. Any other directive in that same leading
run — e.g. ``<!-- borderless -->`` meant for the very first table — is left
untouched so the existing per-block mechanism still sees it.
"""
from .patterns import COMMENT_DIRECTIVE_PATTERN

# Keys recognised as document-wide settings. Anything else found in the
# leading directive run is assumed to be a per-block directive for whatever
# follows and is left in the content.
DOCUMENT_LEVEL_KEYS = {
    "page",
    "orientation",
    "margin",
    "font",
    "font-size",
    "columns",
    "heading-font",
    "smart-quotes",
}


def parse_document_directives(content: str):
    """Extract leading document-level directives from *content*.

    Only the leading run of the document is scanned: consecutive
    ``<!-- key: value -->`` lines (blank lines between them are tolerated) up
    to the first line that is neither blank nor a directive. This mirrors how
    a document typically opens (front-matter-style options before any real
    content) without needing a special delimiter.

    Returns:
        ``(directives, remaining_content)`` — *directives* is a ``{key:
        value}`` dict (empty if none found); *remaining_content* is *content*
        with the recognised document-level directive lines removed (other
        lines, including non-document-level directives and blank lines, are
        preserved verbatim so downstream line-based parsing is unaffected).
    """
    if not content:
        return {}, content
    lines = content.split("\n")
    directives: dict = {}
    consumed = set()
    for idx, line in enumerate(lines):
        stripped = line.strip()
        if not stripped:
            continue
        match = COMMENT_DIRECTIVE_PATTERN.match(stripped)
        if not match:
            break
        key = match.group(1).strip().lower()
        value = (match.group(2) or "").strip()
        if key in DOCUMENT_LEVEL_KEYS:
            directives[key] = value
            consumed.add(idx)
    if not directives:
        return {}, content
    remaining = "\n".join(line for i, line in enumerate(lines) if i not in consumed)
    return directives, remaining
