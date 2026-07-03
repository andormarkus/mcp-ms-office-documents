# AGENTS.md

## Project Overview

MCP (Model Context Protocol) server built with **FastMCP 3.0** that exposes Office document generation as MCP tools. Runs as a Docker container (Python 3.12, Alpine) on port **8958** at `/mcp` using streamable-HTTP transport. Entry point: `main.py`.

## Architecture

```
main.py                  ← Registers all MCP tools on a single FastMCP instance
├── {docx,xlsx,pptx,email,xml}_tools/
│   ├── __init__.py      ← Re-exports the public function (e.g. markdown_to_word)
│   ├── base_*_tool.py   ← Core conversion logic (markdown → document bytes)
│   ├── helpers.py        ← Parsing, formatting, shared utilities
│   └── dynamic_*_tools.py  ← YAML-driven tool registration (docx, email only)
├── upload_tools/
│   ├── main.py          ← upload_file() dispatches to strategy backend
│   └── backends/{local,s3,gcs,azure,minio}.py
├── config.py            ← Singleton Config from env vars (Pydantic v2), logging setup
├── template_utils.py    ← Template resolution: custom_templates/ → default_templates/
└── middleware.py         ← Optional API key auth (Bearer / x-api-key header)
```

**Data flow:** Every tool converts input → in-memory bytes → calls `upload_file(file_obj, suffix)` → backend saves/uploads → returns URL or path string to the MCP client.

## Key Conventions

- **Config is centralized in `config.py`** — no module reads `os.environ` directly. Access via `get_config()` singleton.
- **Template resolution** (`template_utils.py`): searches `custom_templates/` before `default_templates/`, with `/app/*` container paths tried first, then local paths. Never hardcode template paths.
- **Dynamic tool registration**: YAML files in `config/` define parameterized email/docx templates. Each entry becomes a separate MCP tool at startup via `register_*_template_tools_from_yaml(mcp, path)`. Placeholders use Mustache syntax `{{name}}`. See `config/docx_templates.yaml` for the canonical example.
  - **Merged loading**: the master `config/<kind>_templates.yaml` is merged with UI-managed per-template files in `config/<kind>_templates.d/<name>.yaml` (the `.d` file wins on a name clash). The documented master files are never rewritten by tooling. Merge logic lives in `template_registry.py`.
  - **Live (un)registration**: `register_docx_template` / `register_email_template` (and their `unregister_*` counterparts) can add/replace/remove tools on a running `FastMCP` instance via `local_provider`, so the admin UI applies template edits without a restart. A module-level registry tracks live dynamic tool names.
- **Admin UI** (`admin/`, opt-in via `ADMIN_ENABLED`): a FastHTML template-admin for creating/editing dynamic docx+email templates without YAML. When enabled, `main.py` runs the **combined ASGI app** (`admin.app.build_combined_app`) under uvicorn — the MCP endpoint mounted at `/` and the admin UI under `ADMIN_PATH` (default `/admin`) in one process, so saving a template registers its MCP tool live. Modules: `store.py` (`TemplateStore` + `FileTemplateStore`, abstracted for a future shared backend), `analysis.py` (placeholder/conditional/style detection on uploaded assets), `preview.py` (render with sample values, no upload backend), `auth.py` (shared-password gate), `app.py` (views + app factory, including the Status page and document re-upload/re-scan). When `ADMIN_ENABLED` is unset, `main.py` keeps the original `mcp.run()` path unchanged.
- **Metrics** (`metrics.py`, project root): tiny always-on in-process counters (per-template calls/errors/last-used, recorded by the dynamic tool wrappers) plus a bounded recent-log ring buffer (installed by the admin app at startup). Surfaced on the admin Status page. No external deps; safe to import from the core tool modules.
- **Pydantic models** for tool arguments are created dynamically with `create_model()` in `dynamic_*_tools.py`. The `TYPE_MAP` dict maps YAML type strings to Python types.
- **Error handling in tools**: raise `fastmcp.exceptions.ToolError` for user-facing errors; use `RuntimeError` in upload/backend layers.
- **Logging**: use `logging.getLogger(__name__)` everywhere. Level controlled by `DEBUG` env var only.

## DOCX-Specific Conventions

- **Markdown directives** are parsed as `<!-- key: value -->` lines. There are two tiers, both handled by `docx_tools/patterns.py`'s `COMMENT_DIRECTIVE_PATTERN`:
  - **Per-block directives** (`markdown_processor.py`): one or more directive lines immediately above a block attach options to that block only — `borderless`, `widths`, `style`, `shade` (table shading, e.g. `header=D5E8F0 alt=F2F2F2`), `bookmark` (wraps the block in a named `w:bookmarkStart`/`End` pair), `tab` (configures an explicit tab stop, e.g. `right` or `right dot`, on the following paragraph).
  - **Document-level directives** (`docx_tools/directives.py::parse_document_directives`): recognised only in the leading run of directive lines at the very top of the markdown, before any real content — `page` (`letter`/`a4`/`legal`), `orientation` (`landscape`), `margin` (`1in` or `top=1in bottom=1in ...`), `font`/`font-size`/`heading-font`, `columns` (`2`, `2 sep`, or `custom:5400,3240`), `smart-quotes` (`on`). These are extracted and stripped from the content before block parsing; unrecognized keys in that same leading run (e.g. a `borderless` meant for the first table) are left in place. Applied in `base_docx_tool.py::_markdown_to_doc` via `document_features.py`'s `apply_page_setup`/`apply_default_font`/`apply_columns` — tool parameters of the same name (`page_size`, `orientation`, `default_font`, `smart_quotes`) override the directive when both are given.
- **Footnotes**: inline `[^footnote text]` syntax (not a two-step id/definition — the bracket contains the footnote body directly). Implemented by `document_features.py`'s `insert_footnote`/`add_footnote`/`add_footnote_reference_run`, which build `word/footnotes.xml` directly via `docx.opc.part.XmlPart` since python-docx (1.2.0) has no footnote API (unlike `NumberingPart`/`CommentsPart`).
- **Internal hyperlinks**: `[text](#bookmark)` — a URL starting with `#` emits `w:anchor` instead of an external relationship (`inline_formatting.py::add_hyperlink`). Pair with a `<!-- bookmark: bookmark -->` directive to define the target.
- **Smart quotes** (`patterns.py::apply_smart_quotes`) are applied as a single whole-document preprocessing pass (opt-in — off by default so existing content asserting on literal quote characters is unaffected), not threaded through every inline-formatting call site. Fenced code blocks are skipped entirely; inline code spans are protected via a placeholder swap.
- All of the above are **pure Python** (`python-docx` + direct `OxmlElement`/`opc` manipulation) — no LibreOffice or other external binary is used anywhere in `docx_tools/`.

## Adding a New Document Tool

1. Create `<type>_tools/` package with `__init__.py`, `base_<type>_tool.py`, and optional `helpers.py`.
2. The base tool function should: accept content → produce an `io.BytesIO` → call `upload_file(buffer, "<ext>")` → return the result string.
3. Register the async wrapper in `main.py` using `@mcp.tool(name=..., description=..., tags=..., annotations=...)`.
4. Use `Annotated[<type>, Field(description=...)]` for all tool parameters — the descriptions are critical because MCP clients (AI models) rely on them.

## Tests

```bash
pytest                        # Run all tests (asyncio_mode=auto in pytest.ini)
pytest tests/test_docx_base.py  # Single module
```

- Tests live in `tests/` and output generated files to `tests/output/{docx,pptx,xlsx}/` for manual inspection.
- Upload is mocked in tests — patch `upload_file` or the specific `*_tool.upload_file` to capture bytes without needing a real backend. See `test_xlsx_creation.py::_create_workbook_from_markdown` for the pattern.
- PPTX tests instantiate `PowerpointPresentation` directly and call `.save()` to get a buffer, bypassing upload entirely.
- No `.env` required for tests — `config.py` defaults to `LOCAL` strategy and INFO logging.
