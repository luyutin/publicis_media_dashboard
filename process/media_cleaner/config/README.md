# Media cleaner editable configuration

These files are loaded when `process.media_cleaner` is imported. Restart the
Streamlit application after editing them.

- `target_columns.txt`: one tab-delimited `template column<TAB>description` per
  line. It must contain exactly the canonical columns defined by
  `process/template_schema.py::TEMPLATE_COLUMNS`; output order follows that schema.
- `aliases.json`: deterministic source-header aliases grouped by target field.
  Fields absent from the aliases and not mapped by the model are still retained
  under their original source headers; this file controls naming, not retention.
- `ollama_system_prompt.txt`: stable model role, guardrails, and exclusions.
- `ollama_layout_prompt.txt`: full-row table-count, range, and shared-date discovery.
  This prompt is read at each layout request. The layout response contract is in
  `structure.py`; it is separate from the field-mapping response contract.
- `ollama_user_prompt.txt`: task template. Keep the placeholders
  `{target_descriptions}` and `{options}`.

When building a PyInstaller executable, include this entire directory at
`process/media_cleaner/config` inside the bundle.
