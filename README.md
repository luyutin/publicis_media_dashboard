# Media Report Dashboard & ROI Modeller+

Streamlit dashboard for formatting media reports, cleaning unknown Excel layouts,
and running ROI analysis. The unknown-data formatter can use a local Ollama model
to assist with header detection and automatically falls back to deterministic rules
when Ollama is unavailable.

## Requirements

- Python 3.11
- Ollama (optional, recommended for unknown Excel layouts)
- Ollama model `qwen3.5:9b` (about 6.6 GB)

## Set up on another computer

Clone this repository, open a terminal in the project directory, and create a new
virtual environment. Do not copy `.venv` from another computer.

### Windows PowerShell

```powershell
py -3.11 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
```

### macOS or Linux

```bash
python3.11 -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
```

For AI-assisted header detection, install Ollama from its official website and run:

```bash
ollama pull qwen3.5:9b
```

Ollama normally exposes its local API at `http://127.0.0.1:11434`. The application
still works in rule-only mode if Ollama is not installed or the option is disabled.

## Run

```bash
python run_main.py
```

Alternatively:

```bash
streamlit run main.py
```

## Repository contents

- `main.py`: Streamlit navigation and page entry points
- `process/`: dashboard pages and media-specific formatters
- `process/media_cleaner/`: unknown Excel cleaning engine and editable text configuration
- `Report Template & All Format 字典.xlsx`: default field dictionary
- `Photos/`: images used by the upload instructions

Customer workbooks, local outputs, virtual environments, generated executables,
and legacy project copies are intentionally excluded from Git.

### Unknown-data formatter maintenance

The cleaner is a package under `process/media_cleaner/`. Its Python API is in
`engine.py`; editable aliases, field descriptions, and Ollama prompts are under
`process/media_cleaner/config/`. The canonical field names and order come from
`process/template_schema.py::TEMPLATE_COLUMNS` (shared with the old formatter).
Every retained dated row also contains `Source`, whose value is the original Excel
filename. Source columns mapped by aliases or the LLM use their canonical template
names. Every other non-total source column is preserved with its original header
after `Source`, including declared columns whose values are all blank. Duplicate
original headers receive an Excel-column suffix so their values cannot overwrite
one another. Restart Streamlit after changing the configuration files.

On the Streamlit page, each uploaded workbook shows its worksheet names. Only
the first worksheet is selected by default; users can explicitly include more.
Each selected XLSX or XLS worksheet may contain multiple tables. Actual merged
cells are expanded in memory; unrelated blank cells are not forward-filled.
Only dated records are exported. Total rows, total columns, aggregate column
groups, and non-date demographic tables are excluded and audited.

The pipeline is split into `workbooks.py` (XLS/XLSX input), `grid.py` (merge
expansion), `detection.py` (rule-based header seeds and fallback), `structure.py`
(LLM table discovery across a complete header row), `mapping.py` (aliases and
LLM field semantics), `extraction.py` (date validation), and `output.py` (Excel
exports). `engine.py` orchestrates these steps. It does not invoke the old
customer-specific formatters or write their ROI session state.

With Ollama enabled, finding one dated seed triggers a full-width row request,
including absolute coordinates, nearby samples, and merge ranges. The model can
return more tables than the seeds, different table widths/metric orders, and a
shared date column anywhere on the row. Proposals must cover the source columns,
reference real cells, avoid duplicate metric extraction, and have date evidence.
Invalid proposals get one correction attempt, then fall back with an audit warning.
Unknown field mapping is a separate model step. Raw model confidence is not a
guarantee that a table is correct.

The default header scan covers the entire sheet (`--scan-rows 0`); lower stacked
tables are discovered by continuing the rule scan. A missing year is resolved only
from local evidence or `--default-year`, never from the current date. Audits include
source range, date source, merged ranges, skipped rows, mapping method, warnings,
and output row intervals. The displayed mapping ratio is field coverage, not an
accuracy/confidence score.

Limits: a discoverable seed is still needed; complex layouts without one may be
missed. Rule-only fallback has narrower layout support. Unfamiliar metrics without
a corresponding standard field remain marked as unmapped in the audit while their
original columns and values are retained in `cleaned_data`; rates and video
durations are not silently relabelled as different metrics. Formula values must have been calculated
and saved by the spreadsheet application. No automatic overlap/deduplication across
independent source files is performed.

The standalone CLI is:

```bash
python -m process.media_cleaner input.xlsx --no-ollama
python -m process.media_cleaner input.xls --all-sheets --default-year 2025
```

Run deterministic regressions with `python -m unittest discover -s tests`.
`python -m tests.manual_validate_media_cleaner --output /tmp/media-validation`
runs opt-in real-model layout variations and checks row counts and metric sums.
Add `--source <directory>` to test private workbooks; these are development data,
not an independent generalization benchmark. Stubbed LLM tests validate the
pipeline contract separately from real-model recognition quality.

## Update an existing Windows installation

For a source-code installation, close the running dashboard, open PowerShell in
the project directory, then run:

```powershell
git pull origin main
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python run_main.py
```

If the user runs a packaged `.exe`, pulling the repository does not update that
executable. Rebuild it on Windows from the latest `main` branch and replace the
old distribution. The build must include both `Report Template & All Format
字典.xlsx` and the entire `process/media_cleaner/config` directory.

## Building a Windows executable

Build the executable on Windows after installing the dependencies. Ollama and the
model remain separate prerequisites and are not embedded in the executable. Ensure
the field dictionary, Streamlit static assets, and the complete
`process/media_cleaner/config` directory are included in the PyInstaller configuration.
