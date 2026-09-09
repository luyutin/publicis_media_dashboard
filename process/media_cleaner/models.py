"""Contracts for independent worksheet regions and their audit trail."""

from dataclasses import dataclass, field


@dataclass
class Candidate:
    sheet: str
    header_row: int
    score: float
    headers: list[str]
    mapped: dict[int, str]
    unmapped: list[str]
    data_rows: int
    start_column: int = 1
    end_column: int = 1
    end_row: int = 1
    header_start: int = 1
    block_id: str = ""
    method: str = "rules"
    excluded_columns: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    shared_date_column: int | None = None
    group_label: str = ""


@dataclass
class AuditRecord:
    input_file: str
    sheet: str
    header_row: int | None
    status: str
    score: float | None
    data_rows: int
    mapped_columns: dict[str, str]
    unmapped_columns: list[str]
    output_file: str | None
    block_id: str = ""
    source_range: str = ""
    detection_method: str = "rules"
    excluded_columns: list[str] = field(default_factory=list)
    skipped_rows: dict[str, int] = field(default_factory=dict)
    warnings: list[str] = field(default_factory=list)
    merged_ranges: list[str] = field(default_factory=list)
    output_start_row: int | None = None
    output_end_row: int | None = None
    date_source_range: str = ""


@dataclass
class OllamaConfig:
    model: str = "qwen3.5:9b"
    url: str = "http://127.0.0.1:11434"
    timeout: float = 120.0
    enabled: bool = True
    # Temporary diagnostic mode: ask Ollama about every dated block, including
    # blocks whose columns were already resolved by aliases.
    force_all: bool = False
    warned: bool = False
    # Diagnostic only. A failed request must not disable later LLM requests.
    failure: str | None = None
