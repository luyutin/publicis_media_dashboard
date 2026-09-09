import hashlib
import io
import tempfile
import zipfile
from pathlib import Path

import streamlit as st

from process import media_cleaner as clean_media_data


def main():
    st.title('🧹 未清理資料格式化')
    st.markdown("#### **從不固定版面的 Excel 找出日期資料表，產生標準格式化數據。**")
    st.info(
        '可一次上傳多個 XLSX / XLS 檔案，並為每個檔案勾選要清整的工作表；'
        '支援同一工作表內左右並排、上下堆疊的多個資料表，並自動填滿合併儲存格。'
        '只保留具有有效日期的資料，排除彙總列與彙總欄。能對應的欄位使用系統標準名稱；'
        '無法對應的欄位保留原名稱與資料，並加入 Source 欄標示來源 Excel。'
    )

    default_dictionary = (
        Path(__file__).resolve().parent.parent
        / "Report Template & All Format 字典.xlsx"
    )
    uploaded_files = st.file_uploader(
        "上傳未清理的媒體 Excel 檔案",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
        key="uncleaned_media_files",
    )

    selected_sheets_by_file: dict[int, list[str]] = {}
    sheet_read_errors: dict[int, str] = {}
    if uploaded_files:
        st.subheader("選擇要清整的工作表")
        st.caption(
            "每個檔案預設只勾選第一張；如果還有其他資料表，請再手動勾選。"
            "程式會對每張工作表的各個資料區塊獨立偵測、清理與稽核。"
        )
        for index, uploaded_file in enumerate(uploaded_files, start=1):
            workbook_data = uploaded_file.getvalue()
            file_digest = hashlib.sha1(workbook_data).hexdigest()[:10]
            try:
                sheet_names = clean_media_data.list_workbook_sheets(workbook_data)
            except Exception as exc:
                sheet_read_errors[index] = str(exc)
                st.error(f"{uploaded_file.name}：無法讀取工作表（{exc}）")
                continue
            selected_sheets_by_file[index] = st.multiselect(
                f"{uploaded_file.name}（共 {len(sheet_names)} 張）",
                options=sheet_names,
                default=sheet_names[:1],
                key=f"uncleaned_media_sheets_{index}_{file_digest}",
            )

    with st.expander("⚙️ 進階設定"):
        scan_rows = st.number_input(
            "表頭搜尋列數（0 表示整張工作表）",
            min_value=0,
            value=0,
            step=100,
            help="預設搜尋整張工作表，避免漏掉下方資料表；正數只限制表頭搜尋，不限制資料列數。",
        )
        default_year = st.number_input(
            "缺少年份的日期使用哪一年（0 表示僅依報表內容判斷）",
            min_value=0, max_value=2100, value=0, step=1,
            help="例如文字 7月28日。完整日期不受影響；無法確認年份的列會排除並列入稽核。",
        )
        dictionary_file = st.file_uploader(
            "自訂欄位字典（選填，XLSX）",
            type=["xlsx"],
            accept_multiple_files=False,
            key="uncleaned_media_dictionary",
        )
        if default_dictionary.exists():
            st.download_button(
                "📥 下載自訂欄位字典範本",
                data=default_dictionary.read_bytes(),
                file_name="自訂欄位字典_範本.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="download_media_dictionary_template",
            )
            st.caption(
                "填寫時請保留「字典欄位說明」工作表與 B/C、H–L 欄位結構；"
                "在對應的系統欄位列填入來源報表欄名。"
            )
        else:
            st.warning("找不到預設欄位字典範本，請聯絡系統管理者。")
        use_ollama = st.checkbox(
            "使用本機 Ollama 辨識並列表格數量與陌生欄位",
            value=True,
            help="找到日期表頭後，會將完整橫列、欄座標及附近樣本送到 Ollama，確認其他表格與共用日期，再辨識陌生欄位。提案未通過驗證時會使用規則並記錄原因。",
        )
        if use_ollama:
            ollama_model = st.text_input("Ollama 模型", value="qwen3.5:2b")
            ollama_url = st.text_input(
                "Ollama API 網址", value="http://127.0.0.1:11434"
            )
            ollama_timeout = st.number_input(
                "Ollama 等待秒數",
                min_value=1.0,
                max_value=600.0,
                value=120.0,
                step=5.0,
            )
        else:
            ollama_model = "qwen3.5:2b"
            ollama_url = "http://127.0.0.1:11434"
            ollama_timeout = 120.0

    if st.button("開始格式化", type="primary", key="clean_uncleaned_media"):
        if not uploaded_files:
            st.error("請先上傳至少一個 XLSX / XLS 檔案。")
        elif default_year and default_year < 1900:
            st.error("請將年份設為 0，或輸入 1900–2100 之間的年份。")
        elif sheet_read_errors:
            st.error("有檔案無法讀取工作表，請確認檔案內容後再試。")
        elif any(
            not selected_sheets_by_file.get(index)
            for index in range(1, len(uploaded_files) + 1)
        ):
            missing_files = [
                uploaded_file.name
                for index, uploaded_file in enumerate(uploaded_files, start=1)
                if not selected_sheets_by_file.get(index)
            ]
            st.error(
                "請為以下檔案至少勾選一張工作表："
                + "、".join(missing_files)
            )
        else:
            st.session_state.pop("uncleaned_media_results", None)
            with st.spinner("正在偵測表頭並格式化資料，請稍候…"):
                with tempfile.TemporaryDirectory(prefix="media_cleaner_") as temp_dir:
                    temp_root = Path(temp_dir)
                    output_dir = temp_root / "cleaned_output"
                    output_dir.mkdir()

                    if dictionary_file is not None:
                        dictionary_path = temp_root / "dictionary.xlsx"
                        dictionary_path.write_bytes(dictionary_file.getvalue())
                    else:
                        dictionary_path = (
                            default_dictionary if default_dictionary.exists() else None
                        )

                    try:
                        aliases = clean_media_data.read_dictionary(dictionary_path)
                    except Exception as exc:
                        st.error(f"欄位字典無法讀取：{exc}")
                        return
                    ollama = clean_media_data.OllamaConfig(
                        model=ollama_model,
                        url=ollama_url,
                        timeout=float(ollama_timeout),
                        enabled=use_ollama,
                        # Temporary diagnostic mode requested for tracing every
                        # dated table through Ollama, even after another request fails.
                        force_all=use_ollama,
                    )
                    all_audit = []
                    cleaned_files = []
                    used_output_names = set()

                    for index, uploaded_file in enumerate(uploaded_files, start=1):
                        safe_name = Path(uploaded_file.name).name
                        input_dir = temp_root / f"input_{index}"
                        input_dir.mkdir()
                        input_path = input_dir / safe_name
                        input_path.write_bytes(uploaded_file.getvalue())

                        output_name = f"{Path(safe_name).stem}_cleaned.xlsx"
                        if output_name in used_output_names:
                            output_name = (
                                f"{Path(safe_name).stem}_cleaned_{index}.xlsx"
                            )
                        used_output_names.add(output_name)
                        output_path = output_dir / output_name

                        try:
                            audit, rows = clean_media_data.clean_workbook_sheets(
                                input_path=input_path,
                                output_path=output_path,
                                aliases=aliases,
                                scan_rows=int(scan_rows),
                                sheet_names=selected_sheets_by_file[index],
                                ollama=ollama,
                                default_year=int(default_year) or None,
                            )
                        except Exception as exc:
                            audit = [
                                clean_media_data.AuditRecord(
                                    input_file=safe_name,
                                    sheet="、".join(selected_sheets_by_file[index]),
                                    header_row=None,
                                    status=f"處理失敗：{exc}",
                                    score=None,
                                    data_rows=0,
                                    mapped_columns={},
                                    unmapped_columns=[],
                                    output_file=None,
                                )
                            ]
                            rows = 0

                        for record in audit:
                            record.input_file = safe_name
                            record.output_file = (
                                output_name if record.data_rows else None
                            )
                        all_audit.extend(audit)
                        if rows and output_path.exists():
                            cleaned_files.append(
                                {
                                    "name": output_name,
                                    "data": output_path.read_bytes(),
                                    "rows": rows,
                                }
                            )

                    clean_media_data.write_audit(all_audit, output_dir)
                    audit_json = (output_dir / "cleaning_audit.json").read_bytes()
                    unmapped_csv = (output_dir / "unmapped_columns.csv").read_bytes()

                    zip_buffer = io.BytesIO()
                    with zipfile.ZipFile(
                        zip_buffer, "w", compression=zipfile.ZIP_DEFLATED
                    ) as archive:
                        for cleaned_file in cleaned_files:
                            archive.writestr(
                                cleaned_file["name"], cleaned_file["data"]
                            )
                        archive.writestr("cleaning_audit.json", audit_json)
                        archive.writestr("unmapped_columns.csv", unmapped_csv)

                    consolidated_path = temp_root / "cleaned_media_results.xlsx"
                    consolidated_rows = (
                        clean_media_data.consolidate_cleaned_workbooks(
                            cleaned_paths=[
                                output_dir / cleaned_file["name"]
                                for cleaned_file in cleaned_files
                            ],
                            records=all_audit,
                            output_path=consolidated_path,
                        )
                    )

                    st.session_state.uncleaned_media_results = {
                        "files": cleaned_files,
                        "consolidated": consolidated_path.read_bytes(),
                        "consolidated_rows": consolidated_rows,
                        "audit": [
                            {
                                "檔案": record.input_file,
                                "工作表": record.sheet,
                                "區塊": record.block_id,
                                "來源範圍": record.source_range,
                                "日期來源": record.date_source_range,
                                "辨識方式": record.detection_method,
                                "表頭列": record.header_row,
                                "狀態": record.status,
                                "欄位對應率": record.score,
                                "有效資料列": record.data_rows,
                                "欄位對應": record.mapped_columns,
                                "未對應欄位（已保留）": record.unmapped_columns,
                                "排除欄位": record.excluded_columns,
                                "略過資料列": record.skipped_rows,
                                "提醒": record.warnings,
                                "已展開合併範圍": record.merged_ranges,
                            }
                            for record in all_audit
                        ],
                        "audit_json": audit_json,
                        "unmapped_csv": unmapped_csv,
                        "zip": zip_buffer.getvalue(),
                        "total": len(uploaded_files),
                    }

    results = st.session_state.get("uncleaned_media_results")
    if not results:
        return

    succeeded = len(results["files"])
    if succeeded:
        st.success(f"格式化完成：{succeeded}/{results['total']} 個檔案成功輸出。")
    else:
        st.warning("處理完成，但沒有檔案產生有效的格式化資料。請查看稽核結果。")

    st.subheader("處理結果")
    summary_rows = [
        {
            key: (
                "、".join(value)
                if isinstance(value, list)
                else str(value) if isinstance(value, dict)
                else value
            )
            for key, value in record.items()
            if key != "欄位對應"
        }
        for record in results["audit"]
    ]
    st.dataframe(summary_rows, use_container_width=True, hide_index=True)

    for record in results["audit"]:
        if record["欄位對應"]:
            with st.expander(
                f"欄位對應：{record['檔案']} / {record['工作表']} / {record.get('區塊', '')} {record.get('來源範圍', '')}"
            ):
                st.json(record["欄位對應"])

    st.subheader("下載")
    with st.expander("📦 輸出檔案", expanded=False):
        if results.get("consolidated") is not None:
            st.download_button(
                f"下載匯總結果（Excel，共 {results.get('consolidated_rows', 0):,} 列）",
                data=results["consolidated"],
                file_name="cleaned_media_results.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="download_consolidated_cleaned_media",
                type="primary",
            )
            st.caption(
                "內含 cleaned_data、cleaning_audit、unmapped_columns 三張工作表；"
                "未對應欄位仍會保留在 cleaned_data，unmapped_columns 用於檢查欄位判別結果。"
            )
        else:
            st.info("目前是更新前的舊結果；請重新執行格式化以產生匯總 Excel。")
        if results["files"]:
            st.download_button(
                "下載全部結果（ZIP）",
                data=results["zip"],
                file_name="cleaned_media_results.zip",
                mime="application/zip",
                key="download_all_cleaned_media",
            )
            for index, cleaned_file in enumerate(results["files"], start=1):
                st.download_button(
                    f"下載 {cleaned_file['name']}（{cleaned_file['rows']} 列）",
                    data=cleaned_file["data"],
                    file_name=cleaned_file["name"],
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"download_cleaned_media_{index}",
                )

        audit_col, unmapped_col = st.columns(2)
        with audit_col:
            st.download_button(
                "下載完整稽核報告（JSON）",
                data=results["audit_json"],
                file_name="cleaning_audit.json",
                mime="application/json",
                key="download_cleaning_audit",
            )
        with unmapped_col:
            st.download_button(
                "下載未對應欄位清單（欄位已保留於結果）",
                data=results["unmapped_csv"],
                file_name="unmapped_columns.csv",
                mime="text/csv",
                key="download_unmapped_columns",
            )
