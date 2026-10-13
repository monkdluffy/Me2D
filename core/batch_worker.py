import os
import time
import logging
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import List, Dict, Any, Callable, Optional

from core.doc_engine import DocumentEngine
from core.pdf_engine import PDFEngine

logger = logging.getLogger("batch_worker")

class BatchProcessor:
    """Multi-threaded Batch Document Processing Engine using ThreadPoolExecutor."""

    def __init__(self, max_workers: int = None):
        self.max_workers = max_workers or min(2, os.cpu_count() or 1)
        self.is_cancelled = False

    def cancel(self):
        """Cancel ongoing batch execution."""
        self.is_cancelled = True

    def process_batch(
        self,
        template_paths: List[str],
        data_records: List[Dict[str, Any]],
        output_dir: str,
        convert_to_pdf: bool = True,
        watermark_text: str = "",
        stamp_image_path: str = "",
        image_map: Optional[Dict[str, Dict[str, Any]]] = None,
        libreoffice_path: str = "",
        custom_find_replace: Optional[Dict[str, Any]] = None,
        on_progress: Optional[Callable[[int, int, str, str], None]] = None,
        on_log: Optional[Callable[[str], None]] = None,
        on_complete: Optional[Callable[[Dict[str, Any]], None]] = None
    ):
        """Executes a batch job concurrently across all template and data row combinations.
        
        Args:
            template_paths: List of DOCX template file paths.
            data_records: List of 78-variable data context dictionaries.
            output_dir: Output folder path.
            convert_to_pdf: If True, automatically convert generated DOCX files to PDF.
            watermark_text: Text for PDF watermark (e.g. "DRAFT", "CONFIDENTIAL", "").
            stamp_image_path: Image path for PDF stamp overlay.
            image_map: Image configuration for docxtpl InlineImage substitution.
            libreoffice_path: Path to soffice executable for PDF conversion.
            custom_find_replace: Optional dict with {"find_text": str, "replace_text": str, "case_sensitive": bool}.
            on_progress: Callback function (current, total, file_name, status).
            on_log: Callback function (log_message).
            on_complete: Callback function (summary_info).
        """
        self.is_cancelled = False
        start_time = time.time()

        tasks = []
        # Build task list: each combination of (template, data_row)
        for t_idx, template_path in enumerate(template_paths):
            t_name = os.path.splitext(os.path.basename(template_path))[0]
            for r_idx, data in enumerate(data_records):
                actual_t_name = t_name.replace("{{inisial_bu}}", str(data.get("inisial_bu", "")).strip())
                out_docx_name = f"{actual_t_name}_{r_idx + 1}.docx" if len(data_records) > 1 else f"{actual_t_name}.docx"
                out_docx_path = os.path.join(output_dir, out_docx_name)

                tasks.append({
                    "task_id": f"T{t_idx+1}-R{r_idx+1}",
                    "template_path": template_path,
                    "data": data,
                    "out_docx_path": out_docx_path,
                    "convert_pdf": convert_to_pdf,
                    "watermark_text": watermark_text,
                    "stamp_image_path": stamp_image_path,
                    "image_map": image_map,
                    "libreoffice_path": libreoffice_path,
                    "custom_find_replace": custom_find_replace
                })

        total_tasks = len(tasks)
        completed_count = 0
        success_count = 0
        error_count = 0
        results_summary = []

        if on_log:
            on_log(f"🚀 Starting batch execution with {total_tasks} task(s) using {self.max_workers} thread worker(s)...")
            if custom_find_replace:
                pairs_to_log = []
                if isinstance(custom_find_replace, list):
                    pairs_to_log = custom_find_replace
                elif isinstance(custom_find_replace, dict):
                    if "pairs" in custom_find_replace and isinstance(custom_find_replace["pairs"], list):
                        default_cs = custom_find_replace.get("case_sensitive", True)
                        for item in custom_find_replace["pairs"]:
                            p_copy = dict(item)
                            if "case_sensitive" not in p_copy:
                                p_copy["case_sensitive"] = default_cs
                            pairs_to_log.append(p_copy)
                    elif custom_find_replace.get("find_text"):
                        pairs_to_log = [custom_find_replace]
                
                valid_pairs = [p for p in pairs_to_log if p.get("find_text")]
                if valid_pairs:
                    on_log(f"🔍 Batch Find & Replace active: {len(valid_pairs)} aturan pencarian/penggantian:")
                    for p_idx, p in enumerate(valid_pairs, 1):
                        cs_label = "Case Sensitive" if p.get("case_sensitive", True) else "Non-Case Sensitive"
                        on_log(f"   [{p_idx}] '{p['find_text']}' ➔ '{p.get('replace_text', '')}' ({cs_label})")

        def _execute_single_task(task_info: Dict[str, Any]) -> Dict[str, Any]:
            if self.is_cancelled:
                return {"success": False, "task_id": task_info["task_id"], "error": "Cancelled"}

            # Step 1: Process DOCX replacement (with custom find & replace)
            success, res_msg = DocumentEngine.process_document(
                template_path=task_info["template_path"],
                output_path=task_info["out_docx_path"],
                context=task_info["data"],
                image_map=task_info["image_map"],
                custom_find_replace=task_info.get("custom_find_replace")
            )

            if not success:
                return {"success": False, "task_id": task_info["task_id"], "docx_path": task_info["out_docx_path"], "error": res_msg}

            # --- Dynamic Highlight based on tgl_daftar_periksa ---
            tgl = str(task_info["data"].get("tgl_daftar_periksa", "")).lower()
            
            # Hanya proses jika file merupakan template DO-06-02-FR-9.2-02
            path_check = str(task_info["template_path"]).lower()
            is_audit_program = "do-06-02-fr-9.2-02" in path_check or "program audit internal ai" in path_check
            
            if is_audit_program:
                month_mapping = {
                    "januari": 1, "jan": 1,
                    "februari": 2, "feb": 2,
                    "maret": 3, "mar": 3,
                    "april": 4, "apr": 4,
                    "mei": 5, "may": 5,
                    "juni": 6, "jun": 6,
                    "juli": 7, "jul": 7,
                    "agustus": 8, "agst": 8, "ags": 8, "aug": 8,
                    "september": 9, "sept": 9, "sep": 9,
                    "oktober": 10, "okt": 10, "oct": 10,
                    "november": 11, "nov": 11,
                    "desember": 12, "des": 12, "dec": 12
                }
                target_month_idx = None
                for word in tgl.split():
                    if word in month_mapping:
                        target_month_idx = month_mapping[word]
                        break
                
                if target_month_idx:
                    DocumentEngine.highlight_audit_month(task_info["out_docx_path"], target_month_idx)
            # -----------------------------------------------------

            final_pdf_path = None
            # Step 2: Convert to PDF if requested
            if task_info["convert_pdf"]:
                pdf_out_path = os.path.splitext(task_info["out_docx_path"])[0] + ".pdf"
                pdf_ok, pdf_msg = PDFEngine.convert_docx_to_pdf(
                    docx_path=task_info["out_docx_path"],
                    output_pdf_path=pdf_out_path,
                    custom_libreoffice_path=task_info["libreoffice_path"]
                )

                if pdf_ok:
                    final_pdf_path = pdf_out_path
                    # Step 3: Apply Watermark if specified
                    if task_info["watermark_text"]:
                        wm_pdf_path = os.path.splitext(pdf_out_path)[0] + "_wm.pdf"
                        wm_ok, wm_msg = PDFEngine.apply_watermark(
                            pdf_path=pdf_out_path,
                            output_pdf_path=wm_pdf_path,
                            watermark_text=task_info["watermark_text"]
                        )
                        if wm_ok:
                            os.remove(pdf_out_path)
                            os.rename(wm_pdf_path, pdf_out_path)

                    # Step 4: Apply Stamp/Signature if specified
                    if task_info["stamp_image_path"] and os.path.exists(task_info["stamp_image_path"]):
                        st_pdf_path = os.path.splitext(pdf_out_path)[0] + "_st.pdf"
                        st_ok, st_msg = PDFEngine.apply_stamp_or_signature(
                            pdf_path=pdf_out_path,
                            output_pdf_path=st_pdf_path,
                            image_path=task_info["stamp_image_path"],
                            position_preset="bottom-right"
                        )
                        if st_ok:
                            os.remove(pdf_out_path)
                            os.rename(st_pdf_path, pdf_out_path)
                            
                    # Remove the intermediate DOCX file since user only wants PDF
                    try:
                        if os.path.exists(task_info["out_docx_path"]):
                            os.remove(task_info["out_docx_path"])
                    except Exception as e:
                        logger.warning(f"Failed to remove intermediate DOCX: {e}")
                else:
                    err_msg = f"❌ PDF conversion failed for '{os.path.basename(task_info['out_docx_path'])}': {pdf_msg}"
                    logger.warning(err_msg)
                    if on_log:
                        on_log(err_msg)

            return {
                "success": True,
                "task_id": task_info["task_id"],
                "docx_path": task_info["out_docx_path"],
                "pdf_path": final_pdf_path,
                "error": None
            }

        # Run ThreadPoolExecutor
        with ThreadPoolExecutor(max_workers=self.max_workers) as executor:
            future_to_task = {executor.submit(_execute_single_task, t): t for t in tasks}

            for future in as_completed(future_to_task):
                if self.is_cancelled:
                    if on_log:
                        on_log("⚠️ Batch operation cancelled by user.")
                    break

                task_info = future_to_task[future]
                completed_count += 1
                try:
                    res = future.result()
                    if res["success"]:
                        success_count += 1
                        msg = f"SUCCESS: Generated {os.path.basename(res['docx_path'])}"
                        if res.get("pdf_path"):
                            msg += f" + PDF"
                        if on_log:
                            on_log(f"[{completed_count}/{total_tasks}] {msg}")
                        if on_progress:
                            on_progress(completed_count, total_tasks, os.path.basename(res["docx_path"]), "Completed")
                    else:
                        error_count += 1
                        err_msg = res.get("error", "Unknown error")
                        if on_log:
                            on_log(f"[{completed_count}/{total_tasks}] ❌ ERROR ({task_info['task_id']}): {err_msg}")
                        if on_progress:
                            on_progress(completed_count, total_tasks, task_info["task_id"], f"Error: {err_msg}")

                    results_summary.append(res)
                except Exception as ex:
                    error_count += 1
                    if on_log:
                        on_log(f"[{completed_count}/{total_tasks}] ❌ EXCEPTION: {ex}")

        elapsed_time = round(time.time() - start_time, 2)
        summary = {
            "total_tasks": total_tasks,
            "completed": completed_count,
            "successful": success_count,
            "errors": error_count,
            "elapsed_seconds": elapsed_time,
            "output_dir": output_dir,
            "cancelled": self.is_cancelled
        }

        if on_log:
            on_log(f"🎉 Batch process finished in {elapsed_time}s. Total: {total_tasks}, Succeeded: {success_count}, Failed: {error_count}.")

        if on_complete:
            on_complete(summary)

        return summary
