import os
import sys
import shutil
import subprocess
import logging
import tempfile
from typing import List, Tuple, Optional
import fitz  # PyMuPDF

from config import find_libreoffice_path

logger = logging.getLogger("pdf_engine")

class PDFEngine:
    """Cross-Platform PDF Converter and PyMuPDF Post-Processing Engine."""

    @staticmethod
    def convert_docx_to_pdf(docx_path: str, output_pdf_path: str, custom_libreoffice_path: str = "") -> Tuple[bool, str]:
        """Converts a DOCX file to PDF using LibreOffice Headless or MS Word COM automation (Windows).
        
        Args:
            docx_path: Path to input .docx file.
            output_pdf_path: Destination path for generated .pdf.
            custom_libreoffice_path: Optional explicit path to soffice executable.
            
        Returns:
            Tuple[bool, str]: (Success status, Error message or output path)
        """
        if not os.path.exists(docx_path):
            return False, f"Input DOCX file not found: {docx_path}"

        abs_docx = os.path.abspath(docx_path)
        abs_output_pdf = os.path.abspath(output_pdf_path)
        out_dir = os.path.dirname(abs_output_pdf)
        os.makedirs(out_dir, exist_ok=True)

        # 1. Detect LibreOffice binary
        lo_bin = custom_libreoffice_path if (custom_libreoffice_path and os.path.exists(custom_libreoffice_path)) else find_libreoffice_path()

        if lo_bin:
            logger.info(f"Using LibreOffice for PDF conversion: {lo_bin}")
            try:
                with tempfile.TemporaryDirectory() as tmp_dir:
                    env_profile = f"-env:UserInstallation=file:///{tmp_dir.replace(chr(92), '/')}/lo_profile"
                    cmd = [
                        lo_bin,
                        env_profile,
                        "--headless",
                        "--nologo",
                        "--nofirststartwizard",
                        "--convert-to", "pdf",
                        "--outdir", tmp_dir,
                        abs_docx
                    ]
                    # Run headless conversion
                    res = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, timeout=120)
                    if res.returncode == 0:
                        base_name = os.path.splitext(os.path.basename(abs_docx))[0]
                        generated_pdf = os.path.join(tmp_dir, f"{base_name}.pdf")
                        if os.path.exists(generated_pdf):
                            if os.path.exists(abs_output_pdf):
                                os.remove(abs_output_pdf)
                            shutil.move(generated_pdf, abs_output_pdf)
                            return True, abs_output_pdf
                    logger.warning(f"LibreOffice process finished with code {res.returncode}. Stderr: {res.stderr}")
            except Exception as lo_err:
                logger.warning(f"LibreOffice conversion error: {lo_err}")

        # 2. Windows Fallback: MS Word COM Automation
        if sys.platform == "win32":
            logger.info("Attempting MS Word COM automation fallback for PDF conversion...")
            try:
                import win32com.client
                import pythoncom
                import time
                
                # Initialize COM for the current thread
                pythoncom.CoInitialize()
                try:
                    max_retries = 3
                    for attempt in range(max_retries):
                        word = None
                        try:
                            word = win32com.client.DispatchEx("Word.Application")
                            word.Visible = False
                            word.DisplayAlerts = 0
                            
                            # Open the DOCX and save as PDF
                            doc = word.Documents.Open(abs_docx, ReadOnly=True)
                            doc.SaveAs(abs_output_pdf, FileFormat=17)
                            doc.Close(0)
                            return True, abs_output_pdf
                        except Exception as e:
                            logger.error(f"Error during Word document processing (Attempt {attempt+1}/{max_retries}): {e}")
                            if attempt == max_retries - 1:
                                return False, str(e)
                            time.sleep(1.5)
                        finally:
                            if word:
                                try:
                                    word.Quit()
                                except:
                                    pass
                finally:
                    # Uninitialize COM
                    pythoncom.CoUninitialize()
            except Exception as e:
                logger.error(f"Failed to execute MS Word COM automation: {e}")
                return False, str(e)

        return False, "PDF conversion failed: Neither LibreOffice nor MS Word COM automation succeeded."

    @staticmethod
    def apply_watermark(
        pdf_path: str,
        output_pdf_path: str,
        watermark_text: str = "DRAFT",
        opacity: float = 0.25,
        angle: float = 45.0,
        font_size: float = 48.0,
        color_rgb: Tuple[float, float, float] = (0.7, 0.7, 0.7)
    ) -> Tuple[bool, str]:
        """Applies a semi-transparent diagonal text watermark across all pages of a PDF using PyMuPDF.
        
        Args:
            pdf_path: Source PDF file path.
            output_pdf_path: Output PDF file path.
            watermark_text: Watermark text string (e.g. "DRAFT", "CONFIDENTIAL", "SALINAN RESMI").
            opacity: Opacity from 0.0 to 1.0.
            angle: Rotation angle in degrees.
            font_size: Font size in points.
            color_rgb: Text color RGB tuple (0.0 to 1.0).
            
        Returns:
            Tuple[bool, str]: (Success status, Error message or output path)
        """
        if not os.path.exists(pdf_path):
            return False, f"Source PDF file not found: {pdf_path}"

        try:
            doc = fitz.open(pdf_path)
            for page in doc:
                rect = page.rect
                
                # Calculate text length to truly center it
                tw = fitz.get_text_length(watermark_text, fontname="helv", fontsize=font_size)
                
                center_point = fitz.Point((rect.width - tw) / 2, rect.height / 2)
                
                # Insert watermark text centered on page with matrix rotation
                mat = fitz.Matrix(angle)
                
                # Use a slightly darker color to make it more visible if it was too faint
                page.insert_text(
                    center_point,
                    watermark_text,
                    fontsize=font_size,
                    fontname="helv",
                    color=(1.0, 0.2, 0.2), # Redish color so it's very visible
                    fill_opacity=opacity,
                    morph=(center_point, mat)
                )

            os.makedirs(os.path.dirname(os.path.abspath(output_pdf_path)), exist_ok=True)
            doc.save(output_pdf_path)
            doc.close()
            return True, output_pdf_path
        except Exception as e:
            logger.error(f"Failed to apply watermark to PDF '{pdf_path}': {e}")
            return False, str(e)

    @staticmethod
    def apply_stamp_or_signature(
        pdf_path: str,
        output_pdf_path: str,
        image_path: str,
        page_num: int = -1,  # -1 means last page, 0 means all pages, >=1 means 1-indexed page
        x_pt: float = 400.0,
        y_pt: float = 650.0,
        width_pt: float = 120.0,
        height_pt: float = 75.0,
        position_preset: Optional[str] = None  # e.g. "bottom-right", "bottom-left", "top-right"
    ) -> Tuple[bool, str]:
        """Inserts an image stamp/signature onto specified PDF page(s) at coordinates or quadrant preset.
        
        Args:
            pdf_path: Source PDF file.
            output_pdf_path: Output PDF file.
            image_path: Image file (PNG/JPG) to stamp.
            page_num: Target page (-1 for last page, 0 for all, 1..N for specific page).
            x_pt: X coordinate in points.
            y_pt: Y coordinate in points.
            width_pt: Image stamp width in points.
            height_pt: Image stamp height in points.
            position_preset: Optional preset position ("bottom-right", "bottom-left", "top-right").
            
        Returns:
            Tuple[bool, str]: (Success status, Error message or output path)
        """
        if not os.path.exists(pdf_path):
            return False, f"Source PDF file not found: {pdf_path}"
        if not os.path.exists(image_path):
            return False, f"Stamp image file not found: {image_path}"

        try:
            doc = fitz.open(pdf_path)
            total_pages = len(doc)
            target_indices = []

            if page_num == 0:
                target_indices = list(range(total_pages))
            elif page_num == -1:
                target_indices = [total_pages - 1]
            else:
                idx = page_num - 1
                if 0 <= idx < total_pages:
                    target_indices = [idx]
                else:
                    target_indices = [total_pages - 1]

            for idx in target_indices:
                page = doc[idx]
                rect = page.rect

                # Calculate placement rect based on preset if specified
                final_x = x_pt
                final_y = y_pt

                if position_preset == "bottom-right":
                    final_x = rect.width - width_pt - 36
                    final_y = rect.height - height_pt - 36
                elif position_preset == "bottom-left":
                    final_x = 36
                    final_y = rect.height - height_pt - 36
                elif position_preset == "top-right":
                    final_x = rect.width - width_pt - 36
                    final_y = 36

                stamp_rect = fitz.Rect(final_x, final_y, final_x + width_pt, final_y + height_pt)
                page.insert_image(stamp_rect, filename=image_path)

            os.makedirs(os.path.dirname(os.path.abspath(output_pdf_path)), exist_ok=True)
            doc.save(output_pdf_path)
            doc.close()
            return True, output_pdf_path
        except Exception as e:
            logger.error(f"Failed to apply stamp/signature to PDF '{pdf_path}': {e}")
            return False, str(e)

    @staticmethod
    def merge_pdfs(pdf_list: List[str], output_pdf_path: str) -> Tuple[bool, str]:
        """Combines multiple PDF files into one consolidated PDF file.
        
        Args:
            pdf_list: List of PDF file paths to merge in order.
            output_pdf_path: Destination path for consolidated PDF.
            
        Returns:
            Tuple[bool, str]: (Success status, Error message or output path)
        """
        valid_pdfs = [p for p in pdf_list if os.path.exists(p)]
        if not valid_pdfs:
            return False, "No valid PDF files provided for merging."

        try:
            merged_doc = fitz.open()
            for pdf_path in valid_pdfs:
                sub_doc = fitz.open(pdf_path)
                merged_doc.insert_pdf(sub_doc)
                sub_doc.close()

            os.makedirs(os.path.dirname(os.path.abspath(output_pdf_path)), exist_ok=True)
            merged_doc.save(output_pdf_path)
            merged_doc.close()
            return True, output_pdf_path
        except Exception as e:
            logger.error(f"Failed to merge PDFs: {e}")
            return False, str(e)

    @staticmethod
    def get_pdf_info(pdf_path: str) -> dict:
        """Retrieves metadata for a PDF file including size, page count, and modified date.
        
        Returns:
            dict: {"size_kb": float, "pages": int, "modified_date": str}
        """
        import datetime
        info = {"size_kb": 0.0, "pages": 0, "modified_date": ""}
        if not os.path.exists(pdf_path):
            return info
        try:
            stat = os.stat(pdf_path)
            info["size_kb"] = round(stat.st_size / 1024, 2)
            info["modified_date"] = datetime.datetime.fromtimestamp(stat.st_mtime).strftime('%Y-%m-%d %H:%M:%S')
            
            doc = fitz.open(pdf_path)
            info["pages"] = len(doc)
            doc.close()
        except Exception as e:
            logger.error(f"Error getting PDF info for {pdf_path}: {e}")
        return info

    @staticmethod
    def split_pdf(pdf_path: str, output_dir: str, mode: str = "all", selected_pages: List[int] = None) -> Tuple[bool, str]:
        """Splits a PDF into multiple single-page PDFs.
        
        Args:
            pdf_path: Source PDF file.
            output_dir: Directory to save split pages.
            mode: "all" (split every page) or "selected" (split only selected pages).
            selected_pages: List of 0-indexed pages to extract if mode=="selected".
            
        Returns:
            Tuple[bool, str]: (Success status, Error message or output directory)
        """
        if not os.path.exists(pdf_path):
            return False, f"Source PDF file not found: {pdf_path}"
            
        try:
            os.makedirs(output_dir, exist_ok=True)
            doc = fitz.open(pdf_path)
            base_name = os.path.splitext(os.path.basename(pdf_path))[0]
            
            total_pages = len(doc)
            target_pages = list(range(total_pages)) if mode == "all" else (selected_pages or [])
            
            extracted_count = 0
            out_msg = ""
            if mode == "selected" and target_pages:
                new_doc = fitz.open()
                valid_pages = [idx for idx in target_pages if 0 <= idx < total_pages]
                for idx in valid_pages:
                    new_doc.insert_pdf(doc, from_page=idx, to_page=idx)
                if valid_pages:
                    out_path = os.path.join(output_dir, f"{base_name}_extracted.pdf")
                    new_doc.save(out_path)
                    extracted_count = len(valid_pages)
                    out_msg = f"Successfully extracted {extracted_count} pages to {out_path}"
                else:
                    out_msg = "No valid pages found to extract."
                new_doc.close()
            else:
                for idx in target_pages:
                    if 0 <= idx < total_pages:
                        new_doc = fitz.open()
                        new_doc.insert_pdf(doc, from_page=idx, to_page=idx)
                        out_path = os.path.join(output_dir, f"{base_name}_page_{idx+1}.pdf")
                        new_doc.save(out_path)
                        new_doc.close()
                        extracted_count += 1
                out_msg = f"Successfully split {extracted_count} pages to {output_dir}"
                    
            doc.close()
            return True, out_msg
        except Exception as e:
            logger.error(f"Failed to split PDF: {e}")
            return False, str(e)

