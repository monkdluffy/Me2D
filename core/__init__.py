"""
Core engine package for SMAP AutoDoc Studio:
- doc_engine: Word DOCX processing (docxtpl & python-docx XML split-run sanitizer)
- pdf_engine: Cross-platform PDF conversion (LibreOffice / Word COM), watermarking, stamp overlay, PDF merge
- profile_manager: JSON and Excel profile serialization/deserialization
- batch_worker: Multi-threaded batch document replacement orchestrator
"""

from .doc_engine import DocumentEngine
from .pdf_engine import PDFEngine
from .profile_manager import ProfileManager
from .batch_worker import BatchProcessor

__all__ = ["DocumentEngine", "PDFEngine", "ProfileManager", "BatchProcessor"]
