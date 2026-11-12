# VoterMail Engine — High-Volume Word & Excel DPT Mail-Merge Automation
# Created by Danang Satria (https://github.com/Smurffahh/votermail-engine)

__author__ = "Danang Satria"

from .security import (
    sanitize_filename,
    validate_safe_output_path,
    sanitize_cell_text,
    safe_delete_file,
    clean_directory,
)
from .word_preparer import (
    WordCheckResult,
    WordSnapshot,
    check_and_prepare_word_template,
    compute_word_snapshot,
    analyze_word_diff,
    get_file_fingerprint,
)
from .excel_scanner import (
    VoterData,
    ScanIssue,
    SheetScanSummary,
    PreFlightScanResult,
    ExcelSnapshot,
    scan_and_validate_excel,
    compute_excel_snapshot,
    analyze_excel_diff,
    get_excel_file_fingerprint,
    transform_gender,
    transform_age,
    transform_address,
)
from .mail_merger import merge_voters_to_pdf_via_word_com, generate_preview_via_word
from .pdf_verifier import (
    VerificationAuditItem,
    VerificationResult,
    verify_kadus_pdf,
    export_audit_excel,
)

__all__ = [
    "sanitize_filename",
    "validate_safe_output_path",
    "sanitize_cell_text",
    "safe_delete_file",
    "clean_directory",
    "WordCheckResult",
    "check_and_prepare_word_template",
    "VoterData",
    "ScanIssue",
    "SheetScanSummary",
    "PreFlightScanResult",
    "scan_and_validate_excel",
    "transform_gender",
    "transform_age",
    "transform_address",
    "merge_voters_to_pdf_via_word_com",
    "generate_preview_via_word",
    "VerificationAuditItem",
    "VerificationResult",
    "verify_kadus_pdf",
    "export_audit_excel",
]
