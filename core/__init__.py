"""
Mail Converter Core Package
"""

from .pst_extractor import PSTExtractor
from .eml_parser import EMLParser
from .mbox_extractor import MBOXExtractor, MboxExtractionResult
from .msg_parser import MSGParser, ParsedMSG
from .attachment_converter import AttachmentConverter
from .pdf_merger import PDFMerger
from .duplicate_detector import DuplicateDetector
from .conversion_pipeline import ConversionPipeline, InputType

# Email Tools
from .email_fingerprint import EmailFingerprint, FingerprintIndex, create_fingerprint
from .mailbox_writer import MailboxWriter, OutputFormat
from .mailbox_comparator import MailboxComparator, ComparisonConfig, ComparisonResult
from .mailbox_merger import MailboxMerger, MergeConfig, MergeResult
from .mailbox_deduplicator import MailboxDeduplicator, DedupeConfig, DedupeResult
from .mailbox_filter import MailboxFilter, FilterConfig, FilterResult

__all__ = [
    'PSTExtractor',
    'EMLParser',
    'MBOXExtractor',
    'MboxExtractionResult',
    'MSGParser',
    'ParsedMSG',
    'AttachmentConverter',
    'PDFMerger',
    'DuplicateDetector',
    'ConversionPipeline',
    'InputType',
    # Email Tools
    'EmailFingerprint',
    'FingerprintIndex',
    'create_fingerprint',
    'MailboxWriter',
    'OutputFormat',
    'MailboxComparator',
    'ComparisonConfig',
    'ComparisonResult',
    'MailboxMerger',
    'MergeConfig',
    'MergeResult',
    'MailboxDeduplicator',
    'DedupeConfig',
    'DedupeResult',
    'MailboxFilter',
    'FilterConfig',
    'FilterResult',
]
