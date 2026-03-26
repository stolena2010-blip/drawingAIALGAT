"""
Enums for DrawingAI Pro
"""
from enum import Enum, auto


class StageType(Enum):
    """Extraction stage types"""
    CLASSIFICATION = 0  # File type classification
    BASIC_INFO = 1      # Customer, Part#, Drawing#, Revision
    PROCESSES = 2       # Material, Coating, Painting, Specs
    NOTES = 3           # Full notes extraction
    GEOMETRY = 4        # Geometric area calculation


class ConfidenceLevel(Enum):
    """Confidence level in extraction results"""
    FULL = "full"       # Exact match - both part# and drawing# in filename
    HIGH = "high"       # Part# in filename
    MEDIUM = "medium"   # Only drawing# in filename
    LOW = "low"         # Neither in filename
    UNKNOWN = "unknown" # Not yet determined


class FileType(Enum):
    """Detected file types"""
    TECHNICAL_DRAWING = "technical_drawing"
    DOCUMENT = "document"
    IMAGE = "image"
    SPREADSHEET = "spreadsheet"
    UNKNOWN = "unknown"


class ProcessingStatus(Enum):
    """File processing status"""
    PENDING = auto()
    IN_PROGRESS = auto()
    COMPLETED = auto()
    FAILED = auto()
    SKIPPED = auto()


class ReviewStatus(Enum):
    """Review requirement status"""
    NO_REVIEW_NEEDED = "no_review"
    NEEDS_CHECK = "needs_check"          # Minor issues
    NEEDS_REVIEW = "needs_review"        # Significant issues
    PROBLEMATIC = "problematic"          # Major problems
