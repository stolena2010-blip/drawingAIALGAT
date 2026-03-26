"""
Drawing data model
"""
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Optional, Dict, Any

from .enums import ConfidenceLevel, FileType, ProcessingStatus, ReviewStatus


@dataclass
class Drawing:
    """
    Represents a technical drawing file with extracted information
    
    Attributes:
        file_path: Full path to the drawing file
        file_name: Name of the file
        file_size_bytes: Size in bytes
        file_type: Detected file type
        modification_date: Last modification date
        
        # Extracted data - Stage 1 (Basic Info)
        customer_name: Customer/client name
        part_number: Part number (מק״ט)
        item_name: Item description
        drawing_number: Drawing number
        revision: Revision/version
        
        # Extracted data - Stage 2 (Processes)
        material: Material specification
        coating_processes: Coating/plating processes
        painting_processes: Painting specifications
        colors: Color specifications
        specifications: Technical specifications
        
        # Extracted data - Stage 3 (Notes)
        notes_full_text: Complete notes text
        process_summary_hebrew: Hebrew process summary
        
        # Extracted data - Stage 4 (Geometry)
        part_area: Calculated geometric area
        
        # Metadata
        confidence_level: Confidence in extraction
        needs_review: Review status
        processing_status: Current processing status
        extraction_date: When extraction was performed
        error_message: Error message if processing failed
        
        # Additional metadata
        has_parts_list: Whether drawing has parts list
        is_rafael_drawing: Whether this is a RAFAEL drawing
        ocr_text: OCR extracted text (if applicable)
        
        # Cost tracking
        extraction_cost_usd: Cost of extraction in USD
        execution_time_seconds: Time taken to process
    """
    
    # File information
    file_path: Path
    file_name: str
    file_size_bytes: int
    file_type: FileType
    modification_date: datetime
    
    # Stage 1: Basic Info
    customer_name: Optional[str] = None
    part_number: Optional[str] = None
    item_name: Optional[str] = None
    drawing_number: Optional[str] = None
    revision: Optional[str] = None
    
    # Stage 2: Processes
    material: Optional[str] = None
    coating_processes: Optional[str] = None
    painting_processes: Optional[str] = None
    colors: Optional[str] = None
    specifications: Optional[str] = None
    
    # Stage 3: Notes
    notes_full_text: Optional[str] = None
    process_summary_hebrew: Optional[str] = None
    
    # Stage 4: Geometry
    part_area: Optional[str] = None
    
    # Metadata
    confidence_level: ConfidenceLevel = ConfidenceLevel.UNKNOWN
    needs_review: ReviewStatus = ReviewStatus.NO_REVIEW_NEEDED
    processing_status: ProcessingStatus = ProcessingStatus.PENDING
    extraction_date: datetime = field(default_factory=datetime.now)
    error_message: Optional[str] = None
    
    # Additional info
    has_parts_list: bool = False
    is_rafael_drawing: bool = False
    ocr_text: Optional[str] = None
    
    # Cost/Performance
    extraction_cost_usd: float = 0.0
    execution_time_seconds: float = 0.0
    
    def __post_init__(self):
        """Post-initialization processing"""
        # Ensure file_path is Path object
        if not isinstance(self.file_path, Path):
            self.file_path = Path(self.file_path)
        
        # Auto-detect RAFAEL
        if self.customer_name:
            self.is_rafael_drawing = "RAFAEL" in self.customer_name.upper()
    
    @property
    def file_size_mb(self) -> float:
        """Get file size in MB"""
        return self.file_size_bytes / (1024 * 1024)
    
    @property
    def is_complete(self) -> bool:
        """Check if basic extraction is complete"""
        return all([
            self.customer_name,
            self.part_number,
            self.drawing_number
        ])
    
    @property
    def extraction_cost_ils(self) -> float:
        """Get extraction cost in ILS"""
        return self.extraction_cost_usd * 3.7
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary for serialization"""
        return {
            'file_name': self.file_name,
            'file_path': str(self.file_path),
            'file_size_mb': self.file_size_mb,
            'file_type': self.file_type.value,
            'modification_date': self.modification_date.isoformat(),
            
            # Extracted data
            'customer_name': self.customer_name,
            'part_number': self.part_number,
            'item_name': self.item_name,
            'drawing_number': self.drawing_number,
            'revision': self.revision,
            'material': self.material,
            'coating_processes': self.coating_processes,
            'painting_processes': self.painting_processes,
            'colors': self.colors,
            'specifications': self.specifications,
            'notes_full_text': self.notes_full_text,
            'process_summary_hebrew': self.process_summary_hebrew,
            'part_area': self.part_area,
            
            # Metadata
            'confidence_level': self.confidence_level.value,
            'needs_review': self.needs_review.value,
            'processing_status': self.processing_status.value,
            'extraction_date': self.extraction_date.isoformat(),
            'error_message': self.error_message,
            'has_parts_list': self.has_parts_list,
            'is_rafael_drawing': self.is_rafael_drawing,
            
            # Cost
            'extraction_cost_usd': self.extraction_cost_usd,
            'extraction_cost_ils': self.extraction_cost_ils,
            'execution_time_seconds': self.execution_time_seconds,
        }
    
    def __repr__(self) -> str:
        return (
            f"Drawing(file='{self.file_name}', "
            f"customer='{self.customer_name}', "
            f"part#='{self.part_number}', "
            f"status={self.processing_status.name})"
        )
