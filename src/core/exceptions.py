"""
Custom exceptions for DrawingAI Pro
"""


class DrawingAIException(Exception):
    """Base exception for all DrawingAI errors"""
    pass


class ConfigurationError(DrawingAIException):
    """Configuration-related errors"""
    pass


class FileValidationError(DrawingAIException):
    """File validation errors"""
    pass


class FileSizeTooLargeError(FileValidationError):
    """File exceeds maximum allowed size"""
    def __init__(self, file_path: str, size_mb: float, max_size_mb: int):
        self.file_path = file_path
        self.size_mb = size_mb
        self.max_size_mb = max_size_mb
        super().__init__(
            f"File {file_path} is {size_mb:.1f}MB, "
            f"exceeds maximum {max_size_mb}MB"
        )


class ImageProcessingError(DrawingAIException):
    """Image processing errors"""
    pass


class OCRError(ImageProcessingError):
    """OCR processing errors"""
    pass


class AIServiceError(DrawingAIException):
    """AI service (Azure OpenAI) errors"""
    pass


class APILimitExceededError(AIServiceError):
    """API rate limit or quota exceeded"""
    pass


class ExtractionError(DrawingAIException):
    """Data extraction errors"""
    pass


class InvalidResponseError(ExtractionError):
    """Invalid or unparseable response from AI"""
    pass


class RepositoryError(DrawingAIException):
    """Data persistence errors"""
    pass
