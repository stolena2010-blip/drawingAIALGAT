"""
Core configuration module for DrawingAI Pro
"""
import os
from dataclasses import dataclass, field
from typing import Optional
from pathlib import Path

def normalize_azure_endpoint(endpoint: str) -> str:
    """Normalize Azure endpoint to resource base URL."""
    value = (endpoint or "").strip()
    if value.endswith("/openai/v1/"):
        value = value[:-11]
    elif value.endswith("/openai/v1"):
        value = value[:-10]
    return value.rstrip("/")


@dataclass
class AzureConfig:
    """Azure OpenAI configuration"""
    endpoint: str
    api_key: str
    api_version: str = "2024-08-01-preview"
    deployment: str = ""
    
    @classmethod
    def from_env(cls) -> 'AzureConfig':
        """Load Azure config from environment variables"""
        return cls(
            endpoint=normalize_azure_endpoint(os.getenv("AZURE_OPENAI_ENDPOINT", "")),
            api_key=os.getenv("AZURE_OPENAI_API_KEY", ""),
            api_version=os.getenv("AZURE_OPENAI_API_VERSION", "2024-08-01-preview"),
            deployment=os.getenv("AZURE_OPENAI_DEPLOYMENT", "")
        )


@dataclass
class FileLimits:
    """File size and dimension limits"""
    max_file_size_mb: int = 100
    warn_file_size_mb: int = 50
    max_image_dimension: int = 4096
    target_image_dimension: int = 2048
    warn_image_dimension: int = 3000


@dataclass
class PricingConfig:
    """API pricing configuration (USD per 1M tokens)"""
    input_price_per_1m: float = field(default_factory=lambda: float(os.getenv("AZURE_MODEL_INPUT_PRICE_PER_1M", "2.50")))
    output_price_per_1m: float = field(default_factory=lambda: float(os.getenv("AZURE_MODEL_OUTPUT_PRICE_PER_1M", "10.00")))
    usd_to_ils_rate: float = field(default_factory=lambda: float(os.getenv("USD_TO_ILS_RATE", "3.7")))


@dataclass
class ProcessingConfig:
    """Processing behavior configuration"""
    enable_retry: bool = True
    max_retry_attempts: int = 3
    enable_contrast_enhancement: bool = True
    enable_ocr: bool = True
    ocr_language: str = "eng+heb"


@dataclass
class GraphAPIConfig:
    """Microsoft Graph API configuration"""
    tenant_id: str
    client_id: str
    client_secret: str
    shared_mailbox: Optional[str] = None
    dest_dir: Optional[str] = None
    authority_url: str = "https://login.microsoftonline.com"
    graph_api_url: str = "https://graph.microsoft.com"
    scopes: list = field(default_factory=lambda: ["https://graph.microsoft.com/.default"])
    
    @classmethod
    def from_env(cls) -> 'GraphAPIConfig':
        """Load Graph API config from environment variables"""
        return cls(
            tenant_id=os.getenv("GRAPH_TENANT_ID", ""),
            client_id=os.getenv("GRAPH_CLIENT_ID", ""),
            client_secret=os.getenv("GRAPH_CLIENT_SECRET", ""),
            shared_mailbox=os.getenv("SHARED_MAILBOX_NAME", None),
            dest_dir=os.getenv("EMAIL_DEST_DIR", None),
            authority_url=os.getenv("GRAPH_AUTHORITY_URL", "https://login.microsoftonline.com"),
            graph_api_url=os.getenv("GRAPH_API_URL", "https://graph.microsoft.com")
        )
    
    def is_configured(self) -> bool:
        """Check if Graph API credentials are properly configured"""
        return bool(self.tenant_id and self.client_id and self.client_secret)


@dataclass
class EmailConfig:
    """Shared mailbox/email connection configuration"""
    method: str = "OUTLOOK_COM"  # OUTLOOK_COM, EWS, or GRAPH_API
    shared_mailbox: Optional[str] = None
    dest_dir: Optional[str] = None  # optional default target for saved content

    # EWS-specific (optional, only if method == 'EWS')
    email_address: Optional[str] = None
    email_password: Optional[str] = None
    ews_server: str = "outlook.office365.com"
    access_type: str = "DELEGATE"  # or IMPERSONATION
    
    # Graph API configuration (if method == 'GRAPH_API')
    graph_api: Optional[GraphAPIConfig] = None

    @classmethod
    def from_env(cls) -> 'EmailConfig':
        """Load email config from environment variables"""
        graph_cfg = GraphAPIConfig.from_env()
        
        return cls(
            method=os.getenv("EMAIL_METHOD", "OUTLOOK_COM").upper(),
            shared_mailbox=os.getenv("SHARED_MAILBOX_NAME", None),
            dest_dir=os.getenv("EMAIL_DEST_DIR", None),
            email_address=os.getenv("EMAIL_ADDRESS", None),
            email_password=os.getenv("EMAIL_PASSWORD", None),
            ews_server=os.getenv("EWS_SERVER", "outlook.office365.com"),
            access_type=os.getenv("EMAIL_ACCESS_TYPE", "DELEGATE").upper(),
            graph_api=graph_cfg if graph_cfg.is_configured() else None
        )


@dataclass
class Config:
    """Main application configuration"""
    azure: AzureConfig
    file_limits: FileLimits = field(default_factory=FileLimits)
    pricing: PricingConfig = field(default_factory=PricingConfig)
    processing: ProcessingConfig = field(default_factory=ProcessingConfig)
    email: EmailConfig = field(default_factory=EmailConfig.from_env)
    output_folder: Optional[Path] = None
    
    # Supported file extensions
    DRAWING_EXTENSIONS = {".pdf", ".png", ".jpg", ".jpeg", ".tif", ".tiff"}
    
    @classmethod
    def from_env(cls) -> 'Config':
        """Load complete configuration from environment"""
        return cls(
            azure=AzureConfig.from_env(),
            file_limits=FileLimits(),
            pricing=PricingConfig(),
            processing=ProcessingConfig(),
            email=EmailConfig.from_env()
        )
    
    def validate(self) -> bool:
        """Validate configuration"""
        if not self.azure.endpoint or not self.azure.api_key:
            raise ValueError("Azure endpoint and API key are required")
        
        if self.file_limits.max_file_size_mb <= 0:
            raise ValueError("max_file_size_mb must be positive")

        # Email config is optional - no strict validation needed here
        return True


# Global config instance (lazy loaded)
_config: Optional[Config] = None


def get_config() -> Config:
    """Get global configuration instance"""
    global _config
    if _config is None:
        _config = Config.from_env()
        _config.validate()
    return _config


def reset_config() -> None:
    """Reset global config (useful for testing)"""
    global _config
    _config = None
