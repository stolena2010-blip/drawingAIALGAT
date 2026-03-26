"""
Email Connector Factory
=======================

בוחר בין שיטות שונות של התחברות למייל (COM, EWS, Graph API)
"""

from typing import Union, Optional, TYPE_CHECKING
from enum import Enum

from src.core.config import get_config

if TYPE_CHECKING:
    from src.services.email.shared_mailbox import SharedMailboxConnector
    from src.services.email.graph_helper import GraphAPIHelper
    from email_connector_ews import EWSEmailConnector


class EmailMethod(str, Enum):
    """שיטות התחברות למייל"""
    OUTLOOK_COM = "OUTLOOK_COM"
    EWS = "EWS"
    GRAPH_API = "GRAPH_API"


class EmailConnectorFactory:
    """
    Factory ליצירת email connectors
    בוחר בין שיטות שונות לפי הגדרה
    """
    
    @staticmethod
    def create(method: Optional[str] = None) -> Union['SharedMailboxConnector', 'GraphAPIHelper', 'EWSEmailConnector']:
        """
        יצור connector לפי שיטה
        
        Args:
            method: שיטה ('OUTLOOK_COM', 'EWS', 'GRAPH_API')
                    אם None, יקח מ-config
        
        Returns:
            Email connector object
        
        Raises:
            ValueError: אם שיטה לא נתמכת או חסרים credentials
        """
        cfg = get_config()
        
        if method is None:
            method = cfg.email.method
        
        method = method.upper()
        
        if method == EmailMethod.OUTLOOK_COM:
            from src.services.email.shared_mailbox import SharedMailboxConnector
            return SharedMailboxConnector(shared_mailbox=cfg.email.shared_mailbox)
        
        elif method == EmailMethod.EWS:
            from email_connector_ews import EWSEmailConnector
            return EWSEmailConnector(
                email_address=cfg.email.email_address or "",
                password=cfg.email.email_password or "",
                shared_mailbox=cfg.email.shared_mailbox,
                access_type=cfg.email.access_type
            )
        
        elif method == EmailMethod.GRAPH_API:
            if not cfg.email.graph_api or not cfg.email.graph_api.is_configured():
                raise ValueError("Graph API not configured. Set GRAPH_TENANT_ID, GRAPH_CLIENT_ID, GRAPH_CLIENT_SECRET")
            
            from src.services.email.graph_helper import GraphAPIHelper
            return GraphAPIHelper(
                tenant_id=cfg.email.graph_api.tenant_id,
                client_id=cfg.email.graph_api.client_id,
                client_secret=cfg.email.graph_api.client_secret,
                shared_mailbox=cfg.email.graph_api.shared_mailbox or cfg.email.shared_mailbox,
                use_config=False
            )
        
        else:
            raise ValueError(f"Unsupported email method: {method}")
    
    @staticmethod
    def get_available_methods() -> list:
        """קבל רשימת שיטות זמינות"""
        cfg = get_config()
        available = []
        
        # COM תמיד זמין ב-Windows
        import os
        if os.name == 'nt':
            available.append(EmailMethod.OUTLOOK_COM)
        
        # EWS זמין אם יש credentials
        if cfg.email.email_address and cfg.email.email_password:
            available.append(EmailMethod.EWS)
        
        # Graph API זמין אם יש credentials
        if cfg.email.graph_api and cfg.email.graph_api.is_configured():
            available.append(EmailMethod.GRAPH_API)
        
        return available
    
    @staticmethod
    def get_recommended_method() -> str:
        """קבל שיטה מומלצת"""
        available = EmailConnectorFactory.get_available_methods()
        
        # סדר העדפה
        for method in [EmailMethod.GRAPH_API, EmailMethod.EWS, EmailMethod.OUTLOOK_COM]:
            if method in available:
                return method
        
        raise ValueError("No email method available")


def create_email_connector(method: Optional[str] = None):
    """
    Factory function להשגת email connector
    
    Args:
        method: שיטה או None להשתמש בברירת מחדל
    
    Returns:
        Email connector object
    """
    return EmailConnectorFactory.create(method)


def test_all_methods() -> dict:
    """בדוק את כל השיטות הזמינות"""
    results = {}
    
    factory = EmailConnectorFactory()
    available = factory.get_available_methods()
    
    for method in available:
        try:
            print(f"\n🧪 Testing {method}...")
            connector = factory.create(method)
            
            # נסה התחברות
            if method == "OUTLOOK_COM":
                connector.connect()
                success = connector.is_connected
            elif method == "EWS":
                success = connector.connect()
            else:  # GRAPH_API
                success = connector.test_connection()
            
            results[method] = {
                "available": True,
                "working": success,
                "error": connector.last_error if not success else None
            }
            
            status = "✅" if success else "❌"
            print(f"   {status} {method}")
        
        except Exception as e:
            results[method] = {
                "available": True,
                "working": False,
                "error": str(e)
            }
            print(f"   ❌ {method}: {e}")
    
    return results
