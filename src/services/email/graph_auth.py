"""
Microsoft Graph API Authentication Module
==========================================

טיפול בהשגת ושמירת access tokens עבור Microsoft Graph API
"""

import os
import json
import time
import sys
from pathlib import Path
from typing import Optional, Dict, Any
from datetime import datetime, timedelta

try:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    if hasattr(sys.stderr, "reconfigure"):
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

try:
    from msal import PublicClientApplication, ConfidentialClientApplication
    HAS_MSAL = True
except ImportError:
    HAS_MSAL = False


class GraphAuthenticator:
    """
    מטפל בהשגת ושמירת access tokens עבור Graph API
    תומך ב-Client Credentials Flow (service principal)
    """
    
    def __init__(
        self,
        tenant_id: str,
        client_id: str,
        client_secret: str,
        authority_url: str = "https://login.microsoftonline.com",
        cache_file: Optional[str] = None,
        scopes: Optional[list] = None
    ):
        """
        אתחול מטפל ההשגה
        
        Args:
            tenant_id: Azure Tenant ID
            client_id: Application (Client) ID
            client_secret: Client Secret Value
            authority_url: Authority URL (default: Microsoft)
            cache_file: Optional path to cache token
            scopes: List of scopes (default: Graph API default)
        """
        if not HAS_MSAL:
            raise ImportError("msal is required. Install with: pip install msal")
        
        self.tenant_id = tenant_id
        self.client_id = client_id
        self.client_secret = client_secret
        self.authority_url = authority_url
        self.scopes = scopes or ["https://graph.microsoft.com/.default"]
        self.cache_file = cache_file or str(Path.home() / ".cache" / "graph_token.json")
        
        # בנה את ה-authority
        self.authority = f"{authority_url}/{tenant_id}"
        
        # אתחול MSAL client (Confidential Client עבור service principal)
        self.client = ConfidentialClientApplication(
            client_id=client_id,
            client_credential=client_secret,
            authority=self.authority
        )
        
        self.access_token: Optional[str] = None
        self.token_expires_at: Optional[float] = None
        self.last_error: Optional[str] = None
    
    def _ensure_cache_dir(self):
        """וודא שתיקיית ה-cache קיימת"""
        cache_path = Path(self.cache_file)
        cache_path.parent.mkdir(parents=True, exist_ok=True)
    
    def _load_cached_token(self) -> Optional[str]:
        """
        טען token מ-cache אם עדיין תקף
        
        Returns:
            Access token או None אם אין/לא תקף
        """
        try:
            if not os.path.exists(self.cache_file):
                return None
            
            with open(self.cache_file, 'r') as f:
                cache_data = json.load(f)
            
            token = cache_data.get("access_token")
            expires_at = cache_data.get("expires_at")
            
            if not token or not expires_at:
                return None
            
            # בדוק אם ה-token עדיין תקף (עם 5 דקות buffer)
            if time.time() < (expires_at - 300):
                return token
            
            return None
        except Exception as e:
            print(f"⚠ Failed to load token from cache: {e}")
            return None
    
    def _save_token_to_cache(self, token: str, expires_in: int):
        """
        שמור token ל-cache
        
        Args:
            token: Access token
            expires_in: זמן תפוגה בשניות
        """
        try:
            self._ensure_cache_dir()
            expires_at = time.time() + expires_in
            
            cache_data = {
                "access_token": token,
                "expires_at": expires_at,
                "saved_at": datetime.now().isoformat()
            }
            
            with open(self.cache_file, 'w') as f:
                json.dump(cache_data, f, indent=2)
        except Exception as e:
            print(f"⚠ Failed to save token to cache: {e}")
    
    def get_access_token(self, force_refresh: bool = False) -> Optional[str]:
        """
        קבל access token
        
        Args:
            force_refresh: אם True, תמיד קבל token חדש (תעלם מ-cache)
        
        Returns:
            Access token או None אם נכשל
        """
        # אם יש token בתוך-זיכרון ותקף, החזר אותו
        if not force_refresh and self.access_token and self.token_expires_at:
            if time.time() < (self.token_expires_at - 300):  # 5 דקות buffer
                return self.access_token
        
        # נסה טעינה מ-cache
        if not force_refresh:
            cached_token = self._load_cached_token()
            if cached_token:
                self.access_token = cached_token
                return cached_token
        
        # קבל token חדש מ-Azure
        try:
            print("🔐 Requesting access token for Graph API...")
            
            result = self.client.acquire_token_for_client(scopes=self.scopes)
            
            if "access_token" not in result:
                error_msg = result.get("error_description", result.get("error", "Unknown error"))
                self.last_error = error_msg
                print(f"❌ Failed to obtain token: {error_msg}")
                return None
            
            token = result["access_token"]
            expires_in = result.get("expires_in", 3600)  # default 1 hour
            
            # שמור בזיכרון
            self.access_token = token
            self.token_expires_at = time.time() + expires_in
            
            # שמור ב-cache
            self._save_token_to_cache(token, expires_in)
            
            # חשב זמן תפוגה קריא
            expiry_time = datetime.fromtimestamp(self.token_expires_at).strftime("%H:%M:%S")
            print(f"✅ Access token received (valid until {expiry_time})")
            
            return token
            
        except Exception as e:
            error_msg = str(e)
            self.last_error = error_msg
            print(f"❌ Failed to obtain access token: {error_msg}")
            return None
    
    def test_connection(self) -> bool:
        """
        בדוק אם ההשגה עובדת
        
        Returns:
            True אם הצליח להשיג token
        """
        token = self.get_access_token(force_refresh=True)
        return token is not None
    
    def diagnose(self) -> Dict[str, Any]:
        """
        אבחון לפי תושג
        
        Returns:
            Dict עם פרטי האבחון
        """
        result = {
            "tenant_id": self.tenant_id,
            "client_id": self.client_id,
            "authority": self.authority,
            "scopes": self.scopes,
            "cache_file": self.cache_file,
            "has_msal": HAS_MSAL,
            "token_status": None,
            "error": None
        }
        
        try:
            token = self.get_access_token(force_refresh=True)
            if token:
                result["token_status"] = "✅ קיבלנו access token בהצלחה"
                result["token_length"] = len(token)
                result["expires_at"] = datetime.fromtimestamp(self.token_expires_at).isoformat()
            else:
                result["token_status"] = f"❌ נכשל: {self.last_error}"
                result["error"] = self.last_error
        except Exception as e:
            result["token_status"] = f"❌ שגיאה: {str(e)}"
            result["error"] = str(e)
        
        return result


def create_authenticator(
    tenant_id: str,
    client_id: str,
    client_secret: str,
    **kwargs
) -> GraphAuthenticator:
    """
    Factory function ליצירת authenticator
    
    Args:
        tenant_id: Azure Tenant ID
        client_id: Application (Client) ID
        client_secret: Client Secret Value
        **kwargs: פרמטרים נוספים (authority_url, scopes, etc.)
    
    Returns:
        GraphAuthenticator instance
    """
    return GraphAuthenticator(
        tenant_id=tenant_id,
        client_id=client_id,
        client_secret=client_secret,
        **kwargs
    )
