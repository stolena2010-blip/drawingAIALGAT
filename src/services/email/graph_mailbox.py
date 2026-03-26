"""
Microsoft Graph API Email Connector
====================================

תחזוקת קבלת מיילים מתיבה משותפת דרך Microsoft Graph API
תומך בהורדת מיילים, מצורפים, וניהול תיקיות
"""

import os
import re
import requests
import json
import sys
from pathlib import Path
from typing import List, Dict, Optional, Any, Callable
from datetime import datetime, timedelta
from urllib.parse import quote

try:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    if hasattr(sys.stderr, "reconfigure"):
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

from src.utils.logger import get_logger
from .graph_auth import GraphAuthenticator

logger = get_logger(__name__)


class GraphMailboxConnector:
    """
    מחבר תיבת משתמש דרך Microsoft Graph API
    משמש לגישה לתיבות משותפות, הורדת מיילים, ניהול תיקיות
    """
    
    def __init__(
        self,
        authenticator: GraphAuthenticator,
        shared_mailbox: str,
        graph_api_url: str = "https://graph.microsoft.com/v1.0"
    ) -> None:
        """
        אתחול מחבר
        
        Args:
            authenticator: GraphAuthenticator instance (עם access token)
            shared_mailbox: כתובת התיבה המשותפת
            graph_api_url: Base URL ל-Graph API
        """
        self.authenticator = authenticator
        self.shared_mailbox = shared_mailbox
        self.graph_api_url = graph_api_url
        self.is_connected = False
        self.last_error: Optional[str] = None
        self.folders_cache: Dict[str, str] = {}  # folder_name -> folder_id
    
    def _get_headers(self) -> Dict[str, str]:
        """קבל headers להרשאה"""
        token = self.authenticator.get_access_token()
        if not token:
            raise RuntimeError("Failed to get access token")
        
        return {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json"
        }
    
    def _make_request(
        self,
        method: str,
        endpoint: str,
        data: Optional[Dict] = None,
        params: Optional[Dict] = None
    ) -> Optional[Dict]:
        """
        בצע request ל-Graph API
        
        Args:
            method: GET, POST, DELETE, etc.
            endpoint: API endpoint
            data: JSON body (for POST/PATCH)
            params: Query parameters
        
        Returns:
            Response JSON או None אם נכשל
        """
        try:
            if str(endpoint).startswith("http://") or str(endpoint).startswith("https://"):
                url = str(endpoint)
            else:
                url = f"{self.graph_api_url}{endpoint}"
            headers = self._get_headers()
            
            if method.upper() == "GET":
                response = requests.get(url, headers=headers, params=params)
            elif method.upper() == "POST":
                response = requests.post(url, headers=headers, json=data, params=params)
            elif method.upper() == "PATCH":
                response = requests.patch(url, headers=headers, json=data, params=params)
            elif method.upper() == "DELETE":
                response = requests.delete(url, headers=headers, params=params)
            else:
                raise ValueError(f"Unsupported method: {method}")
            
            # בדוק status code
            if response.status_code >= 400:
                error_msg = response.text
                self.last_error = f"HTTP {response.status_code}: {error_msg}"
                return None
            
            # אם זה 200-299, החזר עדות
            if response.status_code >= 200 and response.status_code < 300:
                # 204 No Content או 202 Accepted - אין body
                if response.status_code in [202, 204]:
                    return {"status": "success"}
                
                # יש body - נסה לפרסר כ-JSON
                try:
                    return response.json()
                except Exception as e:
                    # אם לא JSON, החזר סטטוס הצלחה
                    logger.info(f"[GRAPH] Response not JSON (status {response.status_code}): {e}")
                    return {"status": "success"}
            
            return None

        except Exception as e:
            error_msg = str(e)
            self.last_error = error_msg
            # גם נסה לקרוא מתוך response אם יש
            if hasattr(e, 'response') and e.response is not None:
                try:
                    error_details = e.response.json()
                    self.last_error = f"{error_msg} | Details: {error_details}"
                except Exception as parse_err:
                    logger.error(f"[GRAPH] Could not parse error response: {parse_err}")
            return None

    def _get_all_pages(
        self,
        endpoint: str,
        params: Optional[Dict] = None,
        max_pages: int = 200
    ) -> List[Dict[str, Any]]:
        """קבל את כל הדפים מ-Graph endpoint שמחזיר value + @odata.nextLink"""
        items: List[Dict[str, Any]] = []
        next_endpoint: Optional[str] = endpoint
        next_params = params.copy() if isinstance(params, dict) else None
        pages = 0

        while next_endpoint and pages < max_pages:
            result = self._make_request("GET", next_endpoint, params=next_params)
            if not result:
                break

            values = result.get("value", [])
            if isinstance(values, list):
                items.extend(values)

            next_link = result.get("@odata.nextLink")
            next_endpoint = next_link if isinstance(next_link, str) and next_link.strip() else None
            next_params = None
            pages += 1

        return items

    def _list_child_folders_all(self, parent_id: str) -> List[Dict[str, Any]]:
        endpoint = f"/users/{quote(self.shared_mailbox)}/mailFolders/{parent_id}/childFolders"
        return self._get_all_pages(
            endpoint,
            params={
                "$top": 200,
                "includeHiddenFolders": "true",
                "$select": "id,displayName,parentFolderId,childFolderCount,totalItemCount,unreadItemCount"
            }
        )
    
    def connect(self) -> bool:
        """
        בדוק חיבור וקבל מידע על התיבה
        
        Returns:
            True אם הצליח
        """
        try:
            # בדוק גישה לתיבה המשותפת
            endpoint = f"/users/{quote(self.shared_mailbox)}"
            result = self._make_request("GET", endpoint)
            
            if result:
                self.is_connected = True
                self.last_error = None
                display_name = result.get("displayName", self.shared_mailbox)
                logger.info(f"✅ Connected successfully to mailbox: {display_name}")
                return True
            else:
                self.is_connected = False
                logger.error(f"❌ Connection failed: {self.last_error}")
                return False
        except Exception as e:
            self.last_error = str(e)
            self.is_connected = False
            return False
    
    def diagnose(self) -> Dict[str, Any]:
        """אבחון חיבור"""
        result = {
            "shared_mailbox": self.shared_mailbox,
            "graph_api_url": self.graph_api_url,
            "is_connected": False,
            "mailbox_info": None,
            "error": None
        }
        
        try:
            # בדוק גישה לתיבה
            endpoint = f"/users/{quote(self.shared_mailbox)}"
            mailbox = self._make_request("GET", endpoint)
            
            if mailbox:
                result["is_connected"] = True
                result["mailbox_info"] = {
                    "displayName": mailbox.get("displayName"),
                    "userPrincipalName": mailbox.get("userPrincipalName"),
                    "id": mailbox.get("id")
                }
            else:
                result["error"] = self.last_error
        except Exception as e:
            result["error"] = str(e)
        
        return result
    
    # Well-known folder names supported by Graph API (language-independent)
    WELL_KNOWN_FOLDERS = {
        "inbox", "drafts", "sentitems", "deleteditems",
        "junkemail", "archive", "outbox",
    }
    # Map user-friendly aliases → Graph API well-known names
    _FOLDER_ALIASES = {
        "inbox": "Inbox",
        "drafts": "Drafts",
        "sent items": "SentItems",
        "sentitems": "SentItems",
        "deleted items": "DeletedItems",
        "deleteditems": "DeletedItems",
        "junk email": "JunkEmail",
        "junkemail": "JunkEmail",
        "archive": "Archive",
        "outbox": "Outbox",
    }

    def get_well_known_folder(self, name: str) -> Optional[Dict[str, Any]]:
        """
        Get a well-known folder by Graph API name (e.g. 'Inbox').
        Works regardless of display language (Hebrew, English, etc.).
        """
        wkn = self._FOLDER_ALIASES.get(name.strip().lower())
        if not wkn:
            return None
        try:
            endpoint = f"/users/{quote(self.shared_mailbox)}/mailFolders/{wkn}"
            result = self._make_request("GET", endpoint, params={
                "$select": "id,displayName,parentFolderId,childFolderCount,totalItemCount,unreadItemCount"
            })
            if result and result.get("id"):
                return result
            return None
        except Exception as e:
            logger.debug(f"Well-known folder '{name}' not found: {e}")
            return None

    def list_folders(self, show_inbox_only: bool = False) -> List[Dict[str, str]]:
        """
        רשימת תיקיות בתיבה
        
        Args:
            show_inbox_only: אם True, הצג רק תת-תיקיות של Inbox
        
        Returns:
            רשימת dict עם 'id' ו-'displayName'
        """
        try:
            endpoint = f"/users/{quote(self.shared_mailbox)}/mailFolders"
            folders = self._get_all_pages(
                endpoint,
                params={
                    "$top": 200,
                    "includeHiddenFolders": "true",
                    "$select": "id,displayName,parentFolderId,childFolderCount,totalItemCount,unreadItemCount"
                }
            )
            
            # אם צריך Inbox בלבד
            if show_inbox_only:
                inbox = next((f for f in folders if str(f.get("displayName", "")).lower() == "inbox"), None)
                # Fallback: use well-known name (works for Hebrew "תיבת דואר נכנס" etc.)
                if not inbox:
                    inbox = self.get_well_known_folder("Inbox")
                if inbox:
                    # קבל את כל תת-התיקיות של Inbox (רקורסיבי - כל העומק)
                    inbox_id = inbox.get("id")
                    folders = []
                    visited_ids = set()
                    queue = [inbox_id] if inbox_id else []

                    while queue:
                        parent_id = queue.pop(0)
                        endpoint = f"/users/{quote(self.shared_mailbox)}/mailFolders/{parent_id}/childFolders"
                        children = self._get_all_pages(
                            endpoint,
                            params={
                                "$top": 200,
                                "includeHiddenFolders": "true",
                                "$select": "id,displayName,parentFolderId,childFolderCount,totalItemCount,unreadItemCount"
                            }
                        )

                        for child in children:
                            child_id = child.get("id")
                            if not child_id or child_id in visited_ids:
                                continue
                            visited_ids.add(child_id)
                            folders.append(child)
                            queue.append(child_id)
                else:
                    folders = []
            
            # עדכן cache
            for folder in folders:
                self.folders_cache[folder.get("displayName")] = folder.get("id")
            
            return folders
        
        except Exception as e:
            logger.error(f"❌ Failed to retrieve folders: {e}")
            return []

    def list_folders_recursive(self) -> List[Dict[str, Any]]:
        """רשימת כל התיקיות בתיבה (כולל תתי-תיקיות בכל עומק) עם path מלא."""
        try:
            roots = self.list_folders(show_inbox_only=False)
            if not roots:
                return []

            result: List[Dict[str, Any]] = []
            visited_ids = set()
            queue: List[Dict[str, Any]] = []

            for root in roots:
                folder_id = root.get("id")
                if not folder_id or folder_id in visited_ids:
                    continue
                visited_ids.add(folder_id)
                root_path = str(root.get("displayName", "")).strip()
                item = dict(root)
                item["path"] = root_path
                result.append(item)
                queue.append(item)

            while queue:
                parent = queue.pop(0)
                parent_id = parent.get("id")
                if not parent_id:
                    continue

                children = self._list_child_folders_all(parent_id)
                for child in children:
                    child_id = child.get("id")
                    if not child_id or child_id in visited_ids:
                        continue
                    visited_ids.add(child_id)

                    child_name = str(child.get("displayName", "")).strip()
                    parent_path = str(parent.get("path", parent.get("displayName", ""))).strip()
                    child_path = f"{parent_path}/{child_name}" if parent_path else child_name

                    child_item = dict(child)
                    child_item["path"] = child_path
                    result.append(child_item)
                    queue.append(child_item)

            return result
        except Exception as e:
            logger.error(f"❌ Failed to retrieve recursive folders: {e}")
            return []
    
    def list_messages(
        self,
        folder_id: str,
        limit: int = 100,
        received_after: Optional[datetime] = None
    ) -> List[Dict[str, Any]]:
        """
        קבל הודעות מתיקייה
        
        Args:
            folder_id: ID של התיקייה
            limit: מספר הודעות מקסימלי
            received_after: סנן הודעות מתאריך זה ואילך
        
        Returns:
            רשימת הודעות
        """
        try:
            # בנה filter אם צריך
            filters = []
            if received_after:
                # Convert to UTC for Graph API filter
                if received_after.tzinfo is None:
                    # Naive datetime = local Israel time → convert to UTC
                    try:
                        from zoneinfo import ZoneInfo
                    except ImportError:
                        from backports.zoneinfo import ZoneInfo
                    local_tz = ZoneInfo("Asia/Jerusalem")
                    utc_dt = received_after.replace(tzinfo=local_tz).astimezone(ZoneInfo("UTC"))
                else:
                    # Already timezone-aware → convert to UTC
                    from datetime import timezone as _tz
                    utc_dt = received_after.astimezone(_tz.utc)
                date_str = utc_dt.strftime("%Y-%m-%dT%H:%M:%S") + "Z"
                filters.append(f"receivedDateTime ge {date_str}")
            
            params = {
                "$top": limit,
                "$select": "id,subject,from,receivedDateTime,hasAttachments,bodyPreview,webLink,categories",
                "$expand": "attachments($select=id)"
            }
            
            if filters:
                params["$filter"] = " and ".join(filters)
            
            endpoint = f"/users/{quote(self.shared_mailbox)}/mailFolders/{folder_id}/messages"
            result = self._make_request("GET", endpoint, params=params)
            
            if result:
                return result.get("value", [])
            return []
        
        except Exception as e:
            logger.error(f"❌ Failed to retrieve messages: {e}")
            return []
    
    def get_message_details(self, message_id: str) -> Optional[Dict[str, Any]]:
        """קבל פרטי הודעה מלאים כולל גוף"""
        try:
            endpoint = f"/users/{quote(self.shared_mailbox)}/messages/{message_id}"
            result = self._make_request("GET", endpoint)
            return result
        except Exception as e:
            logger.error(f"❌ Failed to retrieve message details: {e}")
            return None
    
    def get_attachments(self, message_id: str) -> List[Dict[str, Any]]:
        """
        קבל רשימת מצורפים של הודעה (רק קבצים מצורפים אמיתיים, לא inline images)
        """
        try:
            endpoint = f"/users/{quote(self.shared_mailbox)}/messages/{message_id}/attachments"
            result = self._make_request("GET", endpoint)
            
            if result:
                all_attachments = result.get("value", [])
                # סנן רק inline images/signatures - הורד את כל השאר
                real_attachments = []
                for att in all_attachments:
                    is_inline = att.get("isInline", False)
                    content_id = att.get("contentId")
                    att_name = att.get("name", "").lower()
                    
                    # דלג על inline images (חתימות, לוגואים)
                    # אבל שמור PDF/Excel/Word/ZIP גם אם יש להם contentId
                    if is_inline and content_id and att_name.endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp')):
                        continue  # דלג על תמונות inline
                    
                    real_attachments.append(att)
                
                return real_attachments
            return []
        except Exception as e:
            logger.error(f"❌ Failed to retrieve attachments: {e}")
            return []
    
    def download_attachment(
        self,
        message_id: str,
        attachment_id: str,
        output_path: str
    ) -> bool:
        """
        הורד מצורף
        
        Args:
            message_id: ID הודעה
            attachment_id: ID המצורף
            output_path: נתיב שמירה
        
        Returns:
            True אם הצליח
        """
        try:
            endpoint = f"/users/{quote(self.shared_mailbox)}/messages/{message_id}/attachments/{attachment_id}"
            
            # קבל מידע על המצורף
            att_info = self._make_request("GET", endpoint)
            if not att_info:
                return False
            
            # דל את המצורף (contentBytes בנתונים המוקודדים בBase64)
            if "@odata.type" in att_info and att_info["@odata.type"] == "#microsoft.graph.fileAttachment":
                # זה קובץ - ניתן להורדה
                content = att_info.get("contentBytes")
                if content:
                    import base64
                    file_content = base64.b64decode(content)
                    
                    # וודא תיקייה
                    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
                    
                    with open(output_path, 'wb') as f:
                        f.write(file_content)
                    return True
            
            return False
        except Exception as e:
            logger.error(f"❌ Failed to download attachment: {e}")
            return False
    
    def create_folder(self, folder_name: str, parent_folder_id: str = "") -> Optional[str]:
        """
        יצור תיקייה חדשה
        
        Args:
            folder_name: שם התיקייה החדשה
            parent_folder_id: ID של התיקייה האב (ריק = root)
        
        Returns:
            ID של התיקייה החדשה או None אם נכשל
        """
        try:
            if parent_folder_id:
                endpoint = f"/users/{quote(self.shared_mailbox)}/mailFolders/{parent_folder_id}/childFolders"
            else:
                endpoint = f"/users/{quote(self.shared_mailbox)}/mailFolders"
            
            data = {"displayName": folder_name}
            result = self._make_request("POST", endpoint, data=data)
            
            if result:
                folder_id = result.get("id")
                self.folders_cache[folder_name] = folder_id
                return folder_id
            
            return None
        except Exception as e:
            logger.error(f"❌ Failed to create folder: {e}")
            return None
    
    def move_message(self, message_id: str, target_folder_id: str) -> bool:
        """העבר הודעה לתיקייה אחרת"""
        try:
            endpoint = f"/users/{quote(self.shared_mailbox)}/messages/{message_id}"
            data = {"parentFolderId": target_folder_id}
            result = self._make_request("PATCH", endpoint, data=data)
            return result is not None
        except Exception as e:
            logger.error(f"❌ Failed to move message: {e}")
            return False
    
    def delete_message(self, message_id: str) -> bool:
        """מחק הודעה"""
        try:
            endpoint = f"/users/{quote(self.shared_mailbox)}/messages/{message_id}"
            result = self._make_request("DELETE", endpoint)
            return result is not None
        except Exception as e:
            logger.error(f"❌ Failed to delete message: {e}")
            return False

    def list_categories(self) -> Dict[str, str]:
        """
        Get all master categories for this mailbox.
        Returns dict mapping category displayName → preset color (e.g. "preset9").
        """
        try:
            endpoint = f"/users/{quote(self.shared_mailbox)}/outlook/masterCategories"
            result = self._make_request("GET", endpoint)
            if result and result.get("value"):
                return {
                    cat.get("displayName", ""): cat.get("color", "preset0")
                    for cat in result["value"]
                    if cat.get("displayName")
                }
            return {}
        except Exception as e:
            logger.debug(f"Failed to list categories: {e}")
            return {}

    def ensure_category(self, category_name: str, color: str = "preset0") -> bool:
        """
        ודא שקטגוריה קיימת בתיבה. אם קיימת — לא משנה צבע.
        צבע משמש רק ליצירת קטגוריה חדשה.

        Args:
            category_name: שם הקטגוריה
            color: צבע ליצירה (preset0..preset25) — לא דורס צבע קיים

        Returns:
            True אם קיימת/נוצרה בהצלחה
        """
        try:
            if not category_name:
                logger.warning("⚠️ Empty category name")
                return False

            valid_colors = {f"preset{i}" for i in range(0, 26)}
            if color not in valid_colors:
                logger.warning(f"⚠️ Invalid color: {color}, using preset0")
                color = "preset0"

            logger.info(f"🔍 Checking category '{category_name}' in mailbox {self.shared_mailbox}")

            escaped = category_name.replace("'", "''")
            endpoint = f"/users/{quote(self.shared_mailbox)}/outlook/masterCategories"
            result = self._make_request("GET", endpoint, params={"$filter": f"displayName eq '{escaped}'"})

            if result and result.get("value"):
                existing = result["value"][0]
                existing_color = existing.get("color")
                logger.info(f"✓ Category '{category_name}' exists (color: {existing_color})")
                return True

            logger.info(f"➕ Creating new category '{category_name}' with color '{color}'")
            create_result = self._make_request("POST", endpoint, data={
                "displayName": category_name,
                "color": color
            })
            if create_result:
                logger.info("✓ Category created successfully")
            else:
                logger.error("❌ Category creation failed")
            return create_result is not None
        except Exception as e:
            logger.error(f"❌ Failed to ensure category: {e}")
            return False

    def ensure_category_for_mailbox(self, mailbox_address: str, category_name: str, color: str = "preset0") -> bool:
        """
        ודא שקטגוריה קיימת בתיבה ספציפית. אם קיימת — לא משנה צבע.
        צבע משמש רק ליצירת קטגוריה חדשה.

        Args:
            mailbox_address: כתובת תיבה
            category_name: שם הקטגוריה
            color: צבע ליצירה (preset0..preset25) — לא דורס צבע קיים

        Returns:
            True אם קיימת/נוצרה בהצלחה
        """
        try:
            if not mailbox_address or not category_name:
                logger.warning("⚠️ Mailbox address or category name is empty")
                return False

            valid_colors = {f"preset{i}" for i in range(0, 26)}
            if color not in valid_colors:
                logger.warning(f"⚠️ Invalid color: {color}, using preset0")
                color = "preset0"

            logger.info(f"🔍 Checking category '{category_name}' in mailbox {mailbox_address}")

            escaped = category_name.replace("'", "''")
            endpoint = f"/users/{quote(mailbox_address)}/outlook/masterCategories"
            result = self._make_request("GET", endpoint, params={"$filter": f"displayName eq '{escaped}'"})

            if result and result.get("value"):
                existing = result["value"][0]
                existing_color = existing.get("color")
                logger.info(f"✓ Category '{category_name}' exists in mailbox {mailbox_address} (color: {existing_color})")
                return True

            logger.info(f"➕ Creating new category '{category_name}' with color '{color}' in mailbox {mailbox_address}")
            create_result = self._make_request("POST", endpoint, data={
                "displayName": category_name,
                "color": color
            })
            if create_result:
                logger.info(f"✓ Category created successfully in mailbox {mailbox_address}")
            else:
                logger.error(f"❌ Category creation failed in mailbox {mailbox_address}")
                logger.error(f"Last error: {self.last_error}")
            return create_result is not None
        except Exception as e:
            logger.error(f"❌ Failed to ensure category for mailbox {mailbox_address}: {e}")
            return False
    
    def mark_message_with_category(self, message_id: str, category_name: str = "AI Processed") -> bool:
        """
        סמן הודעה עם קטגוריה (צבע) — הוסף לקטגוריות קיימות, לא החליף
        
        Args:
            message_id: מזהה ההודעה
            category_name: שם הקטגוריה (ברירת מחדל: "AI Processed")
        
        Returns:
            True אם הצליח
        """
        try:
            endpoint = f"/users/{quote(self.shared_mailbox)}/messages/{message_id}"
            
            # קרא קטגוריות קיימות
            get_result = self._make_request("GET", endpoint, params={"$select": "categories"})
            current_categories = []
            if get_result and "categories" in get_result:
                current_categories = get_result["categories"]
            
            # בדוק אם הקטגוריה כבר קיימת
            if category_name not in current_categories:
                current_categories.append(category_name)
            
            # עדכן עם הקטגוריות המוזמגות
            data = {
                "categories": current_categories
            }
            result = self._make_request("PATCH", endpoint, data=data)
            return result is not None
        except Exception as e:
            logger.error(f"❌ Failed to mark message with category: {e}")
            return False

    def replace_message_category(
        self, message_id: str, old_category: str, new_category: str
    ) -> bool:
        """
        Replace one category with another on a message.
        Removes old_category and adds new_category.
        """
        try:
            endpoint = f"/users/{quote(self.shared_mailbox)}/messages/{message_id}"
            get_result = self._make_request("GET", endpoint, params={"$select": "categories"})
            current_categories = []
            if get_result and "categories" in get_result:
                current_categories = get_result["categories"]

            updated = [c for c in current_categories if c != old_category]
            if new_category not in updated:
                updated.append(new_category)

            result = self._make_request("PATCH", endpoint, data={"categories": updated})
            return result is not None
        except Exception as e:
            logger.error(f"❌ Failed to replace category: {e}")
            return False

    def flag_message(self, message_id: str, flag_status: str = "flagged") -> bool:
        """
        סמן הודעה עם דגל
        
        Args:
            message_id: מזהה ההודעה
            flag_status: סטטוס הדגל (flagged, complete, notFlagged)
        
        Returns:
            True אם הצליח
        """
        try:
            endpoint = f"/users/{quote(self.shared_mailbox)}/messages/{message_id}"
            data = {
                "flag": {
                    "flagStatus": flag_status
                }
            }
            result = self._make_request("PATCH", endpoint, data=data)
            return result is not None
        except Exception as e:
            logger.error(f"❌ Failed to flag message: {e}")
            return False
    
    def send_email(
        self,
        to_address: str,
        subject: str,
        body: str = "",
        attachments: Optional[List[Any]] = None,
        body_type: str = "Text",
        replace_display_with_filename: bool = False
    ) -> bool:
        """
        שלח מייל מהתיבה המשותפת
        
        Args:
            to_address: כתובת נמען
            subject: נושא המייל
            body: גוף ההודעה
            attachments: רשימת קבצים מצורפים (Path objects או dicts עם 'path' ו-'display_name')
            body_type: סוג גוף (Text או HTML)
            replace_display_with_filename: אם True, משתמש ב-FILE_NAME בהודעה במקום DISPLAY_NAME
                                           (כלומר attachment_name = filename בלבד, לא display_name)
        
        Returns:
            True אם הצליח
        """
        try:
            import base64
            import time
            
            # בנה את ההודעה
            message = {
                "subject": subject,
                "body": {
                    "contentType": body_type,
                    "content": body
                },
                "toRecipients": [
                    {
                        "emailAddress": {
                            "address": to_address
                        }
                    }
                ]
            }
            
            # הוסף attachments אם יש
            if attachments and len(attachments) > 0:
                message["attachments"] = []
                for att in attachments:
                    display_name = None
                    if isinstance(att, dict):
                        att_path = Path(att.get('path', ''))
                        display_name = att.get('display_name')
                    else:
                        att_path = Path(att)
                    if not att_path.exists():
                        logger.warning(f"⚠️ File not found: {att_path}")
                        continue
                    
                    try:
                        with open(att_path, "rb") as f:
                            content = f.read()
                            content_b64 = base64.b64encode(content).decode('utf-8')
                        
                        # החלט על שם הקובץ בהודעה
                        if replace_display_with_filename:
                            # השתמש ב-FILE_NAME (שם הקובץ בדיסק) במקום DISPLAY_NAME
                            attachment_name = att_path.name
                            if display_name:
                                logger.info(f"📎 Replacing DISPLAY_NAME '{display_name}' with FILE_NAME '{att_path.name}'")
                        else:
                            # השתמש ב-DISPLAY_NAME אם יש, אחרת ב-FILE_NAME
                            attachment_name = display_name or att_path.name
                            if display_name and att_path.suffix:
                                attachment_name = f"{display_name}{att_path.suffix}"
                        
                        message["attachments"].append({
                            "@odata.type": "#microsoft.graph.fileAttachment",
                            "name": attachment_name,
                            "contentBytes": content_b64
                        })
                    except Exception as e:
                        logger.error(f"⚠️ Failed to read file {att_path.name}: {e}")
            
            # שלח (ללא retry - Graph API עשוי לשלוח למרות שגיאה 503)
            endpoint = f"/users/{quote(self.shared_mailbox)}/sendMail"
            data = {
                "message": message,
                "saveToSentItems": "true"
            }
            
            result = self._make_request("POST", endpoint, data=data)
            
            if result is not None:
                logger.info(f"✅ Email sent successfully to {to_address}")
                return True
            else:
                # בדוק אם זו שגיאה ידועה של MailboxStoreUnavailable
                # במקרים רבים המייל נשלח למרות השגיאה
                if self.last_error and 'MailboxStoreUnavailable' in str(self.last_error):
                    logger.error(f"⚠️ Email may have been sent despite error: {self.last_error}")
                    logger.info(f"ℹ️ Check sent items to verify delivery to {to_address}")
                    # נחזיר True כי סביר שהמייל נשלח
                    return True
                else:
                    logger.error(f"❌ Email send failed: {self.last_error}")
                    return False
                
        except Exception as e:
            logger.error(f"❌ Failed to send email: {e}")
            self.last_error = str(e)
            return False


def create_mailbox_connector(
    authenticator: GraphAuthenticator,
    shared_mailbox: str,
    **kwargs
) -> GraphMailboxConnector:
    """
    Factory function ליצירת mailbox connector
    
    Args:
        authenticator: GraphAuthenticator instance
        shared_mailbox: כתובת התיבה המשותפת
        **kwargs: פרמטרים נוספים
    
    Returns:
        GraphMailboxConnector instance
    """
    return GraphMailboxConnector(
        authenticator=authenticator,
        shared_mailbox=shared_mailbox,
        **kwargs
    )
