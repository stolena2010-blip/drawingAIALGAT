"""
Graph API Helper - Simplified Interface
========================================

ממשק מפושט לעבודה עם Graph API או COM
מאפשר קל לעבור בין שיטות שונות
"""

from typing import Optional, List, Dict, Any
import re
import html
from datetime import datetime
from pathlib import Path

from src.core.config import get_config
from .graph_auth import GraphAuthenticator, create_authenticator
from .graph_mailbox import GraphMailboxConnector, create_mailbox_connector

from src.utils.logger import get_logger
logger = get_logger(__name__)


def _sanitize_filename(name: str, max_length: int = 180) -> str:
    """Sanitize a filename by removing invalid characters."""
    safe = re.sub(r'[\t<>:"/\\|?*\x00-\x1F]', "_", str(name)).strip()
    safe = re.sub(r"\s+", " ", safe)
    return safe[:max_length]


def _clean_html_body(raw_body: str) -> str:
    """Convert HTML email body to clean plain text."""
    body = raw_body
    body = re.sub(r"<br\s*/?>", "\n", body, flags=re.IGNORECASE)
    body = re.sub(r"</p>", "\n\n", body, flags=re.IGNORECASE)
    body = re.sub(r"</div>", "\n", body, flags=re.IGNORECASE)
    body = re.sub(r"<[^>]+>", "", body)
    body = html.unescape(body)
    body = re.sub(r"\n{3,}", "\n\n", body)
    body = "\n".join(line.strip() for line in body.split("\n"))
    signature_patterns = [
        r"(?:Sent from|שולח מ)[^\n]*",
        r"(?:Best regards|Regards|Thanks|Thank you)[,\s]*\n.*$",
        r"(?:בברכה|תודה)[,\s]*\n.*$",
    ]
    for pattern in signature_patterns:
        body = re.sub(pattern, "", body, flags=re.IGNORECASE | re.MULTILINE)
    return body.strip()


def _save_email_txt(msg_dir: Path, sender: str, subject: str, received_time: str, body: str, web_link: str = "") -> None:
    """Save email details to a standardized email.txt file."""
    with open(msg_dir / "email.txt", "w", encoding="utf-8") as f:
        f.write(f"{sender}\n")
        f.write(f"נושא: {subject}\n")
        if received_time:
            try:
                dt = datetime.fromisoformat(received_time.replace("Z", "+00:00"))
                f.write(f"תאריך: {dt.strftime('%d/%m/%Y %H:%M')}\n")
            except Exception as e:
                logger.debug(f"Ignored: {e}")
                pass
        if web_link:
            f.write(f"מייל מקור: {web_link}\n")
        f.write("\n" + "=" * 70 + "\n\n")
        f.write(body or "(ללא תוכן)")


def _download_attachments(mailbox, msg_id: str, msg_dir: Path) -> int:
    """Download all attachments for a message. Returns count."""
    count = 0
    attachments = mailbox.get_attachments(msg_id)
    for att in attachments:
        att_id = att.get("id")
        att_name = _sanitize_filename(att.get("name", "attachment"))
        att_path = msg_dir / att_name
        if mailbox.download_attachment(msg_id, att_id, str(att_path)):
            count += 1
    return count


class GraphAPIHelper:
    """
    Helper class לעבודה פשוטה עם Microsoft Graph API
    """
    
    def __init__(
        self,
        tenant_id: Optional[str] = None,
        client_id: Optional[str] = None,
        client_secret: Optional[str] = None,
        shared_mailbox: Optional[str] = None,
        use_config: bool = True
    ) -> None:
        """
        אתחול helper
        
        Args:
            tenant_id: Azure Tenant ID
            client_id: Application (Client) ID
            client_secret: Client Secret Value
            shared_mailbox: כתובת התיבה המשותפת
            use_config: אם True, טען מ-config אם לא סופקו פרמטרים
        """
        self.authenticator: Optional[GraphAuthenticator] = None
        self.mailbox: Optional[GraphMailboxConnector] = None
        self.last_error: Optional[str] = None
        
        # אם צריך טעינה מ-config
        if use_config and not all([tenant_id, client_id, client_secret]):
            cfg = get_config()
            if cfg.email.graph_api and cfg.email.graph_api.is_configured():
                tenant_id = cfg.email.graph_api.tenant_id
                client_id = cfg.email.graph_api.client_id
                client_secret = cfg.email.graph_api.client_secret
                shared_mailbox = shared_mailbox or cfg.email.graph_api.shared_mailbox or cfg.email.shared_mailbox
        
        # וודא שיש את כל הנדרש
        if not all([tenant_id, client_id, client_secret]):
            raise ValueError("Missing required Graph API credentials")
        
        if not shared_mailbox:
            raise ValueError("Missing shared mailbox address")
        
        # בנה authenticator
        try:
            self.authenticator = create_authenticator(
                tenant_id=tenant_id,
                client_id=client_id,
                client_secret=client_secret
            )
            
            # בנה mailbox connector
            self.mailbox = create_mailbox_connector(
                authenticator=self.authenticator,
                shared_mailbox=shared_mailbox
            )
        except Exception as e:
            logger.debug(f"Error handled: {e}")
            self.last_error = str(e)
            raise

    def test_connection(self) -> bool:
        """בדוק אם החיבור עובד"""
        if not self.authenticator or not self.mailbox:
            self.last_error = "Not initialized"
            return False
        
        # בדוק token
        if not self.authenticator.test_connection():
            self.last_error = self.authenticator.last_error
            return False
        
        # בדוק גישה לתיבה
        if not self.mailbox.connect():
            self.last_error = self.mailbox.last_error
            return False
        
        return True
    
    def download_messages(
        self,
        output_dir: str,
        folder_name: str = "Inbox",
        limit: int = 1000,
        save_progress: bool = True,
        progress_callback: Optional[Any] = None
    ) -> Dict[str, Any]:
        """
        הורד מיילים מתיקייה לתיקיית מקומית
        
        Args:
            output_dir: תיקיית פלט
            folder_name: שם התיקייה
            limit: מספר הודעות מקסימלי (ברירת מחדל: 1000)
            save_progress: אם True, שמור רשימת IDs שהורדו
            progress_callback: callback להצגת התקדמות
        
        Returns:
            Dict עם סטטיסטיקות
        """
        result = {
            "success": False,
            "messages_downloaded": 0,
            "attachments_downloaded": 0,
            "total_messages": 0,
            "total_attachments": 0,
            "errors": [],
            "output_dir": output_dir
        }
        
        try:
            # וודא תיקייה
            Path(output_dir).mkdir(parents=True, exist_ok=True)
            
            # קבל ID של התיקייה (חפש גם תתי-תיקיות של Inbox)
            folder_id = None

            folders = self.mailbox.list_folders(show_inbox_only=False)
            for folder in folders:
                if folder.get("displayName", "").lower() == folder_name.lower():
                    folder_id = folder.get("id")
                    break

            if not folder_id:
                inbox_folders = self.mailbox.list_folders(show_inbox_only=True)
                for folder in inbox_folders:
                    if folder.get("displayName", "").lower() == folder_name.lower():
                        folder_id = folder.get("id")
                        break

            if not folder_id:
                result["errors"].append(f"Folder not found: {folder_name}")
                return result
            
            # קבל הודעות
            messages = self.mailbox.list_messages(folder_id, limit=limit)
            
            total = len(messages)
            result["total_messages"] = total
            for idx, msg in enumerate(messages, start=1):
                try:
                    msg_id = msg.get("id")
                    subject = msg.get("subject", "No Subject")
                    received = msg.get("receivedDateTime", "")
                    has_attachments = msg.get("hasAttachments", False)
                    
                    # בנה שם תיקיית הודעה
                    safe_subject = _sanitize_filename(subject, max_length=100)

                    try:
                        received_dt = datetime.fromisoformat(received.replace("Z", "+00:00"))
                        received_safe = received_dt.strftime("%Y%m%d_%H%M%S")
                    except Exception as e:
                        logger.debug(f"Handled: {e}")
                        received_safe = datetime.now().strftime("%Y%m%d_%H%M%S")

                    msg_dir = Path(output_dir) / f"{received_safe}__{safe_subject}"
                    msg_dir.mkdir(exist_ok=True)
                    
                    if progress_callback:
                        progress_callback(idx, total, f"קריאת הודעה {idx}/{total}")
                    
                    # שמור פרטי הודעה בקובץ טקסט רק אם יש צורך
                    details = self.mailbox.get_message_details(msg_id)
                    if details:
                        sender = details.get("from", {}).get("emailAddress", {}).get("address", "")
                        subject_full = details.get("subject", "")
                        body = details.get("body", {}).get("content", "")
                        received_time = details.get("receivedDateTime", "")
                        web_link = msg.get("webLink", "") or details.get("webLink", "")

                        cleaned_body = _clean_html_body(body)
                        _save_email_txt(
                            msg_dir=msg_dir,
                            sender=sender,
                            subject=subject_full,
                            received_time=received_time,
                            body=cleaned_body,
                            web_link=web_link,
                        )
                    
                    # הורד מצורפים בלבד
                    if has_attachments:
                        if progress_callback:
                            progress_callback(idx, total, f"הורדת מצורפים {idx}/{total}")

                        downloaded = _download_attachments(self.mailbox, msg_id, msg_dir)
                        result["attachments_downloaded"] += downloaded
                        result["total_attachments"] += downloaded
                    
                    result["messages_downloaded"] += 1
                    if progress_callback:
                        progress_callback(idx, total, f"הורדת הודעה {idx}/{total} ✓")
                
                except Exception as e:
                    logger.debug(f"Error handled: {e}")
                    result["errors"].append(f"Error with message {msg.get('id')}: {str(e)}")
                    if progress_callback:
                        progress_callback(idx, total, f"שגיאה בהודעה {idx}/{total}")
            
            result["success"] = True
            
        except Exception as e:
            logger.debug(f"Error handled: {e}")
            result["errors"].append(str(e))
        
        return result
    
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
            attachments: רשימת קבצים מצורפים
            body_type: סוג גוף המייל (Text או HTML)
            replace_display_with_filename: אם True, משתמש ב-FILE_NAME בהודעה במקום DISPLAY_NAME
        
        Returns:
            True אם הצליח
        """
        try:
            if not self.mailbox:
                self.last_error = "Mailbox not initialized"
                return False
            
            result = self.mailbox.send_email(
                to_address=to_address,
                subject=subject,
                body=body,
                attachments=attachments,
                body_type=body_type,
                replace_display_with_filename=replace_display_with_filename
            )
            
            if not result:
                self.last_error = self.mailbox.last_error
            
            return result
            
        except Exception as e:
            logger.debug(f"Error handled: {e}")
            self.last_error = str(e)
            return False

    def mark_message_processed(self, message_id: str, category_name: str = "AI Processed") -> bool:
        """
        סמן מייל כמעובד על ידי AI
        
        Args:
            message_id: מזהה המייל
            category_name: שם הקטגוריה (ברירת מחדל: "AI Processed")
        
        Returns:
            True אם הצליח
        """
        try:
            if not self.mailbox:
                self.last_error = "Mailbox not initialized"
                return False
            
            result = self.mailbox.mark_message_with_category(message_id, category_name)
            
            if not result:
                self.last_error = self.mailbox.last_error
            
            return result
            
        except Exception as e:
            logger.debug(f"Error handled: {e}")
            self.last_error = str(e)
            return False
    
    def flag_message(self, message_id: str, flag_status: str = "flagged") -> bool:
        """
        סמן מייל עם דגל
        
        Args:
            message_id: מזהה המייל
            flag_status: סטטוס הדגל (flagged, complete, notFlagged)
        
        Returns:
            True אם הצליח
        """
        try:
            if not self.mailbox:
                self.last_error = "Mailbox not initialized"
                return False
            
            result = self.mailbox.flag_message(message_id, flag_status)
            
            if not result:
                self.last_error = self.mailbox.last_error
            
            return result
            
        except Exception as e:
            logger.debug(f"Error handled: {e}")
            self.last_error = str(e)
            return False

    def replace_category(self, message_id: str, old_category: str, new_category: str) -> bool:
        """Replace one category with another on a message."""
        try:
            if not self.mailbox:
                self.last_error = "Mailbox not initialized"
                return False
            result = self.mailbox.replace_message_category(message_id, old_category, new_category)
            if not result:
                self.last_error = self.mailbox.last_error
            return result
        except Exception as e:
            logger.debug(f"Error handled: {e}")
            self.last_error = str(e)
            return False

    def ensure_category(self, category_name: str, color: str = "preset0") -> bool:
        """
        ודא שקטגוריה קיימת עם צבע

        Args:
            category_name: שם קטגוריה
            color: צבע (preset0..preset25)

        Returns:
            True אם קיימת/נוצרה בהצלחה
        """
        try:
            if not self.mailbox:
                self.last_error = "Mailbox not initialized"
                return False

            result = self.mailbox.ensure_category(category_name, color)

            if not result:
                self.last_error = self.mailbox.last_error

            return result
        except Exception as e:
            logger.debug(f"Error handled: {e}")
            self.last_error = str(e)
            return False

    def ensure_category_for_mailbox(self, mailbox_address: str, category_name: str, color: str = "preset0") -> bool:
        """
        ודא שקטגוריה קיימת עם צבע עבור תיבה ספציפית

        Args:
            mailbox_address: כתובת תיבה
            category_name: שם קטגוריה
            color: צבע (preset0..preset25)

        Returns:
            True אם קיימת/נוצרה בהצלחה
        """
        try:
            if not self.mailbox:
                self.last_error = "Mailbox not initialized"
                return False

            result = self.mailbox.ensure_category_for_mailbox(mailbox_address, category_name, color)

            if not result:
                self.last_error = self.mailbox.last_error

            return result
        except Exception as e:
            logger.debug(f"Error handled: {e}")
            self.last_error = str(e)
            return False

    def download_message_by_id(self, message_id: str, output_dir: str) -> Dict[str, Any]:
        """
        הורד הודעה בודדת לפי ID לתיקייה מקומית.

        Args:
            message_id: מזהה ההודעה
            output_dir: תיקיית פלט

        Returns:
            Dict עם סטטיסטיקות ונתיבים
        """
        result = {
            "success": False,
            "message_id": message_id,
            "message_dir": None,
            "sender": "",
            "subject": "",
            "received": "",
            "attachments_downloaded": 0,
            "errors": []
        }

        try:
            Path(output_dir).mkdir(parents=True, exist_ok=True)

            details = self.mailbox.get_message_details(message_id) if self.mailbox else None
            if not details:
                result["errors"].append("No message details")
                return result

            sender = details.get("from", {}).get("emailAddress", {}).get("address", "")
            subject_full = details.get("subject", "")
            body = details.get("body", {}).get("content", "")
            received_time = details.get("receivedDateTime", "")
            has_attachments = details.get("hasAttachments", False)
            web_link = details.get("webLink", "")
            
            # שמור את גוף המייל המקורי (HTML) לשימוש בשליחה
            original_body_html = body
            original_body_type = details.get("body", {}).get("contentType", "Text")

            # בנה שם תיקיית הודעה
            safe_subject = _sanitize_filename(subject_full, max_length=100)

            try:
                received_dt = datetime.fromisoformat(received_time.replace("Z", "+00:00"))
                received_safe = received_dt.strftime("%Y%m%d_%H%M%S")
            except Exception as e:
                logger.debug(f"Handled: {e}")
                received_safe = datetime.now().strftime("%Y%m%d_%H%M%S")

            msg_dir = Path(output_dir) / f"{received_safe}__{safe_subject}"
            msg_dir.mkdir(exist_ok=True)
            result["message_dir"] = str(msg_dir)

            cleaned_body = _clean_html_body(body)
            _save_email_txt(
                msg_dir=msg_dir,
                sender=sender,
                subject=subject_full,
                received_time=received_time,
                body=cleaned_body,
                web_link=web_link,
            )

            if has_attachments and self.mailbox:
                result["attachments_downloaded"] += _download_attachments(self.mailbox, message_id, msg_dir)

            result["success"] = True
            result["sender"] = sender
            result["subject"] = subject_full
            result["received"] = received_time
            result["original_body_html"] = original_body_html  # גוף מייל מקורי עם HTML
            result["original_body_type"] = original_body_type  # סוג התוכן (HTML/Text)
            result["web_link"] = web_link  # URL לפתיחת המייל ב-Outlook Web
            result["categories"] = details.get("categories", [])  # Outlook categories (color labels)
            return result

        except Exception as e:
            logger.debug(f"Error handled: {e}")
            result["errors"].append(str(e))
            return result
    
    def diagnose_all(self) -> Dict[str, Any]:
        """אבחון מלא"""
        result = {
            "authenticator": self.authenticator.diagnose() if self.authenticator else None,
            "mailbox": self.mailbox.diagnose() if self.mailbox else None,
            "last_error": self.last_error
        }
        return result


def quick_test(
    tenant_id: str,
    client_id: str,
    client_secret: str,
    shared_mailbox: str
) -> bool:
    """
    בדיקה מהירה של חיבור Graph API
    
    Returns:
        True אם הצליח
    """
    try:
        helper = GraphAPIHelper(
            tenant_id=tenant_id,
            client_id=client_id,
            client_secret=client_secret,
            shared_mailbox=shared_mailbox,
            use_config=False
        )
        return helper.test_connection()
    except Exception as e:
        logger.debug(f"Error: {e}")
        print(f"❌ Test failed: {e}")
        return False
