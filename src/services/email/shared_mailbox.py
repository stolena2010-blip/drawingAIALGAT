import os
from pathlib import Path
from typing import Optional, List, Dict, Callable

try:
    import win32com.client  # Requires pywin32
except Exception as e:  # noqa: F401
    win32com = None

from src.core.config import get_config

from src.utils.logger import get_logger
logger = get_logger(__name__)


class SharedMailboxConnector:
    """
    Simple connector for accessing a shared mailbox via Outlook COM on Windows.
    Uses values from `EmailConfig` in the global config.
    """

    def __init__(self, shared_mailbox: Optional[str] = None) -> None:
        cfg = get_config().email
        self.shared_mailbox = shared_mailbox or cfg.shared_mailbox
        self.method = cfg.method
        self._outlook = None
        self._ns = None
        self._inbox = None
        self.is_connected = False
        self.last_error: Optional[str] = None

        if not self.shared_mailbox:
            raise ValueError("Shared mailbox address is required")
        if self.method != "OUTLOOK_COM":
            raise ValueError("SharedMailboxConnector supports OUTLOOK_COM only. Use EWS connector for EWS.")

        if os.name != "nt":
            raise EnvironmentError("Outlook COM is only available on Windows")

    def connect(self) -> bool:
        """Establish Outlook COM connection and resolve shared inbox."""
        try:
            if win32com is None:
                raise ImportError("pywin32 is required for Outlook COM (pip install pywin32)")

            self._outlook = win32com.client.Dispatch("Outlook.Application")
            self._ns = self._outlook.GetNamespace("MAPI")

            recipient = self._ns.CreateRecipient(self.shared_mailbox)
            recipient.Resolve()
            if not recipient.Resolved:
                raise RuntimeError(f"Shared mailbox not resolved: {self.shared_mailbox}")

            # 6 = Inbox
            self._inbox = self._ns.GetSharedDefaultFolder(recipient, 6)
            self.is_connected = True
            self.last_error = None
            return True
        except Exception as e:
            logger.debug(f"Error handled: {e}")
            self.is_connected = False
            self.last_error = str(e)
            return False

    def sync_and_connect(self) -> bool:
        """Force Outlook to fully quit and restart (nuclear option)."""
        import gc
        import time
        
        
        # Close all our references
        self.close()
        
        try:
            # Force Outlook to quit completely
            if self._outlook is None:
                self._outlook = win32com.client.Dispatch("Outlook.Application")
            
            self._outlook.Quit()
            
            # Wait for Outlook to actually close
            time.sleep(2)
            
        except Exception as e:
            logger.debug(f"Ignored: {e}")
            pass
        
        # Reset Outlook object
        self._outlook = None
        gc.collect()
        time.sleep(1)
        
        
        # Reconnect with brand new Outlook instance
        ok = self.connect()
        if not ok:
            return False
        
        return self.is_connected

    def get_inbox_items(self):
        """Return COM Items collection for the shared mailbox inbox."""
        if self._inbox is None:
            raise RuntimeError("Not connected. Call connect() first.")
        items = self._inbox.Items
        items.Sort("[ReceivedTime]", True)
        return items

    def list_folders(self) -> List[str]:
        """Return a list of folder paths from the shared mailbox store."""
        # Reconnect each time to refresh Outlook cache (handles new folders without restart)
        if not self.connect():
            raise RuntimeError(self.last_error or "Not connected. Call connect() first.")
        if self._inbox is None:
            raise RuntimeError("Not connected. Call connect() first.")

        paths: List[str] = []
        
        # Add Inbox itself
        paths.append("Inbox")
        
        # Get subfolders directly from Inbox
        try:
            subfolders = self._inbox.Folders
            count = subfolders.Count
            
            
            for i in range(1, count + 1):
                try:
                    folder = subfolders.Item(i)
                    name = folder.Name
                    full_path = f"Inbox/{name}"
                    paths.append(full_path)
                    print(f"  Found: {full_path}")
                except Exception as e:
                    logger.debug(f"Error: {e}")
                    print(f"  Error getting subfolder {i}: {e}")
                    
        except Exception as e:
            logger.debug(f"Error: {e}")
            print(f"ERROR accessing Inbox.Folders: {e}")
        
        paths.sort(key=lambda s: s.lower())
        return paths

    def download_emails(
        self,
        folder_name: str = "INBOX",
        from_date: Optional[object] = None,
        to_date: Optional[object] = None,
        save_attachments: bool = True,
        output_folder: str = "email_downloads",
        progress_callback: Optional[Callable[[int, int, str], None]] = None,
    ) -> Dict:
        """
        Download emails from a named folder path and optionally save attachments.
        Folder path should match values returned by list_folders().
        """
        if self._ns is None:
            raise RuntimeError("Not connected. Call connect() first.")

        # Resolve folder by path
        target = self._resolve_folder_by_path(folder_name)
        if target is None:
            raise RuntimeError(f"Folder not found: {folder_name}")

        out = Path(output_folder)
        out.mkdir(parents=True, exist_ok=True)

        items = target.Items
        try:
            items.Sort("[ReceivedTime]", True)
        except Exception as e:
            logger.debug(f"Ignored: {e}")
            pass

        # Count total (COM Items may not have count easily when filtered)
        total = getattr(items, "Count", 0) or 0
        stats = {"total_emails": 0, "saved_attachments": 0, "errors": 0}

        def safe_filename(text: str, max_len: int = 150) -> str:
            import re as _re
            if not text:
                text = "No Subject"
            # Remove TAB characters and other invalid filename characters
            text = _re.sub(r'[\t<>:"/\\|?*\x00-\x1F]', "_", str(text)).strip()
            text = _re.sub(r"\s+", " ", text)
            return text[:max_len]

        # Iterate COM items (1-based)
        for i in range(1, total + 1):
            try:
                msg = items.Item(i)
                msg_class = getattr(msg, "MessageClass", "")
                if not str(msg_class).startswith("IPM.Note"):
                    continue

                stats["total_emails"] += 1
                if progress_callback:
                    progress_callback(stats["total_emails"], total, f"מעבד הודעה {stats['total_emails']}/{total}")

                # Save attachments
                if save_attachments:
                    subject = safe_filename(getattr(msg, "Subject", "") or "no_subject")
                    received = getattr(msg, "ReceivedTime", None)
                    received_str = "" if received is None else received.strftime("%Y%m%d_%H%M%S")
                    email_folder = out / f"{received_str}_{subject}"
                    email_folder.mkdir(exist_ok=True)

                    # Save email.txt with email metadata and body
                    try:
                        sender = getattr(msg, "SenderName", "") or ""
                        sender_email = getattr(msg, "SenderEmailAddress", "") or ""
                        body = getattr(msg, "Body", "") or ""
                        
                        email_txt_path = email_folder / "email.txt"
                        with open(email_txt_path, "w", encoding="utf-8") as f:
                            # שורה ראשונה: כתובת המייל של השולח בלבד
                            f.write(f"{sender_email}\n")
                            # שורה שנייה: הנושא
                            f.write(f"Subject: {getattr(msg, 'Subject', '') or 'No Subject'}\n")
                            f.write(f"From: {sender}")
                            if sender_email:
                                f.write(f" <{sender_email}>")
                            f.write(f"\nReceived: {received}\n")
                            f.write("-" * 60 + "\n")
                            f.write(body)
                    except Exception as e:
                        logger.debug(f"Error handled: {e}")
                        stats["errors"] += 1

                    atts = getattr(msg, "Attachments", None)
                    if atts:
                        for k in range(1, atts.Count + 1):
                            att = atts.Item(k)
                            nm = safe_filename(getattr(att, "FileName", f"attachment_{k}"), max_len=180)
                            dest = email_folder / nm
                            if dest.exists():
                                base = dest.stem
                                ext = dest.suffix
                                dest = email_folder / f"{base}__{k}{ext}"
                            try:
                                att.SaveAsFile(str(dest))
                                stats["saved_attachments"] += 1
                            except Exception as e:
                                logger.debug(f"Error handled: {e}")
                                stats["errors"] += 1
            except Exception as e:
                logger.debug(f"Error handled: {e}")
                stats["errors"] += 1

        return stats

    def _resolve_folder_by_path(self, path_str: str):
        """Resolve a folder object from a 'A/B/C' path under the shared mailbox root."""
        # Locate root
        root = None
        try:
            stores = self._ns.Stores
            for i in range(1, stores.Count + 1):
                store = stores.Item(i)
                name = getattr(store, "DisplayName", "") or ""
                if self.shared_mailbox.lower() in name.lower():
                    root = store.GetRootFolder()
                    break
        except Exception as e:
            logger.debug(f"Handled: {e}")
            root = None

        if root is None and self._inbox is not None:
            try:
                root = self._inbox.Parent
            except Exception as e:
                logger.debug(f"Handled: {e}")
                root = None

        if root is None:
            return None

        # Traverse by names (case-insensitive). If first part equals root name, skip it.
        parts = [p for p in (path_str or "").split("/") if p]
        node = root
        try:
            root_name = (getattr(root, "Name", "") or "").lower()
        except Exception as e:
            logger.debug(f"Handled: {e}")
            root_name = ""

        # Special-case for Inbox direct reference
        if len(parts) == 1 and parts[0].lower() in {"inbox", "תיבת דואר נכנס"} and self._inbox is not None:
            return self._inbox

        start_idx = 0
        if parts and parts[0].lower() == root_name:
            start_idx = 1

        for p in parts[start_idx:]:
            subs = getattr(node, "Folders", None)
            found = None
            if subs:
                for j in range(1, subs.Count + 1):
                    f = subs.Item(j)
                    fname = getattr(f, "Name", "") or ""
                    if fname.lower() == p.lower():
                        found = f
                        break
            node = found
            if node is None:
                return None
        return node

    def close(self) -> None:
        """Release COM references and force cleanup (aggressive)."""
        import gc
        
        
        # Release references
        self._inbox = None
        self._ns = None
        self._outlook = None
        self.is_connected = False
        
        # Force garbage collection to release COM locks
        gc.collect()

    # Alias for GUI API parity
    def disconnect(self) -> None:
        self.close()
        self.is_connected = False

    def create_subfolder(self, folder_name: str) -> bool:
        """Create a subfolder under Inbox."""
        if self._inbox is None:
            raise RuntimeError("Not connected. Call connect() first.")
        
        try:
            folders = self._inbox.Folders
            new_folder = folders.Add(folder_name)
            return True
        except Exception as e:
            logger.warning(f"Error (re-raising): {e}")
            raise RuntimeError(f"Cannot create subfolder '{folder_name}': {e}")

    def diagnose(self) -> Dict:
        """Lightweight diagnose for GUI parity (shared mailbox via COM)."""
        results = {
            "server": "Outlook COM",
            "user": "",
            "shared_mailbox": self.shared_mailbox or "",
            "access_type": "OUTLOOK_COM",
            "checks": []
        }

        try:
            ok = self.connect()
            if ok:
                inbox_count = 0
                try:
                    inbox_count = self._inbox.Items.Count
                except Exception as e:
                    logger.debug(f"Ignored: {e}")
                    pass
                results["checks"].append({
                    "address": self.shared_mailbox,
                    "access_type": "OUTLOOK_COM",
                    "ok": True,
                    "inbox_count": inbox_count,
                })
            else:
                msg = self.last_error or "Unknown error"
                results["checks"].append({
                    "address": self.shared_mailbox,
                    "access_type": "OUTLOOK_COM",
                    "ok": False,
                    "error": msg
                })
        except Exception as e:
            logger.debug(f"Handled: {e}")
            results["checks"].append({
                "address": self.shared_mailbox,
                "access_type": "OUTLOOK_COM",
                "ok": False,
                "error": str(e)
            })

        return results

    def diagnose_folder_access(self) -> Dict:
        """Diagnose folder access issues and permissions."""
        if self._inbox is None:
            return {"error": "Not connected"}
        
        results = {
            "inbox_name": "",
            "inbox_path": "",
            "store_name": "",
            "subfolders_count": 0,
            "subfolders": [],
            "can_create": False,
            "can_delete": False,
            "errors": []
        }
        
        # Get Inbox info
        try:
            results["inbox_name"] = self._inbox.Name
            results["inbox_path"] = self._inbox.FolderPath
        except Exception as e:
            logger.debug(f"Error handled: {e}")
            results["errors"].append(f"Cannot read Inbox info: {e}")
        
        # Get Store info
        try:
            store = self._inbox.Store
            results["store_name"] = store.DisplayName
        except Exception as e:
            logger.debug(f"Error handled: {e}")
            results["errors"].append(f"Cannot read Store info: {e}")
        
        # Try to get subfolders
        try:
            folders = self._inbox.Folders
            count = folders.Count
            results["subfolders_count"] = count
            
            print(f"\n=== INBOX SUBFOLDER DIAGNOSTIC ===")
            print(f"Inbox: {results['inbox_name']}")
            print(f"Path: {results['inbox_path']}")
            print(f"Store: {results['store_name']}")
            print(f"Subfolders count: {count}")
            
            for i in range(1, count + 1):
                try:
                    folder = folders.Item(i)
                    name = folder.Name
                    results["subfolders"].append(name)
                    print(f"  {i}. {name}")
                except Exception as e:
                    logger.debug(f"Error handled: {e}")
                    error = f"Error accessing subfolder {i}: {e}"
                    results["errors"].append(error)
                    print(f"  {i}. ERROR: {e}")
            
        except Exception as e:
            logger.debug(f"Error handled: {e}")
            error = f"Cannot access Folders collection: {e}"
            results["errors"].append(error)
            print(f"ERROR: {error}")
        
        # Test write permissions by trying to create/delete a test folder
        test_folder_name = "_TEST_PERMISSIONS_DELETE_ME"
        test_folder = None
        
        try:
            print(f"\nTesting CREATE permission...")
            test_folder = folders.Add(test_folder_name)
            results["can_create"] = True
            print(f"✓ Successfully created test folder '{test_folder_name}'")
            
            # Try to delete it
            try:
                print(f"Testing DELETE permission...")
                test_folder.Delete()
                results["can_delete"] = True
                print(f"✓ Successfully deleted test folder")
            except Exception as e:
                logger.debug(f"Error handled: {e}")
                results["errors"].append(f"Cannot delete test folder: {e}")
                print(f"✗ Cannot delete test folder: {e}")
                print(f"⚠ Please manually delete '{test_folder_name}' from Inbox")
                
        except Exception as e:
            logger.debug(f"Error handled: {e}")
            results["errors"].append(f"Cannot create folder: {e}")
            print(f"✗ Cannot create test folder: {e}")
            print(f"  This likely means no WRITE permissions on this mailbox")
        
        print(f"=== END DIAGNOSTIC ===\n")
        
        return results
        """Return a simple diagnosis structure compatible with the GUI expectations."""
        results = {
            "server": "Outlook COM",
            "user": "",
            "shared_mailbox": self.shared_mailbox or "",
            "access_type": "OUTLOOK_COM",
            "checks": []
        }

        try:
            ok = self.connect()
            if ok:
                inbox_count = getattr(self._inbox.Items, "Count", 0) if self._inbox is not None else 0
                results["checks"].append({
                    "address": self.shared_mailbox,
                    "access_type": "OUTLOOK_COM",
                    "ok": True,
                    "inbox_count": inbox_count,
                })
            else:
                msg = self.last_error or "Unknown error"
                suggestion = None
                if "not resolved" in msg.lower():
                    suggestion = "בדוק/י שהוספת הרשאת Full Access לתיבה המשותפת ב-Outlook/Exchange"
                results["checks"].append({
                    "address": self.shared_mailbox,
                    "access_type": "OUTLOOK_COM",
                    "ok": False,
                    "error": msg,
                    "suggestion": suggestion
                })
        except Exception as e:
            logger.debug(f"Handled: {e}")
            results["checks"].append({
                "address": self.shared_mailbox,
                "access_type": "OUTLOOK_COM",
                "ok": False,
                "error": str(e)
            })

        return results

    def send_email(
        self,
        to_address: str,
        subject: str,
        body: str,
        attachments: Optional[List[Path]] = None,
        cc_addresses: Optional[List[str]] = None
    ) -> bool:
        """
        שליחת מייל דרך Outlook COM
        
        Args:
            to_address: כתובת יעד
            subject: נושא ההודעה
            body: תוכן ההודעה
            attachments: רשימת קבצים מצורפים (paths)
            cc_addresses: רשימת כתובות עותק
        
        Returns:
            True אם נשלח בהצלחה
        """
        if not self.is_connected:
            self.last_error = "Not connected to server"
            print(f"❌ {self.last_error}")
            return False
        
        try:
            print(f"📤 Sending email to {to_address} from {self.shared_mailbox}...")
            
            # צור הודעה חדשה
            mail = self._outlook.CreateItem(0)  # 0 = olMailItem
            
            # הגדר את השולח לתיבה המשותפת
            mail.SentOnBehalfOfName = self.shared_mailbox
            
            # הגדר נמען
            mail.To = to_address
            
            # הוסף עותק אם יש
            if cc_addresses:
                mail.CC = "; ".join(cc_addresses)
            
            # הגדר נושא ותוכן
            mail.Subject = subject
            mail.Body = body
            
            # הוסף קבצים מצורפים אם יש
            if attachments:
                for file_path in attachments:
                    if isinstance(file_path, str):
                        file_path = Path(file_path)
                    
                    if file_path.exists():
                        mail.Attachments.Add(str(file_path.absolute()))
                        print(f"📎 Attached: {file_path.name}")
                    else:
                        print(f"⚠️ File not found: {file_path}")
            
            # שלח
            mail.Send()
            
            print("✅ Email sent successfully!")
            self.last_error = None
            return True
            
        except Exception as e:
            logger.debug(f"Error handled: {e}")
            error_msg = str(e)
            self.last_error = error_msg
            print(f"❌ Failed to send email: {error_msg}")
            return False
