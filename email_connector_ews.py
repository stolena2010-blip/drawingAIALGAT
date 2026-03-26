"""
Exchange Web Services (EWS) Email Connector
============================================

חיבור למייל דרך EWS במקום IMAP - עובד עם Microsoft 365
"""

from exchangelib import Credentials, Account, Configuration, DELEGATE, IMPERSONATION
from exchangelib.protocol import BaseProtocol, NoVerifyHTTPAdapter
from pathlib import Path
from typing import List, Dict, Optional
from datetime import datetime
import os


# השבת אזהרות SSL (לסביבות פיתוח בלבד)
BaseProtocol.HTTP_ADAPTER_CLS = NoVerifyHTTPAdapter


class EWSEmailConnector:
    """
    מתחבר למייל דרך Exchange Web Services
    """
    
    def __init__(
        self,
        email_address: str,
        password: str,
        server: str = 'outlook.office365.com',
        shared_mailbox: Optional[str] = None,
        access_type: str = 'DELEGATE'
    ):
        """
        אתחול
        
        Args:
            email_address: כתובת מייל
            password: סיסמה
            server: שרת Exchange (ברירת מחדל: outlook.office365.com)
            shared_mailbox: כתובת מייל משותפת (אם יש)
            access_type: סוג גישה ('DELEGATE' או 'IMPERSONATION')
        """
        self.email_address = email_address
        self.password = password
        self.server = server
        self.shared_mailbox = shared_mailbox
        self.access_type = access_type.upper() if access_type else 'DELEGATE'
        if self.access_type not in ('DELEGATE', 'IMPERSONATION'):
            self.access_type = 'DELEGATE'
        self.account = None
        self.is_connected = False
        self.last_error = None
        self.effective_address = self.shared_mailbox or self.email_address
    
    def connect(self) -> bool:
        """
        התחברות לשרת Exchange
        
        Returns:
            True אם הצליח
        """
        try:
            print(f"🔌 Connecting to {self.server} via EWS...")
            
            # צור credentials
            credentials = Credentials(username=self.email_address, password=self.password)
            
            # התחבר לחשבון
            self.account = Account(
                primary_smtp_address=self.shared_mailbox or self.email_address,
                credentials=credentials,
                autodiscover=True,  # גילוי אוטומטי של הגדרות השרת
                access_type=DELEGATE if self.access_type == 'DELEGATE' else IMPERSONATION
            )
            
            # בדוק חיבור על ידי גישה לתיקיית INBOX
            inbox_count = self.account.inbox.total_count
            
            self.is_connected = True
            self.last_error = None
            print(f"✅ Connected successfully! ({self.access_type}) user: {self.email_address} → mailbox: {self.account.primary_smtp_address}")
            print(f"   Found {inbox_count} messages in INBOX")
            return True
            
        except Exception as e:
            error_msg = str(e)
            self.last_error = error_msg
            print(f"❌ EWS connection error: {error_msg}")
            self.is_connected = False
            return False
    
    def disconnect(self):
        """ניתוק מהשרת"""
        self.account = None
        self.is_connected = False
        print("🔌 Disconnected from server")

    def diagnose(self) -> Dict:
        """אבחון חיבור: בודק תיבת משתמש ותיבה משותפת אם מוגדרת"""
        results = {
            "server": self.server,
            "user": self.email_address,
            "shared_mailbox": self.shared_mailbox or "",
            "access_type": self.access_type,
            "checks": []
        }

        def _attempt(address: str, access: str) -> Dict:
            try:
                creds = Credentials(username=self.email_address, password=self.password)
                acc = Account(
                    primary_smtp_address=address,
                    credentials=creds,
                    autodiscover=True,
                    access_type=DELEGATE if access == 'DELEGATE' else IMPERSONATION
                )
                count = acc.inbox.total_count
                return {"address": address, "access_type": access, "ok": True, "inbox_count": count}
            except Exception as e:
                msg = str(e)
                suggestion = None
                if 'Access is denied' in msg or 'denied' in msg.lower():
                    suggestion = "בדוק הרשאות Full Access / Application Impersonation לתיבה"
                elif 'not found' in msg.lower():
                    suggestion = "וודא כתובת התיבה המשותפת נכונה ומוגדרת בארגון"
                elif 'Unauthorized' in msg or '401' in msg:
                    suggestion = "בדוק סיסמה/תוקף חשבון והתחברות רב-גורמית"
                return {"address": address, "access_type": access, "ok": False, "error": msg, "suggestion": suggestion}

        # אם מוגדרת תיבה משותפת — עבד רק מולה
        if self.shared_mailbox:
            results["checks"].append(_attempt(self.shared_mailbox, self.access_type))
        else:
            # אחרת, בדוק תיבת המשתמש
            results["checks"].append(_attempt(self.email_address, 'DELEGATE'))

        return results
    
    def list_folders(self) -> List[str]:
        """
        רשימת כל התיקיות
        
        Returns:
            רשימת שמות תיקיות
        """
        if not self.is_connected:
            raise Exception("לא מחובר! קרא ל-connect() קודם")
        
        folders = []
        for folder in self.account.root.walk():
            folders.append(folder.name)
        
        return folders
    
    def download_emails(
        self,
        folder_name: str = 'INBOX',
        from_date: Optional[datetime] = None,
        to_date: Optional[datetime] = None,
        save_attachments: bool = True,
        output_folder: str = 'email_downloads',
        progress_callback: Optional[callable] = None
    ) -> Dict:
        """
        הורדת מיילים מתיקייה
        
        Args:
            folder_name: שם התיקייה
            from_date: תאריך התחלה
            to_date: תאריך סיום
            save_attachments: האם לשמור קבצים מצורפים
            output_folder: תיקיית פלט
            progress_callback: פונקציה לעדכון התקדמות
        
        Returns:
            מילון עם סטטיסטיקות
        """
        if not self.is_connected:
            raise Exception("לא מחובר! קרא ל-connect() קודם")
        
        # בחר תיקייה
        if folder_name.upper() == 'INBOX':
            folder = self.account.inbox
        elif folder_name.upper() == 'SENT':
            folder = self.account.sent
        else:
            # חפש תיקייה מותאמת
            folder = None
            for f in self.account.root.walk():
                if f.name == folder_name:
                    folder = f
                    break
            if not folder:
                raise Exception(f"תיקייה לא נמצאה: {folder_name}")
        
        # צור תיקיית פלט
        output_path = Path(output_folder)
        output_path.mkdir(exist_ok=True)
        
        stats = {
            'total_emails': 0,
            'saved_attachments': 0,
            'errors': 0
        }
        
        # קבל הודעות
        items = folder.all()
        
        # סנן לפי תאריכים אם צורך
        if from_date:
            items = items.filter(datetime_received__gte=from_date)
        if to_date:
            items = items.filter(datetime_received__lte=to_date)
        
        total = items.count()
        print(f"📥 Found {total} messages")
        
        for i, item in enumerate(items, 1):
            try:
                stats['total_emails'] += 1
                
                # עדכן התקדמות
                if progress_callback:
                    progress_callback(i, total, f"מעבד הודעה {i}/{total}")
                
                # שמור קבצים מצורפים
                if save_attachments and item.attachments:
                    subject = self._clean_filename(item.subject or 'no_subject')
                    email_folder = output_path / f"{i}_{subject}"
                    email_folder.mkdir(exist_ok=True)
                    
                    # שמור email.txt עם כתובת המייל וניתונים נוספים
                    try:
                        sender_email = str(item.sender.email_address) if hasattr(item.sender, 'email_address') else ""
                        email_txt_path = email_folder / "email.txt"
                        with open(email_txt_path, "w", encoding="utf-8") as f:
                            # שורה ראשונה: כתובת המייל של השולח בלבד
                            f.write(f"{sender_email}\n")
                            # שורה שנייה: הנושא
                            f.write(f"Subject: {item.subject or 'No Subject'}\n")
                            f.write(f"From: {item.sender}\n")
                            f.write(f"Received: {item.datetime_received}\n")
                            f.write("-" * 60 + "\n")
                            f.write(item.body or "")
                    except Exception:
                        pass
                    
                    for attachment in item.attachments:
                        if hasattr(attachment, 'name'):
                            file_path = email_folder / attachment.name
                            
                            with open(file_path, 'wb') as f:
                                f.write(attachment.content)
                            
                            stats['saved_attachments'] += 1
                            print(f"💾 Saved: {attachment.name}")
            
            except Exception as e:
                print(f"⚠️ Error in message {i}: {e}")
                stats['errors'] += 1
        
        return stats
    
    def send_email(
        self,
        to_address: str,
        subject: str,
        body: str,
        attachments: Optional[List[Path]] = None,
        cc_addresses: Optional[List[str]] = None
    ) -> bool:
        """
        שליחת מייל
        
        Args:
            to_address: כתובת יעד
            subject: נושא ההודעה
            body: תוכן ההודעה (טקסט או HTML)
            attachments: רשימת קבצים מצורפים (paths)
            cc_addresses: רשימת כתובות עותק
        
        Returns:
            True אם נשלח בהצלחה
        """
        if not self.is_connected:
            print("❌ Not connected to server")
            return False
        
        try:
            from exchangelib import Message, Mailbox, FileAttachment
            
            print(f"📤 Sending email to {to_address}...")
            
            # צור הודעה
            msg = Message(
                account=self.account,
                subject=subject,
                body=body,
                to_recipients=[Mailbox(email_address=to_address)]
            )
            
            # הוסף עותק אם יש
            if cc_addresses:
                msg.cc_recipients = [Mailbox(email_address=cc) for cc in cc_addresses]
            
            # הוסף קבצים מצורפים אם יש
            if attachments:
                for file_path in attachments:
                    if isinstance(file_path, str):
                        file_path = Path(file_path)
                    
                    if file_path.exists():
                        with open(file_path, 'rb') as f:
                            content = f.read()
                        
                        attachment = FileAttachment(
                            name=file_path.name,
                            content=content
                        )
                        msg.attach(attachment)
                        print(f"📎 Attached: {file_path.name}")
                    else:
                        print(f"⚠️ File not found: {file_path}")
            
            # שלח
            msg.send()
            
            print("✅ Email sent successfully!")
            return True
            
        except Exception as e:
            print(f"❌ Failed to send email: {e}")
            self.last_error = str(e)
            return False
    
    @staticmethod
    def _clean_filename(name: str) -> str:
        """ניקוי שם קובץ מתווים לא חוקיים"""
        invalid_chars = '<>:"/\\|?*'
        for char in invalid_chars:
            name = name.replace(char, '_')
        return name[:100]  # הגבל אורך


if __name__ == '__main__':
    # Test
    print("🧪 Testing EWS connection...")
    
    email = os.getenv('EMAIL_ADDRESS', 'yelena@algat.co.il')
    password = os.getenv('EMAIL_PASSWORD', '')
    
    connector = EWSEmailConnector(email, password)
    
    if connector.connect():
        print("✅ EWS connection successful!")
        folders = connector.list_folders()
        print(f"📁 Available folders: {folders[:5]}...")
        connector.disconnect()
    else:
        print(f"❌ Connection failed: {connector.last_error}")
