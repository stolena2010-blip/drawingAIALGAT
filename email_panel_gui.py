import os
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from pathlib import Path
from typing import Optional, List
from dotenv import set_key, dotenv_values
from datetime import datetime
import json

# Graph API
from src.services.email.graph_helper import GraphAPIHelper


class EmailPanel(ttk.Frame):
    """
    פאנל ניהול מיילים דרך EWS
    - הזנת פרטי התחברות (מייל, סיסמה, שרת)
    - בדיקת חיבור ורשימת תיקיות
    - הורדת קבצים מצורפים מתיקייה נבחרת
    - שמירת הגדרות לקובץ .env
    """

    def __init__(self, parent: tk.Toplevel) -> None:
        super().__init__(parent)
        self.parent = parent
        self.connector: Optional[GraphAPIHelper] = None
        self.folders: List[str] = []
        self.output_dir: Path = Path.cwd() / "email_downloads"
        self.config_file: Path = Path.cwd() / "email_config.json"

        # טען ברירות מחדל מה-ENV (לא מגלובל config כדי להימנע מ-validation מוקדמת)
        env = dotenv_values()
        # כתובת תיבה משותפת יחידה (ללא רשימת בחירה)
        self.shared_var = tk.StringVar(value=os.getenv("SHARED_MAILBOX_NAME", env.get("SHARED_MAILBOX_NAME", "quotes_check@algat.co.il")))
        # שיטה: GRAPH_API
        self.method_var = tk.StringVar(value="GRAPH_API")
        self.output_var = tk.StringVar(value=str(self.output_dir))
        self.show_inbox_only = tk.BooleanVar(value=True)

        self.status_var = tk.StringVar(value="מצב: מנותק")

        # טען כתובות מייל שמורות
        self._load_saved_emails()

        self._build_ui()

    def _build_ui(self) -> None:
        # מסגרת הפאנל עצמה (שימוש ב-padding של ttk.Frame)
        self.configure(padding=12)

        # פרטי התחברות
        auth_frame = ttk.LabelFrame(self, text="פרטי התחברות", padding=12)
        auth_frame.pack(fill=tk.X, expand=False)
        # הודעת כוונה: עבודה רק מול תיבה משותפת
        ttk.Label(auth_frame, text="האפליקציה עובדת רק מול תיבה משותפת.", foreground="#0066cc", font=("Arial", 9, "bold")).grid(row=0, column=0, columnspan=3, sticky=tk.W, padx=4, pady=(0, 8))

        ttk.Label(auth_frame, text="כתובת תיבה משותפת:").grid(row=1, column=0, sticky=tk.W, padx=4, pady=4)
        ttk.Entry(auth_frame, textvariable=self.shared_var, width=40).grid(row=1, column=1, sticky=tk.W, padx=4, pady=4)
        ttk.Label(auth_frame, text="שיטת חיבור:").grid(row=2, column=0, sticky=tk.W, padx=4, pady=4)
        method_combo = ttk.Combobox(auth_frame, textvariable=self.method_var, values=["OUTLOOK_COM"], state="readonly", width=37)
        method_combo.grid(row=2, column=1, sticky=tk.W, padx=4, pady=4)

        # אין שדות EWS כאשר עובדים עם Outlook COM בלבד

        btns_frame = ttk.Frame(auth_frame)
        btns_frame.grid(row=1, column=2, rowspan=3, sticky=tk.NS, padx=8)

        ttk.Button(btns_frame, text="בדוק חיבור", command=self._test_connect).pack(fill=tk.X, pady=4)
        ttk.Button(btns_frame, text="שמור הגדרות", command=self._save_env).pack(fill=tk.X, pady=4)

        ttk.Label(self, textvariable=self.status_var).pack(anchor=tk.W, pady=(6, 10))

        # רשימת תיקיות
        folders_frame = ttk.LabelFrame(self, text="תיקיות", padding=12)
        folders_frame.pack(fill=tk.BOTH, expand=True)
        self.folders_list = tk.Listbox(folders_frame, height=12)
        self.folders_list.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        folders_scroll = ttk.Scrollbar(folders_frame, orient=tk.VERTICAL, command=self.folders_list.yview)
        folders_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.folders_list.configure(yscrollcommand=folders_scroll.set)

        ttk.Button(self, text="רענן תיקיות", command=self._list_folders).pack(fill=tk.X, pady=4)
        ttk.Button(self, text="סנכרן + רענן", command=self._sync_and_refresh).pack(fill=tk.X, pady=0)
        ttk.Checkbutton(self, text="הצג רק תתי־תיקיות של Inbox", variable=self.show_inbox_only, command=self._list_folders).pack(anchor=tk.W)

        # פלט והורדה
        output_frame = ttk.LabelFrame(self, text="הורדה", padding=12)
        output_frame.pack(fill=tk.X, expand=False)

        ttk.Label(output_frame, text="תיקיית פלט:").grid(row=0, column=0, sticky=tk.W, padx=4, pady=4)
        ttk.Entry(output_frame, textvariable=self.output_var, width=40).grid(row=0, column=1, sticky=tk.W, padx=4, pady=4)
        ttk.Button(output_frame, text="בחר...", command=self._choose_output).grid(row=0, column=2, padx=4, pady=4)

        ttk.Button(output_frame, text="הורד מהתיקייה הנבחרת", command=self._download_selected).grid(row=1, column=0, columnspan=3, sticky=tk.EW, padx=4, pady=8)

        # ===== שליחה מתקייה מעובדת =====
        resend_frame = ttk.LabelFrame(self, text="שליחה מתקייה מעובדת", padding=12)
        resend_frame.pack(fill=tk.X, expand=False, pady=(8, 0))
        
        ttk.Label(resend_frame, text="בחר תקייה:").grid(row=0, column=0, sticky=tk.W, padx=4, pady=4)
        self.resend_folder_var = tk.StringVar(value="לא נבחרה")
        folder_label = ttk.Label(resend_frame, textvariable=self.resend_folder_var, foreground="#0066cc")
        folder_label.grid(row=0, column=1, sticky=tk.EW, padx=4, pady=4)
        ttk.Button(resend_frame, text="בחר תקייה", command=self._choose_resend_folder).grid(row=0, column=2, padx=4, pady=4)
        
        ttk.Label(resend_frame, text="קבצים:").grid(row=1, column=0, sticky=tk.W, padx=4, pady=4)
        self.resend_files_count_var = tk.StringVar(value="0")
        ttk.Label(resend_frame, textvariable=self.resend_files_count_var, foreground="#006600").grid(row=1, column=1, sticky=tk.W, padx=4, pady=4)
        
        ttk.Label(resend_frame, text="נמען חדש:").grid(row=2, column=0, sticky=tk.W, padx=4, pady=4)
        self.resend_recipient_var = tk.StringVar()
        ttk.Entry(resend_frame, textvariable=self.resend_recipient_var, width=30).grid(row=2, column=1, sticky=tk.EW, padx=4, pady=4)
        ttk.Button(resend_frame, text="💾", command=lambda: self._save_email_address('resend'), width=3).grid(row=2, column=2, padx=4, pady=4)
        
        ttk.Button(resend_frame, text="שלח מתקייה", command=self._resend_from_folder).grid(row=3, column=0, columnspan=3, sticky=tk.EW, padx=4, pady=8)
        
        resend_frame.columnconfigure(1, weight=1)

        # ===== שליחת מייל ידני =====
        send_frame = ttk.LabelFrame(self, text="שליחת מייל ידנית", padding=12)
        send_frame.pack(fill=tk.X, expand=False, pady=(8, 0))
        
        ttk.Label(send_frame, text="כתובת יעד:").grid(row=0, column=0, sticky=tk.W, padx=4, pady=4)
        self.to_email_var = tk.StringVar()
        ttk.Entry(send_frame, textvariable=self.to_email_var, width=30).grid(row=0, column=1, sticky=tk.W, padx=4, pady=4)
        ttk.Button(send_frame, text="💾", command=lambda: self._save_email_address('manual'), width=3).grid(row=0, column=2, padx=4, pady=4)
        
        ttk.Label(send_frame, text="נושא:").grid(row=1, column=0, sticky=tk.W, padx=4, pady=4)
        self.subject_var = tk.StringVar()
        ttk.Entry(send_frame, textvariable=self.subject_var, width=40).grid(row=1, column=1, sticky=tk.W, padx=4, pady=4)
        
        ttk.Label(send_frame, text="הודעה:").grid(row=2, column=0, sticky=tk.NW, padx=4, pady=4)
        self.body_text = tk.Text(send_frame, width=40, height=6, wrap=tk.WORD)
        self.body_text.grid(row=2, column=1, sticky=tk.W, padx=4, pady=4)
        
        ttk.Label(send_frame, text="קבצים מצורפים:").grid(row=3, column=0, sticky=tk.W, padx=4, pady=4)
        self.attachments_var = tk.StringVar()
        ttk.Entry(send_frame, textvariable=self.attachments_var, width=30, state="readonly").grid(row=3, column=1, sticky=tk.W, padx=4, pady=4)
        ttk.Button(send_frame, text="בחר קבצים", command=self._choose_attachments).grid(row=3, column=2, padx=4, pady=4)
        
        ttk.Button(send_frame, text="שלח מייל", command=self._send_email).grid(row=4, column=0, columnspan=3, sticky=tk.EW, padx=4, pady=8)

    def _ensure_connector(self) -> None:
        # Graph API
        if not self.connector:
            self.connector = GraphAPIHelper()

    def _choose_attachments(self) -> None:
        """בחירת קבצים מצורפים"""
        # טען כתובת שמורה אם השדה ריק
        if not self.to_email_var.get().strip() and hasattr(self, 'saved_manual_email') and self.saved_manual_email:
            self.to_email_var.set(self.saved_manual_email)
        
        files = filedialog.askopenfilenames(title="בחר קבצים מצורפים")
        if files:
            # סנן בחוץ ZIP files
            filtered_files = [f for f in files if not f.lower().endswith('.zip')]
            if len(filtered_files) < len(files):
                messagebox.showwarning("אזהרה", "קבצי ZIP לא יוצורפו למייל")
            # שמור את הרשימה כמשתנה instance
            self.selected_attachments = list(filtered_files)
            # הצג רק שמות קבצים
            names = [Path(f).name for f in filtered_files]
            self.attachments_var.set(", ".join(names) if len(names) <= 3 else f"{len(names)} קבצים")

    def _send_email(self) -> None:
        """שליחת מייל"""
        try:
            to_addr = self.to_email_var.get().strip()
            subject = self.subject_var.get().strip()
            body = self.body_text.get("1.0", tk.END).strip()
            
            if not to_addr:
                messagebox.showwarning("אזהרה", "אנא הזן כתובת יעד")
                return
            if not subject:
                messagebox.showwarning("אזהרה", "אנא הזן נושא")
                return
            
            # הוסף קידומת קבועה עם timestamp לכותרת
            import time
            timestamp = str(int(time.time()))  # Unix timestamp - מספר אחד
            subject = f"B2B_Quotation_{timestamp} | {subject}"
            
            # וודא חיבור
            if not self.connector:
                self._ensure_connector()
                if not self.connector.test_connection():
                    messagebox.showerror("שגיאה", "לא ניתן להתחבר")
                    return
            
            # שלח מייל
            attachments = getattr(self, 'selected_attachments', [])
            success = self.connector.send_email(
                to_address=to_addr,
                subject=subject,
                body=body,
                attachments=[Path(f) for f in attachments] if attachments else None,
                replace_display_with_filename=True
            )
            
            if success:
                messagebox.showinfo("הצלחה", f"המייל נשלח בהצלחה ל-{to_addr}")
                # נקה שדות
                self.to_email_var.set("")
                self.subject_var.set("")
                self.body_text.delete("1.0", tk.END)
                self.attachments_var.set("")
                self.selected_attachments = []
            else:
                messagebox.showerror("שגיאה", f"שליחת המייל נכשלה: {self.connector.last_error}")
                
        except Exception as e:
            messagebox.showerror("שגיאה", f"שגיאה בשליחת מייל: {e}")

    def _load_saved_emails(self) -> None:
        """טעינת כתובות מייל שמורות מקובץ config"""
        try:
            if self.config_file.exists():
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    # טען כתובות אם קיימות
                    if 'resend_email' in config:
                        self.saved_resend_email = config['resend_email']
                    else:
                        self.saved_resend_email = ""
                    
                    if 'manual_email' in config:
                        self.saved_manual_email = config['manual_email']
                    else:
                        self.saved_manual_email = ""
            else:
                self.saved_resend_email = ""
                self.saved_manual_email = ""
        except Exception as e:
            print(f"Failed to load saved addresses: {e}")
            self.saved_resend_email = ""
            self.saved_manual_email = ""

    def _save_email_address(self, email_type: str) -> None:
        """שמירת כתובת מייל לקובץ config"""
        try:
            # טען config קיים או צור חדש
            config = {}
            if self.config_file.exists():
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
            
            # עדכן כתובת לפי סוג
            if email_type == 'resend':
                email = self.resend_recipient_var.get().strip()
                if not email:
                    messagebox.showwarning("אזהרה", "אנא הזן כתובת מייל לפני השמירה")
                    return
                config['resend_email'] = email
                self.saved_resend_email = email
                messagebox.showinfo("הצלחה", f"כתובת נשמרה: {email}")
            elif email_type == 'manual':
                email = self.to_email_var.get().strip()
                if not email:
                    messagebox.showwarning("אזהרה", "אנא הזן כתובת מייל לפני השמירה")
                    return
                config['manual_email'] = email
                self.saved_manual_email = email
                messagebox.showinfo("הצלחה", f"כתובת נשמרה: {email}")
            
            # שמור לקובץ
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
                
        except Exception as e:
            messagebox.showerror("שגיאה", f"שגיאה בשמירת כתובת: {e}")

    def _test_connect(self) -> None:
        try:
            if not self.shared_var.get().strip():
                messagebox.showwarning("אזהרה", "אנא מלא/י כתובת תיבה משותפת.")
                return
            self._ensure_connector()
            ok = self.connector.test_connection()
            if ok:
                self.status_var.set("מצב: מחובר ✓")
                # רענן תיקיות אוטומטית לאחר חיבור מוצלח
                self._list_folders()
            else:
                self.status_var.set("מצב: שגיאת התחברות")
                messagebox.showerror("התחברות נכשלה", self.connector.last_error or "Unknown error")
        except Exception as e:
            self.status_var.set("מצב: שגיאה")
            messagebox.showerror("שגיאה", f"שגיאה בבדיקת חיבור: {e}")

    def _list_folders(self) -> None:
        try:
            if not self.shared_var.get().strip():
                messagebox.showwarning("אזהרה", "אנא מלא/י כתובת תיבה משותפת כדי לשלוף תיקיות.")
                return
            if not self.connector:
                self._ensure_connector()
                if not self.connector.test_connection():
                    messagebox.showerror("שגיאה", self.connector.last_error or "לא ניתן להתחבר")
                    return
                self.status_var.set("מצב: מחובר ✓")
            
            all_folders = self.connector.mailbox.list_folders(show_inbox_only=self.show_inbox_only.get())
            self.folders = [f['displayName'] for f in all_folders]
            
            self.folders_list.delete(0, tk.END)
            for folder in all_folders:
                folder_name = folder.get('displayName', '')
                folder_id = folder.get('id', '')
                
                # קבל מספר הודעות בתיקייה
                try:
                    messages = self.connector.mailbox.list_messages(folder_id, limit=1000)
                    msg_count = len(messages)
                    display_text = f"{folder_name} ({msg_count})"
                except Exception:
                    display_text = folder_name
                
                self.folders_list.insert(tk.END, display_text)
        except Exception as e:
            messagebox.showerror("שגיאה", f"שגיאה בשליפת תיקיות: {e}")

    def _sync_and_refresh(self) -> None:
        """Graph API syncs automatically - just test connection and refresh."""
        try:
            if not self.shared_var.get().strip():
                messagebox.showwarning("אזהרה", "אנא מלא/י כתובת תיבה משותפת כדי לשלוף תיקיות.")
                return
            self._ensure_connector()
            if not self.connector.test_connection():
                messagebox.showerror("שגיאה", self.connector.last_error or "לא ניתן להתחבר")
                return
            self.status_var.set("מצב: מחובר ✓ (מסונכרן)")
            self._list_folders()
        except Exception as e:
            messagebox.showerror("שגיאה", f"שגיאה בסנכרון: {e}")

    def _choose_output(self) -> None:
        folder = filedialog.askdirectory(title="בחר תיקיית פלט", initialdir=str(self.output_dir))
        if folder:
            self.output_var.set(folder)

    def _download_selected(self) -> None:
        try:
            selection = self.folders_list.curselection()
            if not selection:
                messagebox.showwarning("אזהרה", "בחר/י תיקייה מהרשימה.")
                return
            folder_display = self.folders_list.get(selection[0])
            
            # חלץ שם התיקייה בלי הספירה בסוגריים
            # "MANUAL (2)" -> "MANUAL"
            import re
            folder_name = re.sub(r'\s*\(\d+\)\s*$', '', folder_display).strip()

            if not self.connector:
                self._ensure_connector()
                if not self.connector.test_connection():
                    messagebox.showerror("שגיאה", self.connector.last_error or "לא ניתן להתחבר")
                    return
                self.status_var.set("מצב: מחובר ✓")

            out_dir = Path(self.output_var.get() or self.output_dir)
            out_dir.mkdir(parents=True, exist_ok=True)

            def _progress(current, total, msg) -> None:
                self.status_var.set(f"{msg} ({current}/{total})")
                self.update_idletasks()

            stats = self.connector.download_messages(
                folder_name=folder_name,
                output_dir=str(out_dir),
                progress_callback=_progress,
            )
            messagebox.showinfo("סיום", f"הורדה הסתיימה\nהודעות: {stats.get('total_messages', 0)}\nקבצים מצורפים: {stats.get('total_attachments', 0)}\nשגיאות: {stats.get('errors', 0)}")
        except Exception as e:
            messagebox.showerror("שגיאה", f"שגיאה בהורדה: {e}")

    def _save_env(self) -> None:
        try:
            env_path = Path.cwd() / ".env"
            # ודא שקובץ ה-.env קיים
            if not env_path.exists():
                env_path.touch()
            # עדכן ערכים
            # שמור שיטת חיבור קבועה
            set_key(str(env_path), "EMAIL_METHOD", self.method_var.get())
            set_key(str(env_path), "SHARED_MAILBOX", self.shared_var.get())
            messagebox.showinfo("נשמר", ".env עודכן בהצלחה")
        except Exception as e:
            messagebox.showerror("שגיאה", f"שגיאה בשמירת .env: {e}")

    def _choose_resend_folder(self) -> None:
        """בחר תקייה מעובדת לשליחה מחדש"""
        folder = filedialog.askdirectory(title="בחר תקייה עם קבצים לשליחה")
        if folder:
            self.resend_folder = Path(folder)
            self.resend_folder_var.set(folder)
            
            # ספור קבצים
            files = [f for f in self.resend_folder.glob("*") if f.is_file()]
            self.resend_files_count_var.set(f"{len(files)} קבצים")
            
            # טען כתובת שמורה אם קיימת
            if hasattr(self, 'saved_resend_email') and self.saved_resend_email:
                self.resend_recipient_var.set(self.saved_resend_email)

    def _resend_from_folder(self) -> None:
        """שלח קבצים מתקייה לנמען חדש"""
        try:
            # וולידציה
            if not hasattr(self, 'resend_folder') or not self.resend_folder:
                messagebox.showwarning("אזהרה", "אנא בחר תקייה תחילה")
                return
            
            recipient = self.resend_recipient_var.get().strip()
            if not recipient:
                messagebox.showwarning("אזהרה", "אנא הזן כתובת נמען")
                return
            
            # אסוף קבצים (בלי ZIP files)
            files = [f for f in self.resend_folder.glob("*") if f.is_file() and not f.suffix.lower() == '.zip']
            if not files:
                messagebox.showwarning("אזהרה", "אין קבצים בתקייה")
                return
            
            # בדוק חיבור
            if not self.connector:
                self._ensure_connector()
                if not self.connector.test_connection():
                    messagebox.showerror("שגיאה", "לא ניתן להתחבר")
                    return
            
            # שלח מייל
            import time
            timestamp = str(int(time.time()))  # Unix timestamp - מספר אחד
            
            # קרא כתובת מייל מ-email.txt בתקייה - מהשורה הראשונה
            email_from_file = ""
            email_file = self.resend_folder / "email.txt"
            if email_file.exists():
                try:
                    with open(email_file, "r", encoding="utf-8") as f:
                        first_line = f.readline().strip()
                        # השורה הראשונה היא הכתובת - נקה פורמט ישן אם קיים
                        if first_line and "@" in first_line:
                            # הסר "כתובת שולח:" או "From:" אם קיים
                            email_from_file = first_line.replace("כתובת שולח:", "").replace("From:", "").strip()
                except Exception as e:
                    print(f"GUI error ignored: {e}")
                    pass
            
            # בנה את הנושא עם כתובת המייל אם קיימת
            if email_from_file:
                subject = f"B2B_Quotation_{timestamp} | {email_from_file}"
            else:
                subject = f"B2B_Quotation_{timestamp}"
            
            success = self.connector.send_email(
                to_address=recipient,
                subject=subject,
                body="",  # גוף ריק
                attachments=files,
                replace_display_with_filename=True
            )
            
            if success:
                messagebox.showinfo(
                    "הצלחה",
                    f"✓ המייל נשלח בהצלחה\nל: {recipient}\nקבצים: {len(files)}"
                )
                # נקה שדות
                self.resend_recipient_var.set("")
                self.resend_folder_var.set("לא נבחרה")
                self.resend_files_count_var.set("0")
            else:
                messagebox.showerror(
                    "שגיאה",
                    f"❌ שליחת המייל נכשלה:\n{self.connector.last_error}"
                )
                
        except Exception as e:
            messagebox.showerror("שגיאה", f"שגיאה בשליחה: {e}")
