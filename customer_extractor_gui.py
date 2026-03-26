"""
Customer Extractor - GUI Version
גרסה עם ממשק גרפי לבחירת שלבים
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime
from pathlib import Path
import threading
import sys
import os
import customtkinter as ctk

# Import the main extractor functions
from customer_extractor_v3_dual import (
    scan_folder,
    DRAWING_EXTS
)

# Import email panel if available
try:
    from email_panel_gui import EmailPanel
    HAS_EMAIL_PANEL = True
except ImportError:
    HAS_EMAIL_PANEL = False

# Import dashboard
try:
    from dashboard_gui import DashboardWindow
    HAS_DASHBOARD = True
except ImportError:
    HAS_DASHBOARD = False

# Import automation panel if available
try:
    from automation_panel_gui import AutomationPanel
    from automation_runner import AutomationRunner
    HAS_AUTOMATION_PANEL = True
except ImportError:
    HAS_AUTOMATION_PANEL = False




class ExtractorGUI:
    def __init__(self, root) -> None:
        from src.utils.logger import setup_logging
        from pathlib import Path
        setup_logging(log_level="INFO", log_dir=Path("logs"))

        self.root = root
        self.root.title("Green Coat — DrawingAI Pro")
        self.root.geometry("650x850")
        self.root.resizable(True, True)
        
        # Variables
        self.folder_path = tk.StringVar()
        self.tosend_path = tk.StringVar()
        self.date_from = tk.StringVar()
        self.date_to = tk.StringVar()
        self.use_date_filter = tk.BooleanVar(value=False)
        self.stage1_var = tk.BooleanVar(value=True)
        self.stage2_var = tk.BooleanVar(value=True)
        self.stage3_var = tk.BooleanVar(value=True)
        self.stage4_var = tk.BooleanVar(value=True)
        self.stage5_var = tk.BooleanVar(value=True)
        self.stage6_var = tk.BooleanVar(value=True)
        self.stage7_var = tk.BooleanVar(value=True)
        self.stage8_var = tk.BooleanVar(value=True)
        self.recursive_var = tk.BooleanVar(value=True)
        self.enable_retry_var = tk.BooleanVar(value=True)  # ניסיונות נוספים - כברירת מחדל דלוק (V מסומן)
        self.stop_requested = False
        self.skip_requested = False

        self.automation_runner = None
        
        self.setup_ui()

        self._maybe_start_automation()
    
    def setup_ui(self) -> None:
        """בניית הממשק"""
        
        # Header
        header = ctk.CTkLabel(
            self.root,
            text="Green Coat — DrawingAI Pro — חילוץ נתונים משרטוטים",
            font=("Arial", 16, "bold"),
            fg_color="#1a1a2e",
            text_color="#00d4aa",
            height=50,
            corner_radius=0
        )
        header.pack(fill=tk.X)
        
        # Create scrollable main frame
        canvas_frame = ttk.Frame(self.root)
        canvas_frame.pack(fill=tk.BOTH, expand=True)
        
        canvas = tk.Canvas(canvas_frame, bg="white", highlightthickness=0)
        scrollbar = ttk.Scrollbar(canvas_frame, orient="vertical", command=canvas.yview)
        main_frame = ttk.Frame(canvas, padding="20")
        
        main_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=main_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Mouse wheel scrolling
        def _on_mousewheel(event) -> None:
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel)
        
        # ===== Folder Selection =====
        folder_frame = ttk.LabelFrame(main_frame, text="📁 תיקייה", padding="10")
        folder_frame.pack(fill=tk.X, pady=(0, 10))
        
        folder_entry = ttk.Entry(folder_frame, textvariable=self.folder_path, width=50)
        folder_entry.pack(side=tk.LEFT, padx=(0, 10), fill=tk.X, expand=True)
        
        browse_btn = ttk.Button(folder_frame, text="עיון...", command=self.browse_folder)
        browse_btn.pack(side=tk.LEFT)
        
        # Recursive checkbox
        recursive_cb = ttk.Checkbutton(
            folder_frame,
            text="כולל תת-תיקיות",
            variable=self.recursive_var
        )
        recursive_cb.pack(side=tk.LEFT, padx=(10, 0))
        
        # ===== TO_SEND Backup Path =====
        tosend_frame = ttk.LabelFrame(main_frame, text="💾 TO_SEND תיקיית עותק", padding="10")
        tosend_frame.pack(fill=tk.X, pady=(0, 10))
        
        tosend_entry = ttk.Entry(tosend_frame, textvariable=self.tosend_path, width=50)
        tosend_entry.pack(side=tk.LEFT, padx=(0, 10), fill=tk.X, expand=True)
        
        tosend_browse_btn = ttk.Button(tosend_frame, text="עיון...", command=self.browse_tosend_folder)
        tosend_browse_btn.pack(side=tk.LEFT)
        
        ttk.Label(tosend_frame, text="(תיקייה לשמירת עותקים עם סיומת _TO_SEND)", foreground="gray", font=("Arial", 8)).pack(anchor=tk.W, pady=(5, 0))
        
        # ===== Date Filter =====
        date_frame = ttk.LabelFrame(main_frame, text="📅 סינון לפי תאריך", padding="10")
        date_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Enable date filter checkbox
        date_enable_cb = ttk.Checkbutton(
            date_frame,
            text="הפעל סינון תאריכים",
            variable=self.use_date_filter,
            command=self.toggle_date_filter
        )
        date_enable_cb.pack(anchor=tk.W, pady=(0, 5))
        
        # Date range frame
        self.date_range_frame = ttk.Frame(date_frame)
        self.date_range_frame.pack(fill=tk.X)
        
        # From date
        ttk.Label(self.date_range_frame, text="מתאריך:").grid(row=0, column=0, padx=(0, 5), sticky=tk.W)
        self.date_from_entry = ttk.Entry(self.date_range_frame, textvariable=self.date_from, width=15)
        self.date_from_entry.grid(row=0, column=1, padx=(0, 20))
        self.date_from_entry.insert(0, "DD/MM/YYYY")
        self.date_from_entry.config(state='disabled')
        
        # To date
        ttk.Label(self.date_range_frame, text="עד תאריך:").grid(row=0, column=2, padx=(0, 5), sticky=tk.W)
        self.date_to_entry = ttk.Entry(self.date_range_frame, textvariable=self.date_to, width=15)
        self.date_to_entry.grid(row=0, column=3, padx=(0, 10))
        self.date_to_entry.insert(0, "DD/MM/YYYY")
        self.date_to_entry.config(state='disabled')
        
        # Help text
        self.date_help = ttk.Label(
            self.date_range_frame, 
            text="(סינון לפי תאריך שינוי קובץ)", 
            foreground="gray",
            font=("Arial", 8)
        )
        self.date_help.grid(row=0, column=4, padx=(0, 0), sticky=tk.W)
        
        # ===== Stage Selection =====
        stages_frame = ttk.LabelFrame(main_frame, text="⚙️ בחירת שלבים להרצה", padding="10")
        stages_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Info about preliminary stage and improvements
        info_frame = ttk.Frame(stages_frame)
        info_frame.pack(fill=tk.X, pady=(0, 10))
        ttk.Label(
            info_frame, 
            text="ℹ️  שלב 0: זיהוי אוטומטי של סוג קבצים + Layout + סיבוב תמונה - תמיד פעיל",
            font=("Arial", 9),
            foreground="gray"
        ).pack(anchor=tk.W)
        ttk.Label(
            info_frame, 
            text="🏢  מודל מתקדם לרפאל: זיהוי אוטומטי + Title Block מוגדל x3 לדיוק מקסימלי",
            font=("Arial", 9),
            foreground="blue"
        ).pack(anchor=tk.W)
        
        ttk.Label(stages_frame, text="בחר שלבי עיבוד לשרטוטים:", font=("Arial", 10, "bold")).pack(anchor=tk.W, pady=(0, 10))
        
        # Stage 1
        stage1_frame = ttk.Frame(stages_frame)
        stage1_frame.pack(fill=tk.X, pady=2)
        ttk.Checkbutton(stage1_frame, text="שלב 1:", variable=self.stage1_var).pack(side=tk.LEFT)
        ttk.Label(stage1_frame, text="מידע בסיסי (לקוח, מק״ט, שרטוט, רוויזיה)", foreground="blue").pack(side=tk.LEFT, padx=(5, 0))
        
        # Stage 2
        stage2_frame = ttk.Frame(stages_frame)
        stage2_frame.pack(fill=tk.X, pady=2)
        ttk.Checkbutton(stage2_frame, text="שלב 2:", variable=self.stage2_var).pack(side=tk.LEFT)
        ttk.Label(stage2_frame, text="תהליכי ייצור (חומר, ציפוי, צביעה, מפרטים)", foreground="green").pack(side=tk.LEFT, padx=(5, 0))
        
        # Stage 3
        stage3_frame = ttk.Frame(stages_frame)
        stage3_frame.pack(fill=tk.X, pady=2)
        ttk.Checkbutton(stage3_frame, text="שלב 3:", variable=self.stage3_var).pack(side=tk.LEFT)
        ttk.Label(stage3_frame, text="הנחיות מלאות (NOTES)", foreground="orange").pack(side=tk.LEFT, padx=(5, 0))
        
        # Stage 4
        stage4_frame = ttk.Frame(stages_frame)
        stage4_frame.pack(fill=tk.X, pady=2)
        ttk.Checkbutton(stage4_frame, text="שלב 4:", variable=self.stage4_var).pack(side=tk.LEFT)
        ttk.Label(stage4_frame, text="חישוב שטח גאומטרי", foreground="purple").pack(side=tk.LEFT, padx=(5, 0))
        
        # Stage 5
        stage5_frame = ttk.Frame(stages_frame)
        stage5_frame.pack(fill=tk.X, pady=2)
        ttk.Checkbutton(stage5_frame, text="שלב 5:", variable=self.stage5_var).pack(side=tk.LEFT)
        ttk.Label(stage5_frame, text="Fallback - תהליכים מטקסט (אם שלב 2 נכשל)", foreground="red").pack(side=tk.LEFT, padx=(5, 0))
        
        # ── Separator between drawing stages and post-processing stages ──
        ttk.Separator(stages_frame, orient='horizontal').pack(fill=tk.X, pady=(10, 5))
        ttk.Label(stages_frame, text="שלבי עיבוד נוספים (לאחר סיווג קבצים):", 
                  font=("Arial", 10, "bold")).pack(anchor=tk.W, pady=(5, 10))

        # Stage 6
        stage6_frame = ttk.Frame(stages_frame)
        stage6_frame.pack(fill=tk.X, pady=2)
        ttk.Checkbutton(stage6_frame, text="שלב 6:", variable=self.stage6_var).pack(side=tk.LEFT)
        ttk.Label(stage6_frame, text="חילוץ נתוני PL (Parts List)", 
                  foreground="#009688").pack(side=tk.LEFT, padx=(5, 0))

        # Stage 7
        stage7_frame = ttk.Frame(stages_frame)
        stage7_frame.pack(fill=tk.X, pady=2)
        ttk.Checkbutton(stage7_frame, text="שלב 7:", variable=self.stage7_var).pack(side=tk.LEFT)
        ttk.Label(stage7_frame, text="כמויות מתוך email (הזמנות/הצעות מחיר)", 
                  foreground="#795548").pack(side=tk.LEFT, padx=(5, 0))

        # Stage 8
        stage8_frame = ttk.Frame(stages_frame)
        stage8_frame.pack(fill=tk.X, pady=2)
        ttk.Checkbutton(stage8_frame, text="שלב 8:", variable=self.stage8_var).pack(side=tk.LEFT)
        ttk.Label(stage8_frame, text="פרטי פריט מהזמנות/הצעות (תיאור, כמות, מחיר)", 
                  foreground="#607D8B").pack(side=tk.LEFT, padx=(5, 0))

        # Select/Deselect All
        select_frame = ttk.Frame(stages_frame)
        select_frame.pack(fill=tk.X, pady=(10, 0))
        ttk.Button(select_frame, text="בחר הכל", command=self.select_all).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(select_frame, text="בטל הכל", command=self.deselect_all).pack(side=tk.LEFT)
        
        # ===== Advanced Options =====
        advanced_frame = ttk.LabelFrame(main_frame, text="🔧 אפשרויות מתקדמות", padding="10")
        advanced_frame.pack(fill=tk.X, pady=(0, 10))
        
        retry_cb = ttk.Checkbutton(
            advanced_frame,
            text="✓ אפשר 3 ניסיונות עם הגדלת תמונה ופילטרים (מצב V - מומלץ לדיוק מקסימלי)",
            variable=self.enable_retry_var
        )
        retry_cb.pack(anchor=tk.W)
        
        # Initialize confidence level variable
        self.confidence_level = tk.StringVar(value="LOW")
        
        ttk.Label(
            advanced_frame, 
            text="ℹ️  ניסיון 1: רגיל | ניסיון 2: הגדלה 50%×45% | ניסיון 3: הגדלה 70%×65% + פילטרים",
            font=("Arial", 8),
            foreground="gray"
        ).pack(anchor=tk.W, padx=(20, 0))
        
        # Confidence level selection
        confidence_frame = ttk.Frame(advanced_frame)
        confidence_frame.pack(fill=tk.X, pady=(15, 0))
        ttk.Label(confidence_frame, text="📊 בחר רמת ביטחון לקבצי B2B:", font=("Arial", 9, "bold")).pack(anchor=tk.W, pady=(0, 5))
        
        confidence_subframe = ttk.Frame(confidence_frame)
        confidence_subframe.pack(fill=tk.X, padx=(20, 0))
        
        ttk.Radiobutton(
            confidence_subframe,
            text="LOW - כל השורות (ברירת מחדל)",
            variable=self.confidence_level,
            value="LOW"
        ).pack(anchor=tk.W, pady=2)
        
        ttk.Radiobutton(
            confidence_subframe,
            text="MEDIUM - בינוני ומעלה (MEDIUM + HIGH + FULL)",
            variable=self.confidence_level,
            value="MEDIUM"
        ).pack(anchor=tk.W, pady=2)
        
        ttk.Radiobutton(
            confidence_subframe,
            text="HIGH - גבוה בלבד (HIGH + FULL)",
            variable=self.confidence_level,
            value="HIGH"
        ).pack(anchor=tk.W, pady=2)
        
        ttk.Label(
            confidence_frame,
            text="💾 *קבצי B2B מיוצרים בשלוש גרסאות - בחירה משפיעה על העתקה ל-TO_SEND",
            font=("Arial", 8),
            foreground="gray"
        ).pack(anchor=tk.W, pady=(10, 0))
        
        # ===== Progress =====
        self.progress_frame = ttk.LabelFrame(main_frame, text="📊 התקדמות", padding="10")
        self.progress_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.status_label = ttk.Label(self.progress_frame, text="מוכן להרצה", foreground="green")
        self.status_label.pack(anchor=tk.W)
        
        self.progress_bar = ctk.CTkProgressBar(self.progress_frame, mode='indeterminate')
        self.progress_bar.set(0)
        self.progress_bar.pack(fill=tk.X, pady=(5, 0))
        
        # ===== Buttons =====
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(10, 0))
        
        self.run_btn = ctk.CTkButton(
            button_frame,
            text="▶️ הפעל ניתוח",
            command=self.run_extraction,
            font=("Arial", 12, "bold"),
            fg_color="#00d4aa",
            hover_color="#00b894",
            text_color="#1a1a2e",
            height=40,
            corner_radius=10
        )
        self.run_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        self.skip_btn = ctk.CTkButton(
            button_frame,
            text="⏭️ דלג",
            command=self.skip_current,
            font=("Arial", 11, "bold"),
            fg_color="#9C27B0",
            hover_color="#7B1FA2",
            text_color="white",
            height=40,
            corner_radius=10,
            cursor="hand2",
            state="disabled"
        )
        self.skip_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        self.stop_btn = ctk.CTkButton(
            button_frame,
            text="🛑 עצור",
            command=self.stop_extraction,
            font=("Arial", 11, "bold"),
            fg_color="#FF9800",
            hover_color="#F57C00",
            text_color="white",
            height=40,
            corner_radius=10,
            state="disabled"
        )
        self.stop_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        ctk.CTkButton(
            button_frame,
            text="✖ סגור",
            command=self.root.quit,
            font=("Arial", 10),
            fg_color="#f44336",
            hover_color="#d32f2f",
            text_color="white",
            height=40,
            corner_radius=10
        ).pack(side=tk.LEFT)
        
        # הוסף כפתור למיילים אם EmailPanel זמין
        if HAS_EMAIL_PANEL:
            ctk.CTkButton(
                button_frame,
                text="📧 שלח מייל",
                command=self.open_email_window,
                font=("Arial", 12, "bold"),
                fg_color="#6c5ce7",
                hover_color="#5a4bd1",
                text_color="white",
                height=40,
                corner_radius=10
            ).pack(side=tk.RIGHT, padx=(10, 0))

        # הוסף כפתור אוטומציה אם זמין
        if HAS_AUTOMATION_PANEL:
            ctk.CTkButton(
                button_frame,
                text="🤖 אוטומציה",
                command=self.open_automation_window,
                font=("Arial", 12, "bold"),
                fg_color="#0984e3",
                hover_color="#0773c5",
                text_color="white",
                height=40,
                corner_radius=10
            ).pack(side=tk.RIGHT, padx=(10, 0))

        # הוסף כפתור Dashboard
        if HAS_DASHBOARD:
            ctk.CTkButton(
                button_frame,
                text="📊 Dashboard",
                command=lambda: DashboardWindow(self.root),
                font=("Arial", 12, "bold"),
                fg_color="#FF9800",
                hover_color="#F57C00",
                text_color="white",
                height=40,
                corner_radius=10
            ).pack(side=tk.RIGHT, padx=(10, 0))
        
        # Info label
        info_label = ttk.Label(
            main_frame,
            text="� Powered by GPT-4 Vision | 4 שלבים בלתי תלויים | חיסכון בזמן ועלות",
            foreground="#0066cc",
            font=("Arial", 9, "italic")
        )
        info_label.pack(pady=(10, 0))
        
    def browse_folder(self) -> None:
        """פתיחת דיאלוג לבחירת תיקייה"""
        folder = filedialog.askdirectory(title="בחר תיקייה")
        if folder:
            self.folder_path.set(folder)
    
    def browse_tosend_folder(self) -> None:
        """פתיחת דיאלוג לבחירת תיקיית TO_SEND"""
        folder = filedialog.askdirectory(title="בחר תיקייה ליעד TO_SEND")
        if folder:
            self.tosend_path.set(folder)
    
    def toggle_date_filter(self) -> None:
        """הפעלה/כיבוי של שדות התאריכים"""
        if self.use_date_filter.get():
            self.date_from_entry.config(state='normal')
            self.date_to_entry.config(state='normal')
            # Clear placeholder text
            if self.date_from.get() == "" or self.date_from.get() == "DD/MM/YYYY":
                self.date_from_entry.delete(0, tk.END)
            if self.date_to.get() == "" or self.date_to.get() == "DD/MM/YYYY":
                self.date_to_entry.delete(0, tk.END)
        else:
            self.date_from_entry.config(state='disabled')
            self.date_to_entry.config(state='disabled')
    
    def select_all(self) -> None:
        """בחירת כל השלבים"""
        self.stage1_var.set(True)
        self.stage2_var.set(True)
        self.stage3_var.set(True)
        self.stage4_var.set(True)
        self.stage5_var.set(True)
        self.stage6_var.set(True)
        self.stage7_var.set(True)
        self.stage8_var.set(True)
    
    def deselect_all(self) -> None:
        """ביטול בחירת כל השלבים"""
        self.stage1_var.set(False)
        self.stage2_var.set(False)
        self.stage3_var.set(False)
        self.stage4_var.set(False)
        self.stage5_var.set(False)
        self.stage6_var.set(False)
        self.stage7_var.set(False)
        self.stage8_var.set(False)
    
    def skip_current(self) -> None:
        """דילוג על הקובץ הנוכחי"""
        self.skip_requested = True
        self.status_label.config(text="⏭️ מדלג על קובץ נוכחי...", foreground="purple")
    
    def stop_extraction(self) -> None:
        """עצירת העיבוד ושמירת התוצאות"""
        self.stop_requested = True
        self.skip_requested = False  # ביטול דילוג אם לחצו עצור
        self.status_label.config(text="⏳ עוצר ושומר תוצאות...", foreground="orange")
        self.stop_btn.config(state="disabled")
        self.skip_btn.config(state="disabled")
    
    def validate_inputs(self) -> bool:
        """בדיקת תקינות הקלטים"""
        # Check folder
        if not self.folder_path.get():
            messagebox.showerror("שגיאה", "יש לבחור תיקייה")
            return False
        
        folder = Path(self.folder_path.get())
        if not folder.exists():
            messagebox.showerror("שגיאה", "התיקייה לא קיימת")
            return False
        
        # Check at least one stage selected
        if not any([self.stage1_var.get(), self.stage2_var.get(), 
                   self.stage3_var.get(), self.stage4_var.get(),
                   self.stage6_var.get(), self.stage7_var.get(),
                   self.stage8_var.get()]):
            messagebox.showerror("שגיאה", "יש לבחור לפחות שלב אחד")
            return False
        
        # Validate date range if filter enabled
        if self.use_date_filter.get():
            date_from_str = self.date_from.get()
            date_to_str = self.date_to.get()
            
            if not date_from_str or date_from_str == "DD/MM/YYYY":
                messagebox.showerror("שגיאה", "יש להזין תאריך התחלה")
                return False
            
            if not date_to_str or date_to_str == "DD/MM/YYYY":
                messagebox.showerror("שגיאה", "יש להזין תאריך סיום")
                return False
            
            try:
                date_from = datetime.strptime(date_from_str, "%d/%m/%Y")
                date_to = datetime.strptime(date_to_str, "%d/%m/%Y")
                
                if date_from > date_to:
                    messagebox.showerror("שגיאה", "תאריך התחלה חייב להיות לפני תאריך סיום")
                    return False
                    
            except ValueError:
                messagebox.showerror("שגיאה", "פורמט תאריך לא תקין (DD/MM/YYYY)")
                return False
        
        return True
    
    def run_extraction(self) -> None:
        """הרצת תהליך החילוץ"""
        if not self.validate_inputs():
            return
        
        # Disable button and enable stop button
        self.stop_requested = False
        self.skip_requested = False
        self.run_btn.config(state="disabled")
        self.stop_btn.config(state="normal")
        self.skip_btn.config(state="normal")
        self.status_label.config(text="מריץ...", foreground="blue")
        self.progress_bar.start()
        
        # Get selected stages
        selected_stages = {
            1: self.stage1_var.get(),
            2: self.stage2_var.get(),
            3: self.stage3_var.get(),
            4: self.stage4_var.get(),
            5: self.stage5_var.get(),
            6: self.stage6_var.get(),
            7: self.stage7_var.get(),
            8: self.stage8_var.get(),
        }
        
        # Get date filter range
        date_range = None
        if self.use_date_filter.get():
            date_from = datetime.strptime(self.date_from.get(), "%d/%m/%Y")
            date_to = datetime.strptime(self.date_to.get(), "%d/%m/%Y")
            # Set time to end of day for "to" date
            date_to = date_to.replace(hour=23, minute=59, second=59)
            date_range = (date_from, date_to)
        
        # Run in thread
        thread = threading.Thread(
            target=self._run_extraction_thread,
            args=(selected_stages, date_range),
            daemon=True
        )
        thread.start()
    
    def _run_extraction_thread(self, selected_stages, date_range) -> None:
        """הרצה ב-thread נפרד"""
        try:
            # Import here to avoid circular import
            import sys
            import os
            sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
            
            from customer_extractor_v3_dual import scan_folder
            import customer_extractor_v3_dual
            
            # Set global stop and skip functions
            customer_extractor_v3_dual._gui_should_stop = lambda: self.stop_requested
            customer_extractor_v3_dual._gui_should_skip = lambda: self.skip_requested
            
            # Reset function for skip flag
            def reset_skip() -> None:
                self.skip_requested = False
                self.root.after(0, lambda: self.status_label.config(text="מעבד...", foreground="blue"))
            customer_extractor_v3_dual._gui_skip_reset = reset_skip
            
            folder_path = Path(self.folder_path.get())
            recursive = self.recursive_var.get()
            enable_retry = self.enable_retry_var.get()
            confidence_selection = self.confidence_level.get()  # Get confidence level selection
            
            # Run scan_folder with selected stages and date range
            results, output_folder, output_path, cost_summary, file_classifications = scan_folder(
                folder_path, 
                recursive=recursive, 
                date_range=date_range,
                selected_stages=selected_stages,
                enable_image_retry=enable_retry,  # העברת דגל הניסיונות הנוספים
                tosend_folder=self.tosend_path.get() or None,  # העברת נתיב TO_SEND
                confidence_level=confidence_selection  # העברת בחירת רמת ביטחון B2B
            )
            
            # Prepare summary (works for both completed and stopped runs)
            summary_data = {
                'count': len(results),
                'folder': str(output_folder),
                'output_path': str(output_path) if output_path else None,
                'cost_summary': cost_summary
            }
            
            # Show success with summary (even if user stopped, results are returned)
            self.root.after(0, self._on_success, summary_data, results)
            
        except Exception as e:
            print(f"GUI error ignored: {e}")
            import traceback
            error_msg = f"{str(e)}\n\n{traceback.format_exc()}"
            self.root.after(0, self._on_error, error_msg)
    
    def _on_success(self, summary_data, results) -> None:
        """טיפול בהצלחה"""
        self.progress_bar.stop()
        self.run_btn.config(state="normal")
        self.stop_btn.config(state="disabled")
        self.skip_btn.config(state="disabled")
        
        count = summary_data['count']
        folder = summary_data['folder']
        output_path = summary_data.get('output_path')
        
        self.status_label.config(text=f"✅ הושלם! עובדו {count} קבצים", foreground="green")
        
        # Show results in beautiful window
        if results and output_path:
            self._show_results_window(results, output_path, summary_data)
        elif results:
            messagebox.showwarning("תוצאות", f"✅ עובדו {count} קבצים\n\n⚠️ לא נוצר קובץ Excel")
        else:
            messagebox.showinfo("תוצאות", "לא נמצאו קבצים לעיבוד")
    
    def _on_user_stopped(self) -> None:
        """טיפול בעצירה על ידי המשתמש"""
        self.progress_bar.stop()
        self.run_btn.config(state="normal")
        self.stop_btn.config(state="disabled")
        self.skip_btn.config(state="disabled")
        self.status_label.config(text="🛑 העיבוד נעצר על ידי המשתמש - הקבצים נשמרו", foreground="orange")
        messagebox.showinfo(
            "עצירה", 
            "🛑 העיבוד נעצר על ידי המשתמש\n\n"
            "✅ כל התוצאות עד כה נשמרו בקבצי Excel\n"
            "בדוק את תיקיית NEW FILES"
        )
    
    def _show_results_window(self, results, output_path, summary_data) -> None:
        """הצגת חלון תוצאות מעוצב - קריאה ישירה מהקבצים"""
        import pandas as pd
        from pathlib import Path
        
        # Get the folder path
        output_path_obj = Path(output_path)
        output_folder = output_path_obj.parent
        
        summary_file = output_path_obj  # This is the SUMMARY_all_results file
        
        # חפש קובץ זיהוי - חפש בכל הקבצים בתיקייה
        classification_file = None
        
        # ראשון: חפש קובץ SUMMARY_all_classifications
        for file in output_folder.glob("*"):
            if file.is_file() and file.name.startswith("SUMMARY_all_classifications_") and file.suffix == ".xlsx":
                # בחר את הקובץ האחרון שנוצר
                if classification_file is None or file.stat().st_mtime > classification_file.stat().st_mtime:
                    classification_file = file
        
        # אם לא נמצא, חפש קובצי זיהוי ישנים
        if not classification_file:
            for file in output_folder.glob("file_classification_*.xlsx"):
                if classification_file is None or file.stat().st_mtime > classification_file.stat().st_mtime:
                    classification_file = file
        
        # Read both files
        try:
            # Read classification file for file statistics
            if classification_file and classification_file.exists():
                df_classification = pd.read_excel(classification_file)
            else:
                df_classification = None
            
            # Read SUMMARY file for drawing statistics
            df_summary = pd.read_excel(summary_file)
        except Exception as e:
            messagebox.showerror("שגיאה", f"לא ניתן לקרוא את הקבצים:\n{e}")
            return
        
        # Create new window
        results_win = tk.Toplevel(self.root)
        results_win.title("📊 תוצאות ניתוח - Green Coat DrawingAI Pro")
        results_win.geometry("900x750")
        results_win.resizable(True, True)
        
        # Header
        header = tk.Label(
            results_win,
            text="✅ החילוץ הושלם בהצלחה!",
            font=("Arial", 16, "bold"),
            bg="#4CAF50",
            fg="white",
            pady=15
        )
        header.pack(fill=tk.X)
        
        # Main content frame with scrollbar
        main_frame = ttk.Frame(results_win)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Canvas for scrolling
        canvas = tk.Canvas(main_frame)
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Pack canvas and scrollbar FIRST
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Helper function for non-empty count
        non_empty = lambda col: sum(1 for v in df_summary[col] if v and str(v).strip() and str(v).lower() != 'nan') if col in df_summary.columns else 0
        
        # === FOLDER STATISTICS SECTION ===
        # Get folder statistics from summary_data
        folder_stats = summary_data.get('cost_summary', {}).get('folder_stats', [])
        
        if folder_stats:
            folders_section = ttk.LabelFrame(scrollable_frame, text="📬 Processed Emails Summary", padding="15")
            folders_section.pack(fill=tk.X, pady=(0, 15))
            
            # Header with total folders count
            tk.Label(
                folders_section,
                text=f"📊 Total Folders (Emails): {len(folder_stats)}",
                font=("Arial", 12, "bold"),
                fg="#0066cc"
            ).pack(anchor="w", pady=(0, 10))
            
            # Create a frame with table-like display for each folder
            table_frame = ttk.Frame(folders_section)
            table_frame.pack(fill=tk.BOTH, expand=True)
            
            # Headers
            header_frame = tk.Frame(table_frame, bg="#2c3e50", relief=tk.RAISED, borderwidth=1)
            header_frame.pack(fill=tk.X, pady=(0, 1))
            
            tk.Label(header_frame, text="Folder Name", font=("Arial", 9, "bold"), bg="#2c3e50", fg="white", width=16, anchor="w").pack(side=tk.LEFT, padx=2, pady=4)
            tk.Label(header_frame, text="Drawings", font=("Arial", 9, "bold"), bg="#2c3e50", fg="white", width=8, anchor="center").pack(side=tk.LEFT, padx=2, pady=4)
            tk.Label(header_frame, text="Other Files", font=("Arial", 9, "bold"), bg="#2c3e50", fg="white", width=18, anchor="w").pack(side=tk.LEFT, padx=2, pady=4)
            tk.Label(header_frame, text="High", font=("Arial", 9, "bold"), bg="#2c3e50", fg="white", width=5, anchor="center").pack(side=tk.LEFT, padx=2, pady=4)
            tk.Label(header_frame, text="Med", font=("Arial", 9, "bold"), bg="#2c3e50", fg="white", width=5, anchor="center").pack(side=tk.LEFT, padx=2, pady=4)
            tk.Label(header_frame, text="Low", font=("Arial", 9, "bold"), bg="#2c3e50", fg="white", width=5, anchor="center").pack(side=tk.LEFT, padx=2, pady=4)
            tk.Label(header_frame, text="Cost ($)", font=("Arial", 9, "bold"), bg="#2c3e50", fg="white", width=8, anchor="center").pack(side=tk.LEFT, padx=2, pady=4)
            tk.Label(header_frame, text="Time (s)", font=("Arial", 9, "bold"), bg="#2c3e50", fg="white", width=8, anchor="center").pack(side=tk.LEFT, padx=2, pady=4)
            
            # Data rows
            for folder_info in folder_stats:
                folder_name = folder_info['name']
                file_types = folder_info.get('file_types', {})
                total_drawings = folder_info.get('total_drawings', 0)
                conf_high = folder_info.get('confidence_high', 0)
                conf_medium = folder_info.get('confidence_medium', 0)
                conf_low = folder_info.get('confidence_low', 0)
                total_cost = folder_info.get('total_cost', 0)
                processing_time = folder_info.get('processing_time', 0)
                
                # Determine row color based on confidence levels
                # Green if all drawings have high confidence, red otherwise
                if total_drawings > 0:
                    if conf_high == total_drawings:
                        row_bg = "#c8e6c9"  # Light green
                    else:
                        row_bg = "#ffcdd2"  # Light red
                else:
                    row_bg = "#ffffff"  # White if no drawings
                
                row_frame = tk.Frame(table_frame, bg=row_bg, relief=tk.SOLID, borderwidth=1)
                row_frame.pack(fill=tk.X, pady=1)
                
                # Folder name
                tk.Label(row_frame, text=folder_name, font=("Courier", 8), bg=row_bg, width=16, anchor="w").pack(side=tk.LEFT, padx=2, pady=3)
                
                # Drawings count
                tk.Label(row_frame, text=str(total_drawings), font=("Arial", 9, "bold"), bg=row_bg, fg="#1976D2", width=8, anchor="center").pack(side=tk.LEFT, padx=2, pady=3)
                
                # Other file types (non-drawings)
                other_files_parts = []
                for ftype, count in sorted(file_types.items(), key=lambda x: -x[1]):
                    if ftype != 'DRAWING':
                        # Shorten type names for display
                        display_name = {
                            'PURCHASE_ORDER': 'Order',
                            'QUOTE': 'Quote',
                            'INVOICE': 'Invoice',
                            'PARTS_LIST': 'PL',
                            '3D_MODEL': '3D',
                            '3D_IMAGE': '3D Img',
                            'OTHER': 'Other'
                        }.get(ftype, ftype)
                        other_files_parts.append(f"{count} {display_name}")
                
                other_files_str = ", ".join(other_files_parts) if other_files_parts else "-"
                tk.Label(row_frame, text=other_files_str, font=("Arial", 8), bg=row_bg, width=18, anchor="w").pack(side=tk.LEFT, padx=2, pady=3)
                
                # Confidence levels in separate columns
                tk.Label(row_frame, text=str(conf_high), font=("Arial", 9, "bold"), bg=row_bg, fg="#4CAF50", width=5, anchor="center").pack(side=tk.LEFT, padx=2, pady=3)
                tk.Label(row_frame, text=str(conf_medium), font=("Arial", 9, "bold"), bg=row_bg, fg="#FF9800", width=5, anchor="center").pack(side=tk.LEFT, padx=2, pady=3)
                tk.Label(row_frame, text=str(conf_low), font=("Arial", 9, "bold"), bg=row_bg, fg="#F44336", width=5, anchor="center").pack(side=tk.LEFT, padx=2, pady=3)
                
                # Cost column
                cost_str = f"${total_cost:.4f}" if total_cost > 0 else "$0.0000"
                tk.Label(row_frame, text=cost_str, font=("Arial", 8, "bold"), bg=row_bg, fg="#9C27B0", width=8, anchor="center").pack(side=tk.LEFT, padx=2, pady=3)
                
                # Time column
                time_str = f"{processing_time:.1f}" if processing_time > 0 else "0.0"
                tk.Label(row_frame, text=time_str, font=("Arial", 8, "bold"), bg=row_bg, fg="#607D8B", width=8, anchor="center").pack(side=tk.LEFT, padx=2, pady=3)
        
        # === File Info Section ===
        file_section = ttk.LabelFrame(scrollable_frame, text="📁 קבצי התוצאות", padding="15")
        file_section.pack(fill=tk.X, pady=(0, 15))
        
        # File 1: Summary Results (Drawings)
        tk.Label(
            file_section,
            text=f"📊 קובץ SUMMARY (שרטוטים):",
            font=("Arial", 10, "bold"),
            anchor="w"
        ).pack(fill=tk.X)
        
        file1_frame = ttk.Frame(file_section)
        file1_frame.pack(fill=tk.X, pady=(5, 10))
        
        path_entry1 = ttk.Entry(file1_frame, font=("Courier", 8), width=60)
        path_entry1.insert(0, str(summary_file))
        path_entry1.config(state="readonly")
        path_entry1.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        
        def open_summary_file() -> None:
            import subprocess
            try:
                if summary_file.exists():
                    subprocess.run(['start', '', str(summary_file)], shell=True)
                else:
                    messagebox.showerror("שגיאה", f"קובץ לא קיים:\n{summary_file}")
            except Exception as e:
                messagebox.showerror("שגיאה", f"לא ניתן לפתוח את הקובץ:\n{e}")
        
        ttk.Button(
            file1_frame,
            text="🔓 פתח",
            command=open_summary_file,
            width=8
        ).pack(side=tk.LEFT)
        
        # File 2: Classification File
        if df_classification is not None and classification_file:
            tk.Label(
                file_section,
                text=f"📋 קובץ זיהוי קבצים:",
                font=("Arial", 10, "bold"),
                anchor="w"
            ).pack(fill=tk.X, pady=(10, 0))
            
            file2_frame = ttk.Frame(file_section)
            file2_frame.pack(fill=tk.X, pady=(5, 0))
            
            path_entry2 = ttk.Entry(file2_frame, font=("Courier", 8), width=60)
            path_entry2.insert(0, str(classification_file))
            path_entry2.config(state="readonly")
            path_entry2.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
            
            def open_classification_file() -> None:
                import subprocess
                try:
                    if classification_file.exists():
                        subprocess.run(['start', '', str(classification_file)], shell=True)
                    else:
                        messagebox.showerror("שגיאה", f"קובץ לא קיים:\n{classification_file}")
                except Exception as e:
                    messagebox.showerror("שגיאה", f"לא ניתן לפתוח את הקובץ:\n{e}")
            
            ttk.Button(
                file2_frame,
                text="🔓 פתח",
                command=open_classification_file,
                width=8
            ).pack(side=tk.LEFT)
        
        # === DRAWINGS Statistics Section ===
        # Main section header with blue background
        stage1_header = tk.Label(
            scrollable_frame,
            text="🔵 שלבים 1-4: עיבוד שרטוטים",
            font=("Arial", 13, "bold"),
            bg="#2196F3",
            fg="white",
            pady=8,
            anchor="w",
            padx=10
        )
        stage1_header.pack(fill=tk.X, pady=(0, 10))
        
        drawings_section = ttk.LabelFrame(scrollable_frame, text="📊 סיכום שרטוטים", padding="15")
        drawings_section.pack(fill=tk.X, pady=(0, 15))
        
        total_drawings = len(df_summary)
        
        tk.Label(
            drawings_section,
            text=f"📊 סה״כ שרטוטים שעובדו: {total_drawings}",
            font=("Arial", 12, "bold"),
            fg="#0066cc"
        ).pack(anchor="w", pady=(0, 10))
        
        # Create grid of statistics - STAGE 1: BASIC FIELDS
        stats_frame1 = ttk.LabelFrame(drawings_section, text="📋 שלב 1: שדות בסיסיים וזיהוי", padding="10")
        stats_frame1.pack(fill=tk.X, pady=(0, 10))
        
        stats_data_stage1 = [
            ("👤 לקוח", 'customer_name', "#2196F3"),
            ("🔧 מק״ט", 'part_number', "#FF9800"),
            ("📦 שם פריט", 'item_name', "#8BC34A"),
            ("📋 מס׳ שרטוט", 'drawing_number', "#4CAF50"),
            ("📝 רוויזיה", 'revision', "#9C27B0"),
            ("⚙️ חומר גלם", 'material', "#F44336"),
        ]
        
        for row, (label, col, color) in enumerate(stats_data_stage1):
            count_val = non_empty(col)
            percentage = (count_val / total_drawings * 100) if total_drawings > 0 else 0
            
            tk.Label(
                stats_frame1,
                text=f"{label}:",
                font=("Arial", 10),
                anchor="w",
                width=15
            ).grid(row=row, column=0, sticky="w", padx=5, pady=2)
            
            tk.Label(
                stats_frame1,
                text=f"{count_val}/{total_drawings}",
                font=("Arial", 10, "bold"),
                fg=color,
                anchor="e",
                width=10
            ).grid(row=row, column=1, sticky="e", padx=5, pady=2)
            
            tk.Label(
                stats_frame1,
                text=f"({percentage:.0f}%)",
                font=("Arial", 9),
                fg="gray",
                anchor="w",
                width=8
            ).grid(row=row, column=2, sticky="w", padx=5, pady=2)
        
        # STAGE 2: PROCESSES & SPECIFICATIONS
        stats_frame2 = ttk.LabelFrame(drawings_section, text="🛠️ שלב 2: תהליכים ומפרטים", padding="10")
        stats_frame2.pack(fill=tk.X, pady=(0, 10))
        
        stats_data_stage2 = [
            ("🎨 ציפוי", 'coating_processes', "#00BCD4"),
            ("🖌️ צביעה", 'painting_processes', "#FFC107"),
            ("🌈 צבעים", 'colors', "#E91E63"),
            ("📜 מפרטים", 'specifications', "#673AB7"),
            ("📝 NOTES מלא", 'notes_full_text', "#607D8B"),
            ("📄 תקציר תהליכים", 'process_summary_hebrew', "#FF5722"),
        ]
        
        for row, (label, col, color) in enumerate(stats_data_stage2):
            count_val = non_empty(col)
            percentage = (count_val / total_drawings * 100) if total_drawings > 0 else 0
            
            tk.Label(
                stats_frame2,
                text=f"{label}:",
                font=("Arial", 10),
                anchor="w",
                width=15
            ).grid(row=row, column=0, sticky="w", padx=5, pady=2)
            
            tk.Label(
                stats_frame2,
                text=f"{count_val}/{total_drawings}",
                font=("Arial", 10, "bold"),
                fg=color,
                anchor="e",
                width=10
            ).grid(row=row, column=1, sticky="e", padx=5, pady=2)
            
            tk.Label(
                stats_frame2,
                text=f"({percentage:.0f}%)",
                font=("Arial", 9),
                fg="gray",
                anchor="w",
                width=8
            ).grid(row=row, column=2, sticky="w", padx=5, pady=2)
        
        # CONFIDENCE & QUALITY
        quality_section = ttk.LabelFrame(scrollable_frame, text="✨ איכות ובטחון", padding="15")
        quality_section.pack(fill=tk.X, pady=(0, 15))
        
        # Count by confidence level - CHECK ALL POSSIBLE VALUES
        if 'confidence_level' in df_summary.columns:
            # Debug: print unique values to understand what we have
            unique_values = df_summary['confidence_level'].unique()
            
            confidence_counts = df_summary['confidence_level'].value_counts()
            
            confidence_frame = ttk.Frame(quality_section)
            confidence_frame.pack(fill=tk.X, pady=(0, 10))
            
            confidence_levels = [
                ("🟢 ביטחון גבוה", "HIGH", "#4CAF50"),
                ("🟡 ביטחון בינוני", "MEDIUM", "#FFC107"),
                ("🔴 ביטחון נמוך", "LOW", "#F44336")
            ]
            
            for idx, (label, level, color) in enumerate(confidence_levels):
                # Count all variations (case-insensitive)
                count_val = sum(1 for v in df_summary['confidence_level'] 
                              if str(v).strip().upper() == level)
                percentage = (count_val / total_drawings * 100) if total_drawings > 0 else 0
                
                tk.Label(
                    confidence_frame,
                    text=f"{label}:",
                    font=("Arial", 10),
                    anchor="w",
                    width=20
                ).grid(row=idx, column=0, sticky="w", padx=5, pady=2)
                
                tk.Label(
                    confidence_frame,
                    text=f"{count_val}/{total_drawings}",
                    font=("Arial", 10, "bold"),
                    fg=color,
                    anchor="e",
                    width=10
                ).grid(row=idx, column=1, sticky="e", padx=5, pady=2)
                
                tk.Label(
                    confidence_frame,
                    text=f"({percentage:.0f}%)",
                    font=("Arial", 9),
                    fg="gray",
                    anchor="w",
                    width=8
                ).grid(row=idx, column=2, sticky="w", padx=5, pady=2)
        
        # Needs review
        if 'needs_review' in df_summary.columns:
            needs_review_count = sum(1 for v in df_summary['needs_review'] if str(v).strip() and str(v).lower() not in ['', 'nan', 'none'])
            tk.Label(
                quality_section,
                text=f"⚠️ שרטוטים הדורשים ביקורת: {needs_review_count}/{total_drawings}",
                font=("Arial", 10),
                fg="#FF5722"
            ).pack(anchor="w", pady=(10, 0))
        
        # === DRAWING PROCESSING COST & TIME INFO ===
        cost_drawings_section = ttk.LabelFrame(scrollable_frame, text="💰 עלויות וביצועים - שלבים 1-4 (עיבוד שרטוטים)", padding="15")
        cost_drawings_section.pack(fill=tk.X, pady=(0, 15))
        
        # Calculate total costs from SUMMARY file
        total_extraction_cost = 0
        total_execution_time = 0
        
        if 'extraction_cost_usd' in df_summary.columns:
            total_extraction_cost = df_summary['extraction_cost_usd'].sum()
        
        if 'execution_time_seconds' in df_summary.columns:
            total_execution_time = df_summary['execution_time_seconds'].sum()
        
        # Get token info from summary_data if available
        total_tokens_in = 0
        total_tokens_out = 0
        total_cost_with_classification = 0
        classification_cost = 0
        
        if summary_data and 'cost_summary' in summary_data:
            cost_sum = summary_data['cost_summary']
            total_tokens_in = cost_sum.get('input_tokens', 0)
            total_tokens_out = cost_sum.get('output_tokens', 0)
            total_cost_with_classification = cost_sum.get('total_cost_with_classification', 0)
            classification_cost = cost_sum.get('classification_cost', 0)
        
        cost_drawings_text = tk.Text(cost_drawings_section, height=13, font=("Courier", 10), bg="#e8f5e9", relief=tk.FLAT)
        cost_drawings_text.pack(fill=tk.X)
        
        if total_extraction_cost > 0 or total_tokens_in > 0:
            cost_drawings_text.insert(tk.END, "💵 עלויות עיבוד שרטוטים:\n")
            cost_drawings_text.insert(tk.END, "=" * 50 + "\n")
            
            # Token details
            input_cost = (total_tokens_in / 1_000_000) * 2.50
            output_cost = (total_tokens_out / 1_000_000) * 10.00
            cost_drawings_text.insert(tk.END, f"📥 Input tokens:       {total_tokens_in:,} (${input_cost:.4f})\n")
            cost_drawings_text.insert(tk.END, f"📤 Output tokens:      {total_tokens_out:,} (${output_cost:.4f})\n")
            cost_drawings_text.insert(tk.END, f"📊 סה\"כ tokens:       {total_tokens_in + total_tokens_out:,}\n")
            cost_drawings_text.insert(tk.END, f"💰 עלות עיבוד:        ${total_extraction_cost:.4f} USD (₪{total_extraction_cost*3.7:.2f})\n")
            
            avg_per_drawing = total_extraction_cost / total_drawings if total_drawings > 0 else 0
            cost_drawings_text.insert(tk.END, f"📊 ממוצע לשרטוט:      ${avg_per_drawing:.6f} USD\n")
            
            # Time details
            if total_execution_time > 0:
                minutes = int(total_execution_time // 60)
                seconds = total_execution_time % 60
                cost_drawings_text.insert(tk.END, f"\n⏱️  זמן עיבוד כולל:     {minutes}:{seconds:05.2f} ({total_execution_time:.1f} שניות)\n")
                avg_time = total_execution_time / total_drawings if total_drawings > 0 else 0
                cost_drawings_text.insert(tk.END, f"⏱️  ממוצע לשרטוט:       {avg_time:.2f} שניות\n")
            
            # Total cost summary
            if total_cost_with_classification > 0 and classification_cost > 0:
                cost_drawings_text.insert(tk.END, f"\n" + "=" * 50 + "\n")
                cost_drawings_text.insert(tk.END, f"💰 עלות זיהוי (שלב 0):  ${classification_cost:.4f} USD\n")
                cost_drawings_text.insert(tk.END, f"💰 עלות עיבוד (1-4):    ${total_extraction_cost:.4f} USD\n")
                cost_drawings_text.insert(tk.END, f"💵 סה\"כ כולל:          ${total_cost_with_classification:.4f} USD (₪{total_cost_with_classification*3.7:.2f})\n")
                
                # Add classification time info
                classification_time = summary_data.get('cost_summary', {}).get('classification_time', 0)
                classification_folder_count = summary_data.get('cost_summary', {}).get('classification_folder_count', 0)
                avg_classification_time = summary_data.get('cost_summary', {}).get('avg_classification_time_per_folder', 0)
                
                if classification_time > 0:
                    cost_drawings_text.insert(tk.END, f"\n⏱️  זמן זיהוי כולל:      {classification_time:.1f} שניות\n")
                    if classification_folder_count > 0:
                        cost_drawings_text.insert(tk.END, f"⏱️  ממוצע פר תת-תקייה:   {avg_classification_time:.2f} שניות\n")
        else:
            cost_drawings_text.insert(tk.END, "⚠️  מידע על עלויות לא זמין\n")
        
        cost_drawings_text.config(state=tk.DISABLED)
        
        # Bottom buttons frame
        buttons_frame = ttk.Frame(results_win)
        buttons_frame.pack(fill=tk.X, padx=20, pady=(0, 20))
        
        # Open folder button
        def open_folder() -> None:
            import subprocess
            subprocess.run(['explorer', str(output_folder)])
        
        ttk.Button(
            buttons_frame,
            text="📂 פתח תיקייה",
            command=open_folder
        ).pack(side=tk.LEFT, padx=5)
        
        # Close button
        ttk.Button(
            buttons_frame,
            text="סגור",
            command=results_win.destroy
        ).pack(side=tk.RIGHT, padx=5)
    
    def _on_error(self, error) -> None:
        """טיפול בשגיאה"""
        self.progress_bar.stop()
        self.run_btn.config(state="normal")
        self.status_label.config(text="❌ שגיאה", foreground="red")
        messagebox.showerror("שגיאה", f"התרחשה שגיאה:\n{error}")
    
    def open_email_window(self) -> None:
        """פתח חלון נפרד לניהול מיילים"""
        if not HAS_EMAIL_PANEL:
            messagebox.showerror("שגיאה", "פאנל המיילים לא זמין.\nהעתק את email_panel_gui.py לתיקייה הראשית.")
            return
        
        email_window = ctk.CTkToplevel(self.root)
        email_window.title("📧 הורדת וניהול מיילים — Green Coat")
        email_window.geometry("900x800")
        email_window.resizable(True, True)
        
        try:
            panel = EmailPanel(email_window)
            panel.pack(fill=tk.BOTH, expand=True)
        except Exception as e:
            messagebox.showerror("שגיאה", f"שגיאה בעת פתיחת פאנל המיילים:\n{e}")
            email_window.destroy()

    def open_automation_window(self) -> None:
        """פתח חלון נפרד להגדרות אוטומציה"""
        if not HAS_AUTOMATION_PANEL:
            messagebox.showerror("שגיאה", "פאנל האוטומציה לא זמין.\nהעתק את automation_panel_gui.py לתיקייה הראשית.")
            return

        automation_window = ctk.CTkToplevel(self.root)
        automation_window.title("אוטומציה - Green Coat DrawingAI Pro")
        automation_window.geometry("1100x850")
        automation_window.minsize(900, 700)
        automation_window.resizable(True, True)

        try:
            panel = AutomationPanel(automation_window)
            panel.pack(fill=tk.BOTH, expand=True)
        except Exception as e:
            messagebox.showerror("שגיאה", f"שגיאה בעת פתיחת פאנל האוטומציה:\n{e}")
            automation_window.destroy()

    def _maybe_start_automation(self) -> None:
        """הפעל אוטומציה אוטומטית אם הוגדרה בקובץ ההגדרות"""
        if not HAS_AUTOMATION_PANEL:
            return

        try:
            config_path = Path.cwd() / "automation_config.json"
            if not config_path.exists():
                return

            import json
            with open(config_path, "r", encoding="utf-8") as f:
                cfg = json.load(f)

            if not cfg.get("auto_start", False):
                return

            state_path = Path.cwd() / "automation_state.json"
            self.automation_runner = AutomationRunner(config_path, state_path, self._set_status_from_automation)
            self.automation_runner.start()
            self.status_label.config(text="אוטומציה פעילה ✓", foreground="green")
        except Exception as e:
            print(f"GUI error ignored: {e}")
            pass

    def _set_status_from_automation(self, text: str) -> None:
        try:
            self.status_label.config(text=text, foreground="blue")
        except Exception as e:
            print(f"GUI error ignored: {e}")
            pass
            email_window.destroy()


def main() -> None:
    from dotenv import load_dotenv
    load_dotenv()
    
    ctk.set_appearance_mode("dark")
    ctk.set_default_color_theme("green")
    root = ctk.CTk()
    root.title("Green Coat — DrawingAI Pro")
    root.minsize(750, 850)
    app = ExtractorGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()

