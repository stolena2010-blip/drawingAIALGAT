#!/usr/bin/env python3
"""
DrawingAI Pro - Main Entry Point
Automation Panel as primary window, Manual GUI available via button.
"""
import tkinter as tk
from tkinter import messagebox
from pathlib import Path

from dotenv import load_dotenv
load_dotenv()

from src.utils.logger import setup_logging

try:
    import customtkinter as ctk
    HAS_CTK = True
except ImportError:
    HAS_CTK = False

from automation_panel_gui import AutomationPanel


def open_manual_gui(parent):
    """Open the manual extraction GUI as a secondary window."""
    try:
        from customer_extractor_gui import ExtractorGUI
        
        manual_window = ctk.CTkToplevel(parent) if HAS_CTK else tk.Toplevel(parent)
        manual_window.title("🔧 DrawingAI Pro — עיבוד ידני")
        manual_window.geometry("900x900")
        manual_window.minsize(750, 850)
        manual_window.resizable(True, True)
        
        app = ExtractorGUI(manual_window)
        
    except Exception as e:
        messagebox.showerror("שגיאה", f"שגיאה בפתיחת GUI ידני:\n{e}")


def main():
    setup_logging(log_level="INFO", log_dir=Path("logs"))
    
    if HAS_CTK:
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("green")
        root = ctk.CTk()
    else:
        root = tk.Tk()
    
    root.title("⚙️ DrawingAI Pro — הגדרות אוטומציה")
    root.geometry("1100x850")
    root.minsize(900, 700)
    root.resizable(True, True)
    
    panel = AutomationPanel(root)
    panel.pack(fill=tk.BOTH, expand=True)
    
    # Store reference for manual GUI button
    panel._open_manual_gui_func = lambda: open_manual_gui(root)
    
    root.mainloop()


if __name__ == "__main__":
    main()
