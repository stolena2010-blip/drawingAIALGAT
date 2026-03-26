import json
import os
import re
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
from typing import Dict, Any, List
from datetime import datetime, timezone, timedelta
import customtkinter as ctk

from automation_runner import AutomationRunner
from src.services.email.graph_helper import GraphAPIHelper
from dashboard_gui import DashboardWindow


# ─── Stdout redirector ────────────────────────────────────────────────
class LogRedirector:
    """Redirects stdout to both console and GUI log widget."""
    def __init__(self, text_widget, original_stdout) -> None:
        self.text_widget = text_widget
        self.original = original_stdout

    def write(self, message) -> None:
        try:
            self.original.write(message)
        except Exception:
            # Windows console can't handle Hebrew/emoji — safe to ignore
            pass
        if message.strip():
            try:
                timestamp = datetime.now().strftime("%H:%M:%S")
                self.text_widget.configure(state=tk.NORMAL)
                self.text_widget.insert(tk.END, f"[{timestamp}] {message}\n")
                line_count = int(self.text_widget.index('end-1c').split('.')[0])
                if line_count > 500:
                    self.text_widget.delete("1.0", "100.0")
                self.text_widget.see(tk.END)
                self.text_widget.configure(state=tk.DISABLED)
                self.text_widget.update_idletasks()
            except Exception:
                pass

    def flush(self) -> None:
        self.original.flush()


# ─── Tooltip helper ───────────────────────────────────────────────────
class ToolTip:
    """Simple tooltip on hover."""
    def __init__(self, widget, text) -> None:
        self.widget = widget
        self.text = text
        self.tooltip = None
        widget.bind("<Enter>", self._show)
        widget.bind("<Leave>", self._hide)

    def _show(self, event=None) -> None:
        x = self.widget.winfo_rootx() + 25
        y = self.widget.winfo_rooty() + 25
        self.tooltip = tk.Toplevel(self.widget)
        self.tooltip.wm_overrideredirect(True)
        self.tooltip.wm_geometry(f"+{x}+{y}")
        label = tk.Label(
            self.tooltip, text=self.text,
            background="#2c3e50", foreground="white",
            font=("Arial", 9), padx=8, pady=4,
            relief=tk.SOLID, borderwidth=1
        )
        label.pack()

    def _hide(self, event=None) -> None:
        if self.tooltip:
            self.tooltip.destroy()
            self.tooltip = None


# ─── Main panel ───────────────────────────────────────────────────────
class AutomationPanel(ttk.Frame):
    def __init__(self, parent: tk.Toplevel) -> None:
        super().__init__(parent)
        self.parent = parent
        self.config_path = Path.cwd() / "automation_config.json"
        self.state_path = Path.cwd() / "automation_state.json"

        # ── Variables ──────────────────────────────────────────────
        self.shared_mailbox_var = tk.StringVar()
        self.mailbox_select_var = tk.StringVar()
        self.folder_name_var = tk.StringVar(value="Inbox")
        self.rerun_folder_var = tk.StringVar(value="")
        self.rerun_mailbox_var = tk.StringVar(value="")
        self.scan_from_date_var = tk.StringVar(value="")
        self.recipient_email_var = tk.StringVar()
        self.download_root_var = tk.StringVar(value=str(Path.cwd() / "email_downloads"))
        self.tosend_folder_var = tk.StringVar()
        self.output_copy_folder_var = tk.StringVar()
        self.interval_var = tk.StringVar(value="10")
        self.max_messages_var = tk.StringVar(value="200")
        self.max_files_per_email_var = tk.StringVar(value="15")
        self.max_file_size_var = tk.StringVar(value="100")
        self.stage1_skip_retry_resolution_var = tk.StringVar(value="8000")

        self._resolution_map = {
            "2048 (מהיר - OCR טוב)": 2048,
            "3072 (מאזן - איכות מעולה)": 3072,
            "4096 (איכות - OCR מושלם)": 4096,
            "12000 (Overkill - ברזולוציה מקסימה)": 12000,
        }
        self._reverse_resolution_map = {v: k for k, v in self._resolution_map.items()}
        self.max_image_dim_var = tk.StringVar(value="3072 (מאזן - איכות מעולה)")

        self.recursive_var = tk.BooleanVar(value=True)
        self.enable_retry_var = tk.BooleanVar(value=True)
        self.auto_start_var = tk.BooleanVar(value=False)
        self.auto_send_var = tk.BooleanVar(value=False)
        self.archive_full_var = tk.BooleanVar(value=False)
        self.cleanup_download_var = tk.BooleanVar(value=True)
        self.mark_as_processed_var = tk.BooleanVar(value=True)
        self.mark_category_name_var = tk.StringVar(value="AI Processed")
        self.mark_category_color_var = tk.StringVar(value="None")
        self.nodraw_category_name_var = tk.StringVar(value="NO DRAW")
        self.nodraw_category_color_var = tk.StringVar(value="None")

        self.stage1_var = tk.BooleanVar(value=True)
        self.stage2_var = tk.BooleanVar(value=True)
        self.stage3_var = tk.BooleanVar(value=True)
        self.stage4_var = tk.BooleanVar(value=True)
        self.stage5_var = tk.BooleanVar(value=True)
        self.stage6_var = tk.BooleanVar(value=True)
        self.stage7_var = tk.BooleanVar(value=True)
        self.stage8_var = tk.BooleanVar(value=True)
        self.stage9_var = tk.BooleanVar(value=True)

        # Per-stage model selection – default = actual model from .env
        self._available_models = self._discover_available_models()
        self._env_stage_defaults = self._get_env_stage_defaults()
        self.stage_model_vars: Dict[int, tk.StringVar] = {}
        for n in range(10):  # stages 0-9
            self.stage_model_vars[n] = tk.StringVar(value=self._env_stage_defaults.get(n, "gpt-4o-vision"))

        self.confidence_level_var = tk.StringVar(value="LOW")
        self.status_var = tk.StringVar(value="מוכן")

        # Advanced settings
        self.debug_mode_var = tk.BooleanVar(value=False)
        self.iai_top_red_var = tk.BooleanVar(value=True)
        self.max_retries_var = tk.StringVar(value="3")
        self.scan_dpi_var = tk.StringVar(value="200")
        self.log_max_size_var = tk.StringVar(value="1")
        self.usd_to_ils_var = tk.StringVar(value="3.7")

        # Timer / cost display
        self._timer_running = False
        self._gollum_animating = False
        self._gollum_x = 0
        self._gollum_dx = 3
        self._gollum_whisper_enabled = tk.BooleanVar(value=False)
        self._gollum_whisper_busy = False

        self._folder_display_to_path: Dict[str, str] = {}

        self._category_color_map = {
            "None": "preset0", "Red": "preset1", "Orange": "preset2",
            "Brown": "preset3", "Yellow": "preset4", "Green": "preset5",
            "Teal": "preset6", "Olive": "preset7", "Blue": "preset8",
            "Purple": "preset9", "Pink": "preset10", "Gray": "preset11",
            "Dark Red": "preset12", "Dark Orange": "preset13", "Dark Brown": "preset14",
            "Dark Yellow": "preset15", "Dark Green": "preset16", "Dark Teal": "preset17",
            "Dark Olive": "preset18", "Dark Blue": "preset19", "Dark Purple": "preset20",
            "Dark Pink": "preset21", "Dark Gray": "preset22", "Black": "preset23",
            "Light Gray": "preset24", "Light Blue": "preset25",
        }

        self.runner = AutomationRunner(self.config_path, self.state_path, self._set_status)

        self._build_ui()
        self._load_config()

        # Close handler — restore stdout
        self.parent.protocol("WM_DELETE_WINDOW", self._on_close)

        if self.auto_start_var.get():
            self.runner.start()
            self._set_status("אוטומציה פעילה")

    @staticmethod
    def _parse_mailboxes_text(raw_text: str) -> list[str]:
        if not raw_text:
            return []
        normalized = raw_text.replace(";", ",").replace("\n", ",")
        seen = set()
        result = []
        for part in normalized.split(","):
            mailbox = part.strip()
            if not mailbox:
                continue
            key = mailbox.lower()
            if key in seen:
                continue
            seen.add(key)
            result.append(mailbox)
        return result

    @staticmethod
    def _format_mailboxes_for_ui(mailboxes: list[str]) -> str:
        return ", ".join([m.strip() for m in mailboxes if str(m).strip()])

    def _refresh_mailbox_selector(self, mailboxes: List[str]) -> None:
        clean = [m.strip() for m in mailboxes if str(m).strip()]
        if hasattr(self, "mailbox_select_combo"):
            self.mailbox_select_combo.configure(values=clean)

        if not clean:
            self.mailbox_select_var.set("")
            return

        current = self.mailbox_select_var.get().strip()
        if current in clean:
            return
        self.mailbox_select_var.set(clean[0])

    def _normalize_folder_name(self, value: str) -> str:
        clean = str(value or "").strip()
        if not clean:
            return ""
        return self._folder_display_to_path.get(clean, clean)

    @staticmethod
    def _format_folder_label(path: str, total_item_count: Any) -> str:
        clean_path = str(path or "").strip()
        if not clean_path:
            return ""
        try:
            if total_item_count is None or str(total_item_count).strip() == "":
                return clean_path
            count = int(total_item_count)
            if count < 0:
                return clean_path
            return f"{clean_path} ({count})"
        except Exception:
            return clean_path

    def _load_folders_for_mailbox(self, mailbox: str, show_message: bool = True) -> None:
        mailbox = (mailbox or "").strip()
        if not mailbox:
            if show_message:
                messagebox.showwarning("אזהרה", "בחר/י תיבה לטעינת תיקיות.")
            return

        try:
            self._set_status(f"טוען תיקיות עבור {mailbox}...")
            helper = GraphAPIHelper(shared_mailbox=mailbox)
            if not helper.test_connection():
                raise RuntimeError("Connection failed")

            all_folders = helper.mailbox.list_folders_recursive() or []

            folder_names: List[str] = []
            raw_paths_seen = set()
            self._folder_display_to_path = {}

            def _add_folder(path: str, total_item_count: Any = None) -> None:
                clean_path = str(path or "").strip()
                if not clean_path:
                    return
                key = clean_path.lower()
                if key in raw_paths_seen:
                    return
                raw_paths_seen.add(key)

                label = self._format_folder_label(clean_path, total_item_count)
                folder_names.append(label)
                self._folder_display_to_path[label] = clean_path

            for folder in all_folders:
                _add_folder(
                    folder.get("path") or folder.get("displayName"),
                    folder.get("totalItemCount")
                )

            if "inbox" not in raw_paths_seen:
                # Inbox may have a Hebrew displayName (e.g. "תיבת דואר נכנס").
                # Use well-known folder name to get its ID and item count.
                wkn_inbox = helper.mailbox.get_well_known_folder("Inbox")
                inbox_count = wkn_inbox.get("totalItemCount") if wkn_inbox else None
                _add_folder("Inbox", inbox_count)

            folder_names = sorted(folder_names, key=lambda x: x.lower())

            if hasattr(self, "folder_combo"):
                self.folder_combo.configure(values=folder_names)

            current_folder_raw = self._normalize_folder_name(self.folder_name_var.get().strip())
            path_to_label = {path: label for label, path in self._folder_display_to_path.items()}
            if current_folder_raw in path_to_label:
                self.folder_name_var.set(path_to_label[current_folder_raw])
            elif "Inbox" in path_to_label:
                self.folder_name_var.set(path_to_label["Inbox"])
            elif folder_names:
                self.folder_name_var.set(folder_names[0])

            self._set_status(f"נטענו {len(folder_names)} תיקיות עבור {mailbox}")
            if show_message:
                messagebox.showinfo("תיקיות נטענו", f"נטענו {len(folder_names)} תיקיות עבור:\n{mailbox}")
        except Exception as e:
            self._set_status(f"שגיאה בטעינת תיקיות: {mailbox}")
            if show_message:
                messagebox.showerror("שגיאה", f"לא ניתן לטעון תיקיות עבור {mailbox}\n\n{e}")

    def _load_folders_for_selected_mailbox(self) -> None:
        self._load_folders_for_mailbox(self.mailbox_select_var.get().strip(), show_message=True)

    # ──────────────────────────────────────────────────────────────
    #  MODEL DISCOVERY
    # ──────────────────────────────────────────────────────────────
    @staticmethod
    def _discover_available_models() -> List[str]:
        """Read distinct model names from .env STAGE_N_MODEL + MODEL_*_ENDPOINT entries."""
        models: set = set()
        env_path = Path.cwd() / ".env"
        if env_path.exists():
            try:
                with open(env_path, "r", encoding="utf-8") as f:
                    for line in f:
                        line = line.strip()
                        if line.startswith("#") or "=" not in line:
                            continue
                        key, _, value = line.partition("=")
                        value = value.strip()
                        if not value:
                            continue
                        if key.startswith("STAGE_") and key.endswith("_MODEL"):
                            models.add(value)
                        elif key == "AZURE_OPENAI_DEPLOYMENT":
                            models.add(value)
                        elif key.startswith("MODEL_") and key.endswith("_ENDPOINT"):
                            pass
            except Exception:
                pass
        # Always include known models
        models.update({"gpt-4o-vision", "gpt-4o-mini-email", "o4-mini", "gpt-5.2", "gpt-5.4"})
        return sorted(models)

    @staticmethod
    def _get_env_stage_defaults() -> Dict[int, str]:
        """Read per-stage default model names from .env (STAGE_N_MODEL keys)."""
        defaults: Dict[int, str] = {}
        env_path = Path.cwd() / ".env"
        fallback = "gpt-4o-vision"
        if env_path.exists():
            try:
                with open(env_path, "r", encoding="utf-8") as f:
                    for line in f:
                        line = line.strip()
                        if line.startswith("#") or "=" not in line:
                            continue
                        key, _, value = line.partition("=")
                        value = value.strip()
                        if not value:
                            continue
                        if key.startswith("STAGE_") and key.endswith("_MODEL"):
                            try:
                                stage_num = int(key.split("_")[1])
                                defaults[stage_num] = value
                            except (ValueError, IndexError):
                                pass
            except Exception:
                pass
        # Fill missing stages with fallback
        for n in range(10):
            if n not in defaults:
                defaults[n] = fallback
        return defaults

    # ──────────────────────────────────────────────────────────────
    #  LOG HELPERS
    # ──────────────────────────────────────────────────────────────
    def _log(self, message: str) -> None:
        """Add a timestamped message to the log terminal."""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)
        self.log_text.configure(state=tk.DISABLED)
        self.update_idletasks()

    def _clear_log(self) -> None:
        """Clear the log terminal."""
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.delete("1.0", tk.END)
        self.log_text.configure(state=tk.DISABLED)

    def _on_close(self) -> None:
        """Stop runner gracefully and close window."""
        try:
            self.runner.stop()
        except Exception:
            pass
        if hasattr(self, "_original_stdout"):
            sys.stdout = self._original_stdout
        self.parent.destroy()

    def _set_status(self, text: str) -> None:
        self.status_var.set(text)

        # Parse cost from stdout (automation_runner prints "Total: $X.XX USD")
        cost_match = re.search(r'\$(\d+\.?\d*)', text)
        files_match = re.search(r'(\d+)\s*(?:files?|drawings?|שרטוט|קבצים)', text)
        if cost_match:
            cost = float(cost_match.group(1))
            files = int(files_match.group(1)) if files_match else 0
            self._update_last_cost(cost, files)

        # Update indicator color based on status keywords
        if hasattr(self, 'status_indicator'):
            if 'פעילה' in text or 'running' in text.lower():
                self._set_status_running()
            elif 'נעצרה' in text or 'stop' in text.lower():
                self._set_status_stopped()
            elif 'מעבד' in text or 'מריץ' in text or 'סבב' in text:
                self._set_status_processing()
            # Stop animation when cycle ends (waiting for next)
            if 'הסבב הבא' in text or 'אין מיילים' in text or 'הושלם' in text:
                self._stop_gollum_animation()

        self._log(text)
        self.update_idletasks()

    # ── Status indicator colors ───────────────────────────────────
    def _set_status_running(self) -> None:
        if hasattr(self, 'status_indicator'):
            self.status_indicator.configure(fg="#00b894")

    def _set_status_stopped(self) -> None:
        if hasattr(self, 'status_indicator'):
            self.status_indicator.configure(fg="#d63031")
        self._stop_gollum_animation()

    def _set_status_processing(self) -> None:
        if hasattr(self, 'status_indicator'):
            self.status_indicator.configure(fg="#fdcb6e")
        self._start_gollum_animation()

    # ── Gollum animation ──────────────────────────────────────────
    def _start_gollum_animation(self) -> None:
        if self._gollum_animating:
            return
        self._gollum_animating = True
        if not self._gollum_bar_visible:
            self._gollum_bar.pack(fill=tk.X, pady=(0, 2), before=self.log_text)
            self._gollum_bar_visible = True
        self._gollum_x = 20
        self._gollum_dx = 3
        self._gollum_step = 0
        self._animate_gollum()

    def _stop_gollum_animation(self) -> None:
        self._gollum_animating = False
        if self._gollum_bar_visible:
            self._gollum_bar.pack_forget()
            self._gollum_bar_visible = False

    def _animate_gollum(self) -> None:
        if not self._gollum_animating:
            return
        try:
            bar_width = self._gollum_bar.winfo_width()
            if bar_width < 50:
                bar_width = 600
            
            # Update track line width
            self._gollum_bar.coords("track", 10, 28, bar_width - 10, 28)
            
            # Move Gollum
            self._gollum_x += self._gollum_dx
            if self._gollum_x >= bar_width - 40:
                self._gollum_dx = -abs(self._gollum_dx)
                # Flip image when changing direction
                if self._gollum_sprite_photo and hasattr(self, '_gollum_sprite_photo_flip'):
                    self._gollum_bar.itemconfigure("gollum", image=self._gollum_sprite_photo_flip)
            elif self._gollum_x <= 40:
                self._gollum_dx = abs(self._gollum_dx)
                if self._gollum_sprite_photo:
                    self._gollum_bar.itemconfigure("gollum", image=self._gollum_sprite_photo)
            
            self._gollum_bar.coords("gollum", self._gollum_x, 28)
            
            # Move the Precious ring ahead of Gollum (he chases it!)
            ring_offset = 60 if self._gollum_dx > 0 else -60
            ring_cx = self._gollum_x + ring_offset
            self._gollum_bar.coords("ring_glow", ring_cx - 15, 10, ring_cx + 15, 46)
            self._gollum_bar.coords("ring", ring_cx - 11, 14, ring_cx + 11, 42)
            # Pulsating glow effect on ring
            self._ring_glow_phase = (self._ring_glow_phase + 1) % 20
            glow_colors = ["#FFD700", "#FFEC8B", "#FFF8DC", "#FFEC8B", "#FFD700",
                           "#DAA520", "#B8860B", "#DAA520", "#FFD700", "#FFEC8B",
                           "#FFF8DC", "#FFEC8B", "#FFD700", "#DAA520", "#B8860B",
                           "#DAA520", "#FFD700", "#FFEC8B", "#FFF8DC", "#FFEC8B"]
            self._gollum_bar.itemconfigure("ring_glow", outline=glow_colors[self._ring_glow_phase])
            
            # Trail sparkle behind Gollum
            self._gollum_step += 1
            sparkles = ["✨", "·", "⭐", "·", "💎", "·"][self._gollum_step % 6]
            trail_x = self._gollum_x - (30 if self._gollum_dx > 0 else -30)
            self._gollum_bar.coords("trail", trail_x, 28)
            self._gollum_bar.itemconfigure("trail", text=sparkles)
            
            # Center status text
            phrases = [
                "Green Coat — DrawingAI Pro",
                "analyzing drawings...",
                "מנתח שרטוטים...",
                "processing files...",
                "מעבד נתונים...",
                "scanning pages...",
                "Algat — אלגט",
                "extracting data...",
                "חילוץ אוטומטי משרטוטים",
            ]
            if self._gollum_step % 40 == 0:
                phrase = phrases[(self._gollum_step // 40) % len(phrases)]
                self._gollum_bar.itemconfigure("bartext", text=phrase)
                # Whisper the phrase if enabled
                if self._gollum_whisper_enabled.get():
                    self._whisper_gollum(phrase)
            self._gollum_bar.coords("bartext", bar_width // 2, 28)
            
        except Exception:
            pass
        
        if self._gollum_animating:
            self.after(30, self._animate_gollum)

    # ── Gollum whisper (TTS) ──────────────────────────────────────
    def _whisper_gollum(self, text: str) -> None:
        """Speak a Gollum phrase in a background thread (creepy whispery voice)."""
        if self._gollum_whisper_busy:
            return
        import threading

        def _speak():
            try:
                self._gollum_whisper_busy = True
                import pyttsx3
                engine = pyttsx3.init()
                engine.setProperty('rate', 95)      # very slow, whispery
                engine.setProperty('volume', 0.45)   # quiet whisper
                # Try to pick the creepiest available voice
                voices = engine.getProperty('voices')
                # Prefer David (deep male) on Windows
                chosen = None
                for v in voices:
                    name_lower = v.name.lower()
                    if 'david' in name_lower:
                        chosen = v.id
                        break
                    elif 'male' in name_lower and not chosen:
                        chosen = v.id
                if chosen:
                    engine.setProperty('voice', chosen)
                # Strip emoji for TTS
                clean = text.replace('💍', '').replace('💎', '').replace('✨', '').strip()
                if clean:
                    # Add dramatic pauses for Gollum effect
                    gollum_text = clean.replace('...', '... ... ')
                    engine.say(gollum_text)
                    engine.runAndWait()
                    engine.stop()
            except Exception:
                pass
            finally:
                self._gollum_whisper_busy = False

        threading.Thread(target=_speak, daemon=True).start()

    # ── Countdown timer ───────────────────────────────────────────
    def _start_timer(self) -> None:
        self._timer_running = True
        self._update_timer()

    def _stop_timer(self) -> None:
        self._timer_running = False
        if hasattr(self, 'timer_var'):
            self.timer_var.set("")

    def _update_timer(self) -> None:
        if not self._timer_running:
            return
        try:
            if self.state_path.exists():
                with open(self.state_path, "r") as f:
                    state = json.load(f)
                last_checked = state.get("last_checked", "")
                if last_checked:
                    last_time = datetime.fromisoformat(
                        last_checked.replace("Z", "+00:00")
                    )
                    config = {}
                    if self.config_path.exists():
                        with open(self.config_path, "r") as f:
                            config = json.load(f)
                    interval = int(config.get("poll_interval_minutes", 10))
                    next_run = last_time + timedelta(minutes=interval)
                    now = datetime.now(timezone.utc)
                    remaining = next_run - now
                    if remaining.total_seconds() > 0:
                        mins = int(remaining.total_seconds() // 60)
                        secs = int(remaining.total_seconds() % 60)
                        self.timer_var.set(f"⏱ סבב הבא: {mins}:{secs:02d}")
                    else:
                        self.timer_var.set("⏱ סבב רץ עכשיו...")
                        self._set_status_processing()
        except Exception:
            pass
        if self._timer_running:
            self.after(1000, self._update_timer)

    # ── Last cost display ─────────────────────────────────────────
    def _update_last_cost(self, cost_usd: float, files: int) -> None:
        if not hasattr(self, 'last_cost_var'):
            return
        try:
            usd_to_ils = float(self.usd_to_ils_var.get() or "3.7")
        except ValueError:
            usd_to_ils = 3.7
        cost_ils = cost_usd * usd_to_ils
        self.last_cost_var.set(f"💰 ${cost_usd:.2f} (₪{cost_ils:.2f}) — {files} קבצים")

    # ── Skip-senders list management ─────────────────────────────
    def _add_skip_sender(self) -> None:
        """Add email address(es) from the entry to the skip list."""
        raw = self._skip_sender_entry.get().strip()
        if not raw:
            return
        # Support pasting multiple addresses separated by ; , or space
        addresses = [a.strip().lower() for a in re.split(r'[;,\s]+', raw) if a.strip()]
        existing = set(self.skip_senders_listbox.get(0, tk.END))
        added = 0
        for addr in addresses:
            if addr and addr not in existing:
                self.skip_senders_listbox.insert(tk.END, addr)
                existing.add(addr)
                added += 1
        self._skip_sender_entry.delete(0, tk.END)
        if added == 0 and addresses:
            messagebox.showinfo("קיים", f"הכתובת כבר ברשימה")

    def _remove_skip_sender(self) -> None:
        """Remove selected email from skip list."""
        sel = self.skip_senders_listbox.curselection()
        if not sel:
            messagebox.showinfo("בחר", "סמן כתובת ברשימה כדי להסיר")
            return
        for idx in reversed(sel):
            self.skip_senders_listbox.delete(idx)

    def _add_skip_category(self) -> None:
        """Add category name to skip-categories list."""
        raw = self._skip_category_entry.get().strip()
        if not raw:
            return
        categories = [c.strip() for c in re.split(r'[;,]+', raw) if c.strip()]
        existing = set(self.skip_categories_listbox.get(0, tk.END))
        added = 0
        for cat in categories:
            if cat and cat not in existing:
                self.skip_categories_listbox.insert(tk.END, cat)
                existing.add(cat)
                added += 1
        self._skip_category_entry.delete(0, tk.END)
        if added == 0 and categories:
            messagebox.showinfo("קיים", f"הקטגוריה כבר ברשימה")

    def _remove_skip_category(self) -> None:
        """Remove selected category from skip list."""
        sel = self.skip_categories_listbox.curselection()
        if not sel:
            messagebox.showinfo("בחר", "סמן קטגוריה ברשימה כדי להסיר")
            return
        for idx in reversed(sel):
            self.skip_categories_listbox.delete(idx)

    # ── Open folder ───────────────────────────────────────────────
    def _open_folder(self, path_var) -> None:
        folder = path_var.get().strip()
        if folder and os.path.isdir(folder):
            subprocess.Popen(f'explorer "{folder}"')
        else:
            messagebox.showwarning("שגיאה", "התיקייה לא קיימת")

    def _build_ui(self) -> None:
        self.configure(padding=8)

        # Increase default font size
        import tkinter.font as tkFont
        default_font = tkFont.nametofont("TkDefaultFont")
        default_font.configure(size=10)

        # ── ACTION BUTTONS (pinned at bottom, always visible) ─────
        actions = ctk.CTkFrame(self)
        actions.pack(side=tk.BOTTOM, fill=tk.X, pady=(4, 0))

        for text, cmd, fg, hover, w in [
            ("💾 שמור", self._save_config, "#00d4aa", "#00b894", 100),
            ("🔌 בדוק חיבור", self._test_mailboxes, "#0984e3", "#0773c5", 120),
            ("▶️ הרץ סבב", self._run_once, "#6c5ce7", "#5a4bd1", 120),
            ("🏋️ הרץ כבדים", self._run_heavy, "#e17055", "#d35400", 120),
            ("🚀 הפעל אוטומציה", self._start, "#00b894", "#00a381", 140),
            ("⏹ עצור", self._stop, "#d63031", "#b71c1c", 80),
            ("🔄 Reset", self._reset_state, "#636e72", "#4a5568", 80),
            ("📊 Dashboard", self._open_dashboard, "#FF9800", "#F57C00", 120),
            ("🔧 ידני", self._open_manual_gui, "#78909C", "#546E7A", 80),
        ]:
            ctk.CTkButton(
                actions, text=text, command=cmd,
                fg_color=fg, hover_color=hover, text_color="#1a1a2e",
                width=w, height=35, corner_radius=8,
            ).pack(side=tk.RIGHT, padx=4)

        ctk.CTkLabel(
            actions, textvariable=self.status_var,
            text_color="#00d4aa", font=("Arial", 11),
        ).pack(side=tk.RIGHT, padx=8)

        # ── PanedWindow: settings (top) + log (bottom) ────────────
        paned = ttk.PanedWindow(self, orient=tk.VERTICAL)
        paned.pack(fill=tk.BOTH, expand=True)

        # Scrollable settings area
        settings_outer = ttk.Frame(paned)
        paned.add(settings_outer, weight=3)

        settings_canvas = tk.Canvas(settings_outer, highlightthickness=0)
        settings_scrollbar = ttk.Scrollbar(settings_outer, orient=tk.VERTICAL, command=settings_canvas.yview)
        settings_frame = ttk.Frame(settings_canvas)

        settings_frame.bind(
            "<Configure>",
            lambda e: settings_canvas.configure(scrollregion=settings_canvas.bbox("all"))
        )
        settings_canvas.create_window((0, 0), window=settings_frame, anchor="nw")
        settings_canvas.bind("<Configure>", lambda e: settings_canvas.itemconfigure("all", width=e.width))
        settings_canvas.configure(yscrollcommand=settings_scrollbar.set)

        settings_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        settings_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Bind mousewheel to scroll settings area
        def _on_mousewheel(event) -> None:
            settings_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        settings_canvas.bind_all("<MouseWheel>", _on_mousewheel)

        log_frame = ttk.LabelFrame(paned, text="📋 לוג ריצה", padding=8)
        paned.add(log_frame, weight=2)

        # ═══════════════════════════════════════════════════════════
        # HEADER
        # ═══════════════════════════════════════════════════════════
        # Header with company logo
        header_frame = ttk.Frame(settings_frame)
        header_frame.pack(fill=tk.X, pady=(0, 8))
        
        # Load company logo
        try:
            from PIL import Image, ImageTk
            import os
            icon_path = os.path.join(os.path.dirname(__file__), "company_logo.png")
            if os.path.exists(icon_path):
                logo_img = Image.open(icon_path).resize((64, 64), Image.LANCZOS)
                self._logo_photo = ImageTk.PhotoImage(logo_img)
                logo_label = tk.Label(header_frame, image=self._logo_photo, bg="#2b2b2b")
                logo_label.pack(side=tk.RIGHT, padx=(0, 8))
        except Exception:
            pass  # No icon, no problem
        
        header = ctk.CTkLabel(
            header_frame,
            text="Green Coat — DrawingAI Pro — אוטומציה",
            font=("Arial", 22, "bold"),
            text_color="#00d4aa",
        )
        header.pack(side=tk.RIGHT)

        # ═══════════════════════════════════════════════════════════
        # STATUS BAR
        # ═══════════════════════════════════════════════════════════
        status_bar = ttk.Frame(settings_frame)
        status_bar.pack(fill=tk.X, pady=(0, 8))

        self.status_indicator = tk.Label(
            status_bar, text="●", font=("Arial", 16), fg="#d63031"
        )
        self.status_indicator.pack(side=tk.RIGHT, padx=(0, 8))

        self.status_label = ttk.Label(
            status_bar, textvariable=self.status_var, font=("Arial", 11)
        )
        self.status_label.pack(side=tk.RIGHT)

        self.counter_var = tk.StringVar(value="")
        ttk.Label(
            status_bar, textvariable=self.counter_var,
            font=("Arial", 10), foreground="#0984e3"
        ).pack(side=tk.RIGHT, padx=8)

        self.timer_var = tk.StringVar(value="")
        ttk.Label(
            status_bar, textvariable=self.timer_var,
            font=("Arial", 10), foreground="#6c5ce7"
        ).pack(side=tk.RIGHT, padx=8)

        self.last_cost_var = tk.StringVar(value="")
        ttk.Label(
            status_bar, textvariable=self.last_cost_var,
            font=("Arial", 10), foreground="#e17055"
        ).pack(side=tk.RIGHT, padx=8)

        # ═══════════════════════════════════════════════════════════
        # TOP ROW: Email+Folders (right)  +  Run+Stages (left)
        # ═══════════════════════════════════════════════════════════
        top_row = ttk.Frame(settings_frame)
        top_row.pack(fill=tk.X, pady=(0, 8))

        # ─── Email + Folders (RIGHT box) ─────────────────────────
        email_frame = ttk.LabelFrame(top_row, text="📧 הגדרות מייל ותיקיות", padding=10)
        email_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 4))
        email_frame.columnconfigure(1, weight=1)

        r = 0
        ttk.Label(email_frame, text="תיבות משותפות:").grid(row=r, column=0, sticky=tk.W, padx=4, pady=3)
        ttk.Entry(email_frame, textvariable=self.shared_mailbox_var, width=35).grid(row=r, column=1, sticky=tk.EW, padx=4, pady=3)

        r += 1
        ttk.Label(email_frame, text="תת-תיקייה:").grid(row=r, column=0, sticky=tk.W, padx=4, pady=3)
        self.folder_combo = ttk.Combobox(email_frame, textvariable=self.folder_name_var, width=33)
        self.folder_combo.grid(row=r, column=1, sticky=tk.EW, padx=4, pady=3)

        r += 1
        ttk.Label(email_frame, text="תיקיית RERUN:").grid(row=r, column=0, sticky=tk.W, padx=4, pady=3)
        rerun_row_frame = ttk.Frame(email_frame)
        rerun_row_frame.grid(row=r, column=1, sticky=tk.EW, padx=4, pady=3)
        self.rerun_folder_entry = ttk.Entry(rerun_row_frame, textvariable=self.rerun_folder_var, width=14)
        self.rerun_folder_entry.pack(side=tk.LEFT)
        ttk.Label(rerun_row_frame, text=" תיבה:").pack(side=tk.LEFT, padx=(6, 2))
        self.rerun_mailbox_entry = ttk.Entry(rerun_row_frame, textvariable=self.rerun_mailbox_var, width=20)
        self.rerun_mailbox_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)

        r += 1
        ttk.Label(email_frame, text="סרוק מתאריך:").grid(row=r, column=0, sticky=tk.W, padx=4, pady=3)
        scan_from_frame = ttk.Frame(email_frame)
        scan_from_frame.grid(row=r, column=1, sticky=tk.EW, padx=4, pady=3)
        self.scan_from_entry = ttk.Entry(scan_from_frame, textvariable=self.scan_from_date_var, width=20)
        self.scan_from_entry.pack(side=tk.RIGHT, padx=(0, 6))
        ttk.Label(scan_from_frame, text="DD/MM/YYYY HH:MM", foreground="gray").pack(side=tk.RIGHT)

        r += 1
        ttk.Label(email_frame, text="תיבה להצגה:").grid(row=r, column=0, sticky=tk.W, padx=4, pady=3)
        mailbox_tools = ttk.Frame(email_frame)
        mailbox_tools.grid(row=r, column=1, sticky=tk.EW, padx=4, pady=3)
        self.mailbox_select_combo = ttk.Combobox(mailbox_tools, textvariable=self.mailbox_select_var, width=22, state="readonly")
        self.mailbox_select_combo.pack(side=tk.RIGHT, padx=(0, 6))
        ttk.Button(mailbox_tools, text="טען תיקיות", command=self._load_folders_for_selected_mailbox).pack(side=tk.RIGHT)

        r += 1
        ttk.Label(email_frame, text="נמען לשליחה:").grid(row=r, column=0, sticky=tk.W, padx=4, pady=3)
        ttk.Entry(email_frame, textvariable=self.recipient_email_var, width=35).grid(row=r, column=1, sticky=tk.EW, padx=4, pady=3)

        r += 1
        ttk.Label(email_frame, text="קטגוריה לסימון:").grid(row=r, column=0, sticky=tk.W, padx=4, pady=3)
        cat_fr = ttk.Frame(email_frame)
        cat_fr.grid(row=r, column=1, sticky=tk.EW, padx=4, pady=3)
        ttk.Entry(cat_fr, textvariable=self.mark_category_name_var, width=18).pack(side=tk.RIGHT, padx=(0, 6))
        color_values = list(self._category_color_map.keys())
        ttk.Combobox(cat_fr, textvariable=self.mark_category_color_var, values=color_values, width=12, state="readonly").pack(side=tk.RIGHT)

        r += 1
        ttk.Label(email_frame, text="קטגוריה NO DRAW:").grid(row=r, column=0, sticky=tk.W, padx=4, pady=3)
        nodraw_cat_fr = ttk.Frame(email_frame)
        nodraw_cat_fr.grid(row=r, column=1, sticky=tk.EW, padx=4, pady=3)
        ttk.Entry(nodraw_cat_fr, textvariable=self.nodraw_category_name_var, width=18).pack(side=tk.RIGHT, padx=(0, 6))
        ttk.Label(nodraw_cat_fr, text="צבע:").pack(side=tk.RIGHT, padx=(0, 2))
        ttk.Combobox(nodraw_cat_fr, textvariable=self.nodraw_category_color_var, values=color_values, width=12, state="readonly").pack(side=tk.RIGHT)
        ToolTip(nodraw_cat_fr, "קטגוריה למיילים ללא שרטוטים מזוהים.\nלא ישלח מייל B2B.")

        r += 1
        ttk.Label(email_frame, text="דלג על שולחים:").grid(row=r, column=0, sticky=tk.NW, padx=4, pady=3)
        skip_frame = ttk.Frame(email_frame)
        skip_frame.grid(row=r, column=1, sticky=tk.EW, padx=4, pady=3)

        # Listbox showing skip senders
        self.skip_senders_listbox = tk.Listbox(skip_frame, height=4, width=38, font=("Consolas", 9))
        self.skip_senders_listbox.pack(side=tk.TOP, fill=tk.X, expand=True)

        # Entry + Add / Remove buttons
        skip_btn_frame = ttk.Frame(skip_frame)
        skip_btn_frame.pack(side=tk.TOP, fill=tk.X, pady=(3, 0))
        self._skip_sender_entry = ttk.Entry(skip_btn_frame, width=28)
        self._skip_sender_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self._skip_sender_entry.bind("<Return>", lambda e: self._add_skip_sender())
        ttk.Button(skip_btn_frame, text="הוסף", width=5, command=self._add_skip_sender).pack(side=tk.LEFT, padx=(4, 0))
        ttk.Button(skip_btn_frame, text="הסר", width=5, command=self._remove_skip_sender).pack(side=tk.LEFT, padx=(2, 0))
        ToolTip(skip_frame, "הקלד כתובת מייל ולחץ 'הוסף'.\nלהסרה – סמן שורה ולחץ 'הסר'.")

        r += 1
        ttk.Label(email_frame, text="דלג על קטגוריות:").grid(row=r, column=0, sticky=tk.NW, padx=4, pady=3)
        skip_cat_frame = ttk.Frame(email_frame)
        skip_cat_frame.grid(row=r, column=1, sticky=tk.EW, padx=4, pady=3)

        self.skip_categories_listbox = tk.Listbox(skip_cat_frame, height=3, width=38, font=("Consolas", 9))
        self.skip_categories_listbox.pack(side=tk.TOP, fill=tk.X, expand=True)

        skip_cat_btn_frame = ttk.Frame(skip_cat_frame)
        skip_cat_btn_frame.pack(side=tk.TOP, fill=tk.X, pady=(3, 0))
        self._skip_category_entry = ttk.Entry(skip_cat_btn_frame, width=28)
        self._skip_category_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self._skip_category_entry.bind("<Return>", lambda e: self._add_skip_category())
        ttk.Button(skip_cat_btn_frame, text="הוסף", width=5, command=self._add_skip_category).pack(side=tk.LEFT, padx=(4, 0))
        ttk.Button(skip_cat_btn_frame, text="הסר", width=5, command=self._remove_skip_category).pack(side=tk.LEFT, padx=(2, 0))
        ToolTip(skip_cat_frame, "הקלד שם קטגוריה ולחץ 'הוסף'.\nאם מייל מסומן באחת מהקטגוריות — ידולג.\nלהסרה – סמן שורה ולחץ 'הסר'.")

        # Folders inside same box (separator)
        ttk.Separator(email_frame, orient=tk.HORIZONTAL).grid(row=r+1, column=0, columnspan=2, sticky=tk.EW, pady=6)

        r += 2
        for i, (label_text, var, browse_cmd) in enumerate([
            ("תיקיית הורדה:", self.download_root_var, self._browse_download),
            ("TO_SEND:", self.tosend_folder_var, self._browse_tosend),
            ("תיקיית שמירה:", self.output_copy_folder_var, self._browse_output_copy),
        ]):
            ttk.Label(email_frame, text=label_text).grid(row=r+i, column=0, sticky=tk.W, padx=4, pady=3)
            folder_row = ttk.Frame(email_frame)
            folder_row.grid(row=r+i, column=1, sticky=tk.EW, padx=4, pady=3)
            ttk.Entry(folder_row, textvariable=var, width=50).pack(side=tk.LEFT, fill=tk.X, expand=True)
            ttk.Button(folder_row, text="עיון...", command=browse_cmd).pack(side=tk.LEFT, padx=(6, 0))
            ttk.Button(folder_row, text="📂", width=3,
                       command=lambda v=var: self._open_folder(v)).pack(side=tk.LEFT, padx=(2, 0))

        # ─── Run settings + Stages (LEFT box) ────────────────────
        right_col = ttk.Frame(top_row)
        right_col.pack(side=tk.LEFT, fill=tk.BOTH, padx=(4, 0))

        # Stages – two rows, each stage = checkbox + model combo vertically
        stages_frame = ttk.LabelFrame(right_col, text="🔬 שלבי חילוץ ומודלים", padding=6)
        stages_frame.pack(fill=tk.X, pady=(0, 4))

        stage_defs = [
            ("0: זיהוי", None, 0),
            ("1: בסיסי", self.stage1_var, 1),
            ("2: תהליכים", self.stage2_var, 2),
            ("3: NOTES", self.stage3_var, 3),
            ("4: שטח", self.stage4_var, 4),
            ("5: Fallback", self.stage5_var, 5),
            ("6: PL", self.stage6_var, 6),
            ("7: email", self.stage7_var, 7),
            ("8: הזמנות", self.stage8_var, 8),
            ("9: מיזוג", self.stage9_var, 9),
        ]

        # Split into two rows: stages 0-4 (top), stages 5-9 (bottom)
        row1_defs = stage_defs[:5]   # 0,1,2,3,4
        row2_defs = stage_defs[5:]   # 5,6,7,8,9

        for row_defs in (row1_defs, row2_defs):
            row_frame = ttk.Frame(stages_frame)
            row_frame.pack(anchor=tk.E, pady=(2, 2))
            for text, var, stage_n in row_defs:
                cell = ttk.Frame(row_frame)
                cell.pack(side=tk.RIGHT, padx=(2, 6))
                # Top: checkbox or label
                if var is not None:
                    ttk.Checkbutton(cell, text=text, variable=var).pack(anchor=tk.E)
                else:
                    ttk.Label(cell, text=text, foreground="#888").pack(anchor=tk.E)
                # Bottom: model combo
                combo = ttk.Combobox(
                    cell,
                    textvariable=self.stage_model_vars[stage_n],
                    values=self._available_models,
                    width=16,
                    state="readonly",
                )
                combo.pack(anchor=tk.E, pady=(1, 0))

        # Run settings (next to stages)
        run_frame = ttk.LabelFrame(right_col, text="⏱ הגדרות ריצה", padding=10)
        run_frame.pack(fill=tk.X, pady=(4, 0))

        run_top = ttk.Frame(run_frame)
        run_top.pack(anchor=tk.W)
        ttk.Label(run_top, text="דקות בין סבבים:").pack(side=tk.RIGHT, padx=(0, 4))
        self.interval_entry = ttk.Entry(run_top, textvariable=self.interval_var, width=6)
        self.interval_entry.pack(side=tk.RIGHT, padx=(0, 20))
        ttk.Label(run_top, text="כמות מיילים לסבב:").pack(side=tk.RIGHT, padx=(0, 4))
        self.max_messages_entry = ttk.Entry(run_top, textvariable=self.max_messages_var, width=6)
        self.max_messages_entry.pack(side=tk.RIGHT)

        run_top2 = ttk.Frame(run_frame)
        run_top2.pack(anchor=tk.W, pady=(4, 0))
        ttk.Label(run_top2, text="מקסימום קבצים למייל (0=ללא הגבלה):").pack(side=tk.RIGHT, padx=(0, 4))
        self.max_files_per_email_entry = ttk.Entry(run_top2, textvariable=self.max_files_per_email_var, width=6)
        self.max_files_per_email_entry.pack(side=tk.RIGHT)

        run_checks = ttk.Frame(run_frame)
        run_checks.pack(anchor=tk.W, pady=(6, 0))
        ttk.Checkbutton(run_checks, text="הפעל בעת פתיחה", variable=self.auto_start_var).pack(side=tk.RIGHT, padx=(0, 12))
        ttk.Checkbutton(run_checks, text="שלח מייל אוטומטית", variable=self.auto_send_var).pack(side=tk.RIGHT, padx=(0, 12))
        ttk.Checkbutton(run_checks, text="שמור עותק מלא", variable=self.archive_full_var).pack(side=tk.RIGHT, padx=(0, 12))
        ttk.Checkbutton(run_checks, text="מחק אחרי העברה", variable=self.cleanup_download_var).pack(side=tk.RIGHT, padx=(0, 12))
        ttk.Checkbutton(run_checks, text="סמן מעובד", variable=self.mark_as_processed_var).pack(side=tk.RIGHT)

        # ═══════════════════════════════════════════════════════════
        # ADVANCED SETTINGS (under run settings in right column)
        # ═══════════════════════════════════════════════════════════
        advanced_frame = ttk.LabelFrame(right_col, text="⚙️ הגדרות מתקדמות", padding=10)
        advanced_frame.pack(fill=tk.X, pady=(4, 0))

        adv_row = ttk.Frame(advanced_frame)
        adv_row.pack(fill=tk.X)

        # Processing settings (left side)
        proc_col = ttk.Frame(adv_row)
        proc_col.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        pr = ttk.Frame(proc_col)
        pr.pack(anchor=tk.W, pady=2)
        ttk.Label(pr, text="מקסימום רזולוציה:").pack(side=tk.RIGHT, padx=(0, 4))
        ttk.Combobox(
            pr,
            textvariable=self.max_image_dim_var,
            values=list(self._resolution_map.keys()),
            width=30,
            state="readonly",
        ).pack(side=tk.RIGHT)

        pr2 = ttk.Frame(proc_col)
        pr2.pack(anchor=tk.W, pady=2)
        ttk.Label(pr2, text="מקסימום גודל קובץ:").pack(side=tk.RIGHT, padx=(0, 4))
        ttk.Entry(pr2, textvariable=self.max_file_size_var, width=6).pack(side=tk.RIGHT, padx=(0, 4))
        ttk.Label(pr2, text="MB").pack(side=tk.RIGHT, padx=(0, 12))
        ttk.Label(pr2, text="דילוג Retry שלב 1:").pack(side=tk.RIGHT, padx=(0, 4))
        ttk.Entry(pr2, textvariable=self.stage1_skip_retry_resolution_var, width=6).pack(side=tk.RIGHT, padx=(0, 4))
        ttk.Label(pr2, text="px").pack(side=tk.RIGHT)

        pr3 = ttk.Frame(proc_col)
        pr3.pack(anchor=tk.W, pady=2)
        ttk.Label(pr3, text="רמת ביטחון B2B:").pack(side=tk.RIGHT, padx=(0, 4))
        ttk.Radiobutton(pr3, text="LOW", variable=self.confidence_level_var, value="LOW").pack(side=tk.RIGHT, padx=(0, 10))
        ttk.Radiobutton(pr3, text="MEDIUM", variable=self.confidence_level_var, value="MEDIUM").pack(side=tk.RIGHT, padx=(0, 10))
        ttk.Radiobutton(pr3, text="HIGH", variable=self.confidence_level_var, value="HIGH").pack(side=tk.RIGHT, padx=(0, 12))
        ttk.Checkbutton(pr3, text="כולל תת-תיקיות", variable=self.recursive_var).pack(side=tk.RIGHT, padx=(0, 12))
        ttk.Checkbutton(pr3, text="ניסיונות נוספים", variable=self.enable_retry_var).pack(side=tk.RIGHT)

        # Other advanced (right side)
        other_col = ttk.Frame(adv_row)
        other_col.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.debug_check = ttk.Checkbutton(other_col, text="Debug mode (הדפסות מפורטות)", variable=self.debug_mode_var)
        self.debug_check.pack(anchor=tk.W, pady=2)
        self.iai_check = ttk.Checkbutton(other_col, text="IAI top-red fallback", variable=self.iai_top_red_var)
        self.iai_check.pack(anchor=tk.W, pady=2)
        self.whisper_check = ttk.Checkbutton(other_col, text="🗣 TTS notifications", variable=self._gollum_whisper_enabled)
        self.whisper_check.pack(anchor=tk.W, pady=2)

        misc_row = ttk.Frame(other_col)
        misc_row.pack(anchor=tk.W, pady=2)
        ttk.Label(misc_row, text="ניסיונות חוזרים:").pack(side=tk.RIGHT)
        self.retries_entry = ttk.Entry(misc_row, textvariable=self.max_retries_var, width=4)
        self.retries_entry.pack(side=tk.RIGHT, padx=4)
        ttk.Label(misc_row, text="DPI:").pack(side=tk.RIGHT, padx=(0, 4))
        self.dpi_combo = ttk.Combobox(misc_row, textvariable=self.scan_dpi_var,
                     values=["150", "200", "300"], width=5, state="readonly")
        self.dpi_combo.pack(side=tk.RIGHT, padx=(0, 8))
        ttk.Label(misc_row, text="לוג מקס:").pack(side=tk.RIGHT, padx=(0, 4))
        self.log_size_entry = ttk.Entry(misc_row, textvariable=self.log_max_size_var, width=4)
        self.log_size_entry.pack(side=tk.RIGHT, padx=(0, 2))
        ttk.Label(misc_row, text="MB").pack(side=tk.RIGHT, padx=(0, 8))
        ttk.Label(misc_row, text="$/₪:").pack(side=tk.RIGHT, padx=(0, 4))
        self.ils_entry = ttk.Entry(misc_row, textvariable=self.usd_to_ils_var, width=5)
        self.ils_entry.pack(side=tk.RIGHT)

        # ═══════════════════════════════════════════════════════════
        # LIVE LOG TERMINAL (bottom pane)
        # ═══════════════════════════════════════════════════════════
        
        # ─── Animated Gollum progress bar ─────────────────────────
        self._gollum_bar = tk.Canvas(
            log_frame, height=56, bg="#1a1a2e", highlightthickness=0
        )
        self._gollum_bar.pack(fill=tk.X, pady=(0, 2))
        # Draw track line
        self._gollum_bar.create_line(
            10, 28, 800, 28, fill="#2d2d4a", width=2, tags="track"
        )
        # Load company logo sprite for animation
        self._gollum_sprite_photo = None
        try:
            from PIL import Image, ImageTk
            import os
            _sprite_path = os.path.join(os.path.dirname(__file__), "company_logo.png")
            if os.path.exists(_sprite_path):
                _sprite_img = Image.open(_sprite_path).resize((48, 48), Image.LANCZOS)
                self._gollum_sprite_photo = ImageTk.PhotoImage(_sprite_img)
                # Also create a flipped version for when moving left
                _sprite_flip = _sprite_img.transpose(Image.FLIP_LEFT_RIGHT)
                self._gollum_sprite_photo_flip = ImageTk.PhotoImage(_sprite_flip)
        except Exception:
            pass
        # Gollum character on the bar (image or fallback emoji)
        if self._gollum_sprite_photo:
            self._gollum_sprite = self._gollum_bar.create_image(
                20, 28, image=self._gollum_sprite_photo, tags="gollum"
            )
        else:
            self._gollum_sprite = self._gollum_bar.create_text(
                20, 28, text="💍", font=("Segoe UI Emoji", 18), tags="gollum"
            )
        # The Precious ring ahead of Gollum (golden glow + ring)
        self._gollum_bar.create_oval(
            55, 10, 85, 46, outline="#FFD700", width=2, fill="", tags="ring_glow"
        )
        self._gollum_bar.create_oval(
            59, 14, 81, 42, outline="#FFA500", width=3, fill="", tags="ring"
        )
        self._ring_glow_phase = 0
        # Trail sparkles (decorative)
        self._gollum_trail = self._gollum_bar.create_text(
            20, 28, text="", font=("Segoe UI Emoji", 10), fill="#FFD700", tags="trail"
        )
        # Status text on bar
        self._gollum_bar_text = self._gollum_bar.create_text(
            400, 28, text="", font=("Consolas", 12, "bold"), fill="#7777aa", tags="bartext"
        )
        # Hide bar initially
        self._gollum_bar.pack_forget()
        self._gollum_bar_visible = False
        
        self.log_text = tk.Text(
            log_frame,
            height=12,
            bg="#1a1a2e",
            fg="#00ff00",
            font=("Consolas", 9),
            wrap=tk.WORD,
            state=tk.DISABLED,
            borderwidth=0,
            padx=8,
            pady=8,
        )
        log_scrollbar = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_scrollbar.set)
        log_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.pack(fill=tk.BOTH, expand=True)

        ttk.Button(log_frame, text="🗑 נקה לוג", command=self._clear_log).pack(anchor=tk.E, pady=(4, 0))

        # ── Redirect stdout to log terminal ───────────────────────
        self._original_stdout = sys.stdout
        sys.stdout = LogRedirector(self.log_text, self._original_stdout)

        # ── Tooltips ──────────────────────────────────────────────
        ToolTip(self.folder_combo, "תת-תיקייה בתיבה המשותפת. השאר Inbox לברירת מחדל.")
        ToolTip(self.rerun_folder_entry, "שם תיקייה ב-Outlook להרצה חוזרת (RERUN). גררי מייל לכאן כדי לשלוח שוב עם כל הפריטים. השאירי ריק אם לא צריך.")
        ToolTip(self.rerun_mailbox_entry, "תיבה בה נמצאת תיקיית RERUN. השאירי ריק אם היא באותה תיבה ראשית.")
        ToolTip(self.scan_from_entry, "תאריך התחלה לסריקה. מיילים שהתקבלו לפני תאריך זה יתעלמו.\nפורמט: DD/MM/YYYY HH:MM\nדוגמה: 26/02/2026 15:00\nהשאירי ריק לסריקה רגילה (לפי last_checked).")
        ToolTip(self.mailbox_select_combo, "בחר תיבה ולחץ 'טען תיקיות' לראות את כל התיקיות")
        ToolTip(self.interval_entry, "כל כמה דקות לבדוק מיילים חדשים")
        ToolTip(self.max_messages_entry, "כמה מיילים לעבד בסבב אחד (מקסימום)")
        ToolTip(self.max_files_per_email_entry, "מיילים עם יותר קבצים מהסף יסומנו AI HEAVY וידלגו. 0=ללא הגבלה")
        ToolTip(self.debug_check, "מפעיל הדפסות מפורטות לפתרון בעיות. מאט את הריצה.")
        ToolTip(self.iai_check, "שלב 1 ישתמש ב-OCR על הכותרת האדומה של IAI כגיבוי")
        ToolTip(self.retries_entry, "כמה ניסיונות חוזרים לשלב 1 אם לא מצא מספר פריט (1-5)")
        ToolTip(self.dpi_combo, "רזולוציה לסריקת PDF. 200=מהיר, 300=מדויק")
        ToolTip(self.log_size_entry, "כשהלוג עובר גודל זה — מתחיל קובץ חדש")
        ToolTip(self.ils_entry, "לחישוב עלות בשקלים ב-Dashboard")

    def _open_dashboard(self) -> None:
        DashboardWindow(self.parent)

    def _open_manual_gui(self) -> None:
        """Open manual extraction GUI."""
        # If launched from automation_main.py, use stored reference
        if hasattr(self, '_open_manual_gui_func'):
            self._open_manual_gui_func()
            return
        
        try:
            from customer_extractor_gui import ExtractorGUI
            import customtkinter as ctk
            
            manual_window = ctk.CTkToplevel(self.winfo_toplevel())
            manual_window.title("🔧 Green Coat — DrawingAI Pro — עיבוד ידני")
            manual_window.geometry("900x900")
            manual_window.minsize(750, 850)
            manual_window.resizable(True, True)
            
            app = ExtractorGUI(manual_window)
            
        except Exception as e:
            from tkinter import messagebox
            messagebox.showerror("שגיאה", f"שגיאה בפתיחת GUI ידני:\n{e}")

    def _browse_download(self) -> None:
        folder = filedialog.askdirectory(title="בחר תיקיית הורדה")
        if folder:
            self.download_root_var.set(folder)

    def _browse_tosend(self) -> None:
        folder = filedialog.askdirectory(title="בחר תיקיית TO_SEND")
        if folder:
            self.tosend_folder_var.set(folder)

    def _browse_output_copy(self) -> None:
        folder = filedialog.askdirectory(title="בחר תיקיית שמירה")
        if folder:
            self.output_copy_folder_var.set(folder)

    def _load_config(self) -> None:
        if not self.config_path.exists():
            return
        try:
            with open(self.config_path, "r", encoding="utf-8") as f:
                cfg = json.load(f)
        except Exception:
            return

        configured_mailboxes = cfg.get("shared_mailboxes")
        if isinstance(configured_mailboxes, list) and configured_mailboxes:
            parsed = [str(x) for x in configured_mailboxes]
            self.shared_mailbox_var.set(self._format_mailboxes_for_ui(parsed))
            self._refresh_mailbox_selector(parsed)
        else:
            single_mailbox = cfg.get("shared_mailbox", self.shared_mailbox_var.get())
            self.shared_mailbox_var.set(single_mailbox)
            self._refresh_mailbox_selector([single_mailbox] if str(single_mailbox).strip() else [])
        self.folder_name_var.set(cfg.get("folder_name", self.folder_name_var.get()))
        self.rerun_folder_var.set(cfg.get("rerun_folder_name", ""))
        self.rerun_mailbox_var.set(cfg.get("rerun_mailbox", ""))
        skip_list = cfg.get("skip_senders", [])
        if not isinstance(skip_list, list):
            skip_list = [s.strip() for s in re.split(r'[;,\n]+', str(skip_list)) if s.strip()]
        self.skip_senders_listbox.delete(0, tk.END)
        for addr in skip_list:
            self.skip_senders_listbox.insert(tk.END, addr.strip().lower())
        skip_cats = cfg.get("skip_categories", [])
        if not isinstance(skip_cats, list):
            skip_cats = [c.strip() for c in re.split(r'[;,\n]+', str(skip_cats)) if c.strip()]
        self.skip_categories_listbox.delete(0, tk.END)
        for cat in skip_cats:
            self.skip_categories_listbox.insert(tk.END, cat.strip())
        self.scan_from_date_var.set(cfg.get("scan_from_date", ""))
        self.recipient_email_var.set(cfg.get("recipient_email", self.recipient_email_var.get()))
        self.download_root_var.set(cfg.get("download_root", self.download_root_var.get()))
        self.tosend_folder_var.set(cfg.get("tosend_folder", ""))
        self.output_copy_folder_var.set(cfg.get("output_copy_folder", ""))
        self.interval_var.set(str(cfg.get("poll_interval_minutes", self.interval_var.get())))
        self.max_messages_var.set(str(cfg.get("max_messages", self.max_messages_var.get())))
        self.max_files_per_email_var.set(str(cfg.get("max_files_per_email", 15)))
        self.max_file_size_var.set(str(cfg.get("max_file_size_mb", self.max_file_size_var.get())))
        self.stage1_skip_retry_resolution_var.set(str(cfg.get("stage1_skip_retry_resolution_px", 8000)))
        stored_dim = cfg.get("max_image_dimension", 3072)
        # Find matching display label for the stored dimension
        display_label = None
        for label, value in self._resolution_map.items():
            if value == stored_dim:
                display_label = label
                break
        if display_label is None:
            display_label = "3072 (מאזן - איכות מעולה)"
        self.max_image_dim_var.set(display_label)
        self.recursive_var.set(bool(cfg.get("recursive", True)))
        self.enable_retry_var.set(bool(cfg.get("enable_retry", True)))
        self.auto_start_var.set(bool(cfg.get("auto_start", False)))
        self.auto_send_var.set(bool(cfg.get("auto_send", False)))
        self.archive_full_var.set(bool(cfg.get("archive_full", False)))
        self.cleanup_download_var.set(bool(cfg.get("cleanup_download", True)))
        self.mark_as_processed_var.set(bool(cfg.get("mark_as_processed", True)))
        self.mark_category_name_var.set(cfg.get("mark_category_name", "AI Processed"))
        stored_color = cfg.get("mark_category_color", "preset0")
        reverse_color_map = {v: k for k, v in getattr(self, "_category_color_map", {}).items()}
        self.mark_category_color_var.set(reverse_color_map.get(stored_color, "None"))

        self.nodraw_category_name_var.set(cfg.get("nodraw_category_name", "NO DRAW"))
        stored_nodraw_color = cfg.get("nodraw_category_color", "preset1")
        self.nodraw_category_color_var.set(reverse_color_map.get(stored_nodraw_color, "None"))

        stages = cfg.get("selected_stages", {})
        self.stage1_var.set(bool(stages.get("1", True)))
        self.stage2_var.set(bool(stages.get("2", True)))
        self.stage3_var.set(bool(stages.get("3", True)))
        self.stage4_var.set(bool(stages.get("4", True)))
        self.stage5_var.set(bool(stages.get("5", True)))
        self.stage6_var.set(bool(stages.get("6", True)))
        self.stage7_var.set(bool(stages.get("7", True)))
        self.stage8_var.set(bool(stages.get("8", True)))
        self.stage9_var.set(bool(stages.get("9", True)))

        # Per-stage model overrides
        stage_models = cfg.get("stage_models", {})
        for n in range(10):
            model_val = stage_models.get(str(n), "")
            if model_val and model_val in self._available_models:
                self.stage_model_vars[n].set(model_val)
            else:
                self.stage_model_vars[n].set(self._env_stage_defaults.get(n, "gpt-4o-vision"))
        
        # Load confidence level for B2B files
        self.confidence_level_var.set(cfg.get("confidence_level", "LOW"))

        # Advanced settings
        self.debug_mode_var.set(bool(cfg.get("debug_mode", False)))
        self.iai_top_red_var.set(bool(cfg.get("iai_top_red_fallback", True)))
        self.max_retries_var.set(str(cfg.get("max_retries", 3)))
        self.scan_dpi_var.set(str(cfg.get("scan_dpi", 200)))
        self.log_max_size_var.set(str(cfg.get("log_max_size_mb", 1)))
        self.usd_to_ils_var.set(str(cfg.get("usd_to_ils_rate", 3.7)))

    def _save_config(self) -> None:
        parsed_mailboxes = self._parse_mailboxes_text(self.shared_mailbox_var.get().strip())
        primary_mailbox = parsed_mailboxes[0] if parsed_mailboxes else ""

        cfg: Dict[str, Any] = {
            "shared_mailbox": primary_mailbox,
            "shared_mailboxes": parsed_mailboxes,
            "folder_name": self._normalize_folder_name(self.folder_name_var.get().strip()),
            "rerun_folder_name": self.rerun_folder_var.get().strip(),
            "rerun_mailbox": self.rerun_mailbox_var.get().strip(),
            "skip_senders": list(self.skip_senders_listbox.get(0, tk.END)),
            "skip_categories": list(self.skip_categories_listbox.get(0, tk.END)),
            "scan_from_date": self.scan_from_date_var.get().strip(),
            "recipient_email": self.recipient_email_var.get().strip(),
            "download_root": self.download_root_var.get().strip(),
            "tosend_folder": self.tosend_folder_var.get().strip(),
            "output_copy_folder": self.output_copy_folder_var.get().strip(),
            "poll_interval_minutes": int(self.interval_var.get() or 10),
            "max_messages": int(self.max_messages_var.get() or 200),
            "max_files_per_email": max(int(self.max_files_per_email_var.get() or 0), 0),
            "max_file_size_mb": int(self.max_file_size_var.get() or 100),
            "stage1_skip_retry_resolution_px": max(int(self.stage1_skip_retry_resolution_var.get() or 8000), 0),
            "max_image_dimension": self._resolution_map.get(self.max_image_dim_var.get(), 3072),
            "recursive": self.recursive_var.get(),
            "enable_retry": self.enable_retry_var.get(),
            "auto_start": self.auto_start_var.get(),
            "auto_send": self.auto_send_var.get(),
            "archive_full": self.archive_full_var.get(),
            "cleanup_download": self.cleanup_download_var.get(),
            "mark_as_processed": self.mark_as_processed_var.get(),
            "mark_category_name": self.mark_category_name_var.get().strip() or "AI Processed",
            "mark_category_color": self._category_color_map.get(
                self.mark_category_color_var.get().strip(),
                "preset0"
            ),
            "nodraw_category_name": self.nodraw_category_name_var.get().strip() or "NO DRAW",
            "nodraw_category_color": self._category_color_map.get(
                self.nodraw_category_color_var.get().strip(),
                "preset1"
            ),
            "heavy_category_name": "AI HEAVY",
            "heavy_category_color": "preset4",
            "confidence_level": self.confidence_level_var.get(),
            "debug_mode": self.debug_mode_var.get(),
            "iai_top_red_fallback": self.iai_top_red_var.get(),
            "max_retries": min(max(int(self.max_retries_var.get() or 3), 1), 5),
            "scan_dpi": int(self.scan_dpi_var.get() or 200),
            "log_max_size_mb": max(int(self.log_max_size_var.get() or 1), 1),
            "usd_to_ils_rate": float(self.usd_to_ils_var.get() or 3.7),
            "selected_stages": {
                "1": self.stage1_var.get(),
                "2": self.stage2_var.get(),
                "3": self.stage3_var.get(),
                "4": self.stage4_var.get(),
                "5": self.stage5_var.get(),
                "6": self.stage6_var.get(),
                "7": self.stage7_var.get(),
                "8": self.stage8_var.get(),
                "9": self.stage9_var.get()
            },
            "stage_models": {
                str(n): self.stage_model_vars[n].get()
                for n in range(10)
                if self.stage_model_vars[n].get()
            }
        }

        try:
            with open(self.config_path, "w", encoding="utf-8") as f:
                json.dump(cfg, f, ensure_ascii=False, indent=2)
            messagebox.showinfo("נשמר", f"ההגדרות נשמרו בהצלחה\n\n"
                              f"מקסימום גודל קובץ: {cfg['max_file_size_mb']}MB\n"
                              f"מקסימום רזולוציה תמונה: {cfg['max_image_dimension']}px\n"
                              f"סף דילוג Retry שלב 1: {cfg['stage1_skip_retry_resolution_px']}px\n\n"
                              f"קבצים שחוצים את המקסימום יידלגו בניתוח!")
        except Exception as e:
            messagebox.showerror("שגיאה", f"שגיאה בשמירת ההגדרות: {e}")

    def _run_once(self) -> None:
        self._save_config()
        self._set_status("▶️ מריץ סבב אחד...")
        threading.Thread(target=self.runner.run_once, daemon=True).start()

    def _run_heavy(self) -> None:
        """Process only emails marked AI HEAVY (no file-count threshold)."""
        self._save_config()
        self._set_status("🏋️ מריץ עיבוד מיילים כבדים...")
        threading.Thread(target=self.runner.run_heavy, daemon=True).start()

    def _test_mailboxes(self) -> None:
        mailboxes = self._parse_mailboxes_text(self.shared_mailbox_var.get().strip())
        if not mailboxes:
            messagebox.showwarning("אזהרה", "לא הוגדרה אף תיבה לבדיקה.")
            return

        self._set_status("בודק חיבור לתיבות...")
        ok_mailboxes = []
        failed_mailboxes = []

        for mailbox in mailboxes:
            try:
                helper = GraphAPIHelper(shared_mailbox=mailbox)
                if helper.test_connection():
                    ok_mailboxes.append(mailbox)
                else:
                    failed_mailboxes.append(mailbox)
            except Exception:
                failed_mailboxes.append(mailbox)

        self._refresh_mailbox_selector(mailboxes)
        if ok_mailboxes:
            selected = self.mailbox_select_var.get().strip()
            preferred = selected if selected in ok_mailboxes else ok_mailboxes[0]
            self.mailbox_select_var.set(preferred)

        summary_lines = [f"סה\"כ תיבות: {len(mailboxes)}", f"✅ הצליחו: {len(ok_mailboxes)}", f"❌ נכשלו: {len(failed_mailboxes)}"]
        summary_lines.append("\nלטעינת תיקיות: בחר/י תיבה ולחץ/י 'טען תיקיות'.")
        if ok_mailboxes:
            summary_lines.append("\nתיבות תקינות:")
            summary_lines.extend([f"- {m}" for m in ok_mailboxes])
        if failed_mailboxes:
            summary_lines.append("\nתיבות שנכשלו:")
            summary_lines.extend([f"- {m}" for m in failed_mailboxes])

        if failed_mailboxes:
            self._set_status("בדיקת תיבות הושלמה עם שגיאות")
            messagebox.showwarning("תוצאות בדיקה", "\n".join(summary_lines))
        else:
            self._set_status("כל התיבות תקינות")
            messagebox.showinfo("תוצאות בדיקה", "\n".join(summary_lines))

    def _start(self) -> None:
        self._save_config()
        self.runner.start()
        self._set_status("🟢 אוטומציה פעילה")
        self._set_status_running()
        self._start_timer()

    def _stop(self) -> None:
        self.runner.stop()
        self._set_status("🔴 אוטומציה נעצרה")
        self._set_status_stopped()
        self._stop_timer()

    def _reset_state(self) -> None:
        """מחק את automation_state.json - התחל מחדש עם כל המיילים"""
        if not messagebox.askyesno("Reset State", "זה יהפוך את כל המיילים לחדשים!\n\nהמשיך?"):
            return
        
        try:
            if self.state_path.exists():
                self.state_path.unlink()  # מחק את הקובץ
            self._set_status("✓ State reset - כל המיילים יתחשבו כחדשים")
            messagebox.showinfo("Reset", "State נמחק בהצלחה!\n\nבסבב הבא, כל המיילים יתחשבו כחדשים.")
        except Exception as e:
            messagebox.showerror("שגיאה", f"שגיאה ב-Reset: {e}")
            self._set_status(f"שגיאה: {e}")
