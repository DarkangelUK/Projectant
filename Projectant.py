import customtkinter as ctk
import sqlite3
import json
import csv
from tkinter import filedialog, messagebox, simpledialog, font as tkfont
import tkinter as tk
from datetime import datetime, timedelta
import webbrowser
from tkcalendar import Calendar, DateEntry
import os
import tempfile
import ctypes
from ctypes import wintypes

def _get_documents_folder():
    """
    Returns the absolute path to the user's Documents folder,
    correctly handling OneDrive redirection on Windows.
    """
    if os.name == 'nt':
        CSIDL_PERSONAL = 5
        SHGFP_TYPE_CURRENT = 0

        buf = ctypes.create_unicode_buffer(wintypes.MAX_PATH)
        ctypes.windll.shell32.SHGetFolderPathW(None, CSIDL_PERSONAL, None, SHGFP_TYPE_CURRENT, buf)
        
        if buf.value:
            return buf.value

    return os.path.join(os.path.expanduser("~"), "Documents")

# --- Time Formatting Helper ---
def format_timedelta(td: timedelta) -> str:
    """Formats a timedelta object into Hhrs Mmins Ssecs string."""
    total_seconds = int(td.total_seconds())
    if total_seconds < 0:
        return "0hrs 0mins"
    
    hours, remainder = divmod(total_seconds, 3600)
    minutes, seconds = divmod(remainder, 60)
    
    parts = []
    if hours > 0:
        parts.append(f"{hours}hrs")
    if minutes > 0:
        parts.append(f"{minutes}mins")
    if total_seconds < 60 and seconds > 0:
        parts.append(f"{seconds}secs")
        
    return " ".join(parts) if parts else "0mins"

# --- Time Tracking Helper Functions ---
def calculate_total_pause_time(pauses_json: str) -> timedelta:
    """Calculates the total duration of all completed pauses from a JSON string."""
    if not pauses_json:
        return timedelta(0)
    
    try:
        pauses = json.loads(pauses_json)
    except json.JSONDecodeError:
        return timedelta(0)
    
    total_pause_time = timedelta(0)
    for pause in pauses:
        if 'start' in pause and 'end' in pause:
            try:
                start = datetime.strptime(pause['start'], "%Y-%m-%d %H:%M:%S.%f")
                end = datetime.strptime(pause['end'], "%Y-%m-%d %H:%M:%S.%f")
                total_pause_time += (end - start)
            except ValueError:
                continue
    return total_pause_time

def calculate_actual_duration(start_time_str: str, end_time_str: str, pauses_json: str) -> int:
    """Calculates the duration in seconds, excluding pauses."""
    try:
        start_dt = datetime.strptime(start_time_str, "%Y-%m-%d %H:%M:%S")
        end_dt = datetime.strptime(end_time_str, "%Y-%m-%d %H:%M:%S")
    except (ValueError, TypeError):
        return 0

    total_time = end_dt - start_dt
    total_pause = calculate_total_pause_time(pauses_json)
    
    actual_duration = total_time - total_pause
    return max(0, int(actual_duration.total_seconds()))

# --- Windows HTML Clipboard Helper ---
def copy_html_to_clipboard(html_content: str, plain_text_fallback: str):
    """
    Registers both HTML and Plain Text formats into the Windows clipboard
    so pasting into Outlook, Teams, or webmail renders rich formatting automatically.
    """
    if os.name != 'nt':
        return False
    try:
        GMEM_MOVEABLE = 0x0002
        CF_UNICODETEXT = 13
        CF_HTML = ctypes.windll.user32.RegisterClipboardFormatW("HTML Format")

        header_template = (
            "Version:0.9\r\n"
            "StartHTML:{:08d}\r\n"
            "EndHTML:{:08d}\r\n"
            "StartFragment:{:08d}\r\n"
            "EndFragment:{:08d}\r\n"
        )
        fragment_start = "<!--StartFragment-->"
        fragment_end = "<!--EndFragment-->"
        
        full_html = f"<html><body>{fragment_start}{html_content}{fragment_end}</body></html>"
        dummy_header = header_template.format(0, 0, 0, 0)
        
        start_html = len(dummy_header.encode('utf-8'))
        start_fragment = start_html + len(f"<html><body>{fragment_start}".encode('utf-8'))
        end_fragment = start_fragment + len(html_content.encode('utf-8'))
        end_html = start_fragment + len(f"{html_content}{fragment_end}</body></html>".encode('utf-8'))
        
        clipboard_payload = header_template.format(start_html, end_html, start_fragment, end_fragment) + full_html
        payload_bytes = clipboard_payload.encode('utf-8') + b'\x00'

        h_html_mem = ctypes.windll.kernel32.GlobalAlloc(GMEM_MOVEABLE, len(payload_bytes))
        p_html = ctypes.windll.kernel32.GlobalLock(h_html_mem)
        ctypes.cdll.msvcrt.memcpy(ctypes.c_void_p(p_html), payload_bytes, len(payload_bytes))
        ctypes.windll.kernel32.GlobalUnlock(h_html_mem)

        text_bytes = (plain_text_fallback + '\x00').encode('utf-16-le')
        h_text_mem = ctypes.windll.kernel32.GlobalAlloc(GMEM_MOVEABLE, len(text_bytes))
        p_text = ctypes.windll.kernel32.GlobalLock(h_text_mem)
        ctypes.cdll.msvcrt.memcpy(ctypes.c_void_p(p_text), text_bytes, len(text_bytes))
        ctypes.windll.kernel32.GlobalUnlock(h_text_mem)

        if ctypes.windll.user32.OpenClipboard(None):
            ctypes.windll.user32.EmptyClipboard()
            ctypes.windll.user32.SetClipboardData(CF_HTML, h_html_mem)
            ctypes.windll.user32.SetClipboardData(CF_UNICODETEXT, h_text_mem)
            ctypes.windll.user32.CloseClipboard()
            return True
    except Exception:
        pass
    return False


class CustomDialog(ctk.CTkToplevel):
    def __init__(self, parent, title, prompt, initial_value=""):
        super().__init__(parent)
        self.title(title)
        self.geometry("400x150")
        self.transient(parent)
        self.grab_set()
        self.result = None
        
        parent_x = parent.winfo_x()
        parent_y = parent.winfo_y()
        parent_width = parent.winfo_width()
        parent_height = parent.winfo_height()
        dialog_width = 400
        dialog_height = 150
        x = parent_x + (parent_width - dialog_width) // 2
        y = parent_y + (parent_height - dialog_height) // 2
        self.geometry(f"{dialog_width}x{dialog_height}+{x}+{y}")

        self.label = ctk.CTkLabel(self, text=prompt)
        self.label.pack(pady=20, padx=20)
        self.entry = ctk.CTkEntry(self, width=360)
        self.entry.insert(0, initial_value)
        self.entry.pack(pady=0, padx=20)
        self.entry.bind("<Return>", self.ok_event)
        button_frame = ctk.CTkFrame(self, fg_color="transparent")
        button_frame.pack(pady=20)
        self.ok_button = ctk.CTkButton(button_frame, text="OK", command=self.ok_event)
        self.ok_button.pack(side="left", padx=10)
        self.cancel_button = ctk.CTkButton(button_frame, text="Cancel", command=self.cancel_event)
        self.cancel_button.pack(side="right", padx=10)
        self.entry.focus_set()
        self.wait_window()

    def ok_event(self, event=None):
        self.result = self.entry.get()
        self.destroy()

    def cancel_event(self):
        self.result = None
        self.destroy()


# --- Task Creation Dialog with Structured Date Options ---
class TaskCreateDialog(ctk.CTkToplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("New Task")
        self.geometry("460x340")
        self.transient(parent)
        self.grab_set()
        self.result = None

        parent_x = parent.winfo_x()
        parent_y = parent.winfo_y()
        parent_width = parent.winfo_width()
        parent_height = parent.winfo_height()
        dialog_width = 460
        dialog_height = 340
        x = parent_x + (parent_width - dialog_width) // 2
        y = parent_y + (parent_height - dialog_height) // 2
        self.geometry(f"{dialog_width}x{dialog_height}+{x}+{y}")

        self.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(self, text="Task Description:", font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, sticky="w", padx=20, pady=(15, 2))
        self.entry_desc = ctk.CTkEntry(self, width=420, placeholder_text="e.g., Review project roadmap and deliverables")
        self.entry_desc.grid(row=1, column=0, sticky="ew", padx=20, pady=(0, 10))

        ctk.CTkLabel(self, text="Timeline / Scheduling:", font=ctk.CTkFont(weight="bold")).grid(row=2, column=0, sticky="w", padx=20, pady=(5, 2))
        
        self.schedule_mode = ctk.StringVar(value="none")
        mode_frame = ctk.CTkFrame(self, fg_color="transparent")
        mode_frame.grid(row=3, column=0, sticky="w", padx=20, pady=(0, 10))

        r1 = ctk.CTkRadioButton(mode_frame, text="No Dates", variable=self.schedule_mode, value="none", command=self._update_date_pickers)
        r1.pack(side="left", padx=(0, 15))
        r2 = ctk.CTkRadioButton(mode_frame, text="Due Date Only", variable=self.schedule_mode, value="due", command=self._update_date_pickers)
        r2.pack(side="left", padx=15)
        r3 = ctk.CTkRadioButton(mode_frame, text="Date Range", variable=self.schedule_mode, value="range", command=self._update_date_pickers)
        r3.pack(side="left", padx=15)

        self.dates_frame = ctk.CTkFrame(self, fg_color="#24292E", corner_radius=6)
        self.dates_frame.grid(row=4, column=0, sticky="ew", padx=20, pady=(0, 15))

        self.lbl_start = ctk.CTkLabel(self.dates_frame, text="Start Date:", font=ctk.CTkFont(size=12))
        self.cal_start = DateEntry(self.dates_frame, width=12, background="#2563EB", foreground="white", borderwidth=2, date_pattern='yyyy-mm-dd')
        
        self.lbl_due = ctk.CTkLabel(self.dates_frame, text="Due / End Date:", font=ctk.CTkFont(size=12))
        self.cal_due = DateEntry(self.dates_frame, width=12, background="#2563EB", foreground="white", borderwidth=2, date_pattern='yyyy-mm-dd')

        self._update_date_pickers()

        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.grid(row=5, column=0, sticky="e", padx=20, pady=(5, 15))

        btn_cancel = ctk.CTkButton(btn_frame, text="Cancel", width=80, fg_color="#475569", hover_color="#334155", command=self.cancel_event)
        btn_cancel.pack(side="right", padx=(10, 0))

        btn_save = ctk.CTkButton(btn_frame, text="Save Task", width=100, fg_color="#2E8B57", hover_color="#20603C", command=self.save_event)
        btn_save.pack(side="right")

        self.entry_desc.focus_set()
        self.wait_window()

    def _update_date_pickers(self):
        for w in self.dates_frame.winfo_children():
            w.grid_forget()

        mode = self.schedule_mode.get()
        if mode == "none":
            lbl_none = ctk.CTkLabel(self.dates_frame, text="Task has no scheduled deadlines.", text_color="#94A3B8")
            lbl_none.grid(row=0, column=0, padx=15, pady=10)
        elif mode == "due":
            self.lbl_due.grid(row=0, column=0, padx=(15, 5), pady=10, sticky="w")
            self.cal_due.grid(row=0, column=1, padx=(0, 15), pady=10, sticky="w")
        elif mode == "range":
            self.lbl_start.grid(row=0, column=0, padx=(15, 5), pady=10, sticky="w")
            self.cal_start.grid(row=0, column=1, padx=(0, 15), pady=10, sticky="w")
            self.lbl_due.grid(row=0, column=2, padx=(15, 5), pady=10, sticky="w")
            self.cal_due.grid(row=0, column=3, padx=(0, 15), pady=10, sticky="w")

    def save_event(self):
        desc = self.entry_desc.get().strip()
        if not desc:
            messagebox.showerror("Validation Error", "Task description cannot be blank.")
            return

        mode = self.schedule_mode.get()
        start_date = None
        due_date = None

        if mode == "due":
            due_date = self.cal_due.get_date().strftime("%Y-%m-%d")
        elif mode == "range":
            start_date = self.cal_start.get_date().strftime("%Y-%m-%d")
            due_date = self.cal_due.get_date().strftime("%Y-%m-%d")
            if start_date > due_date:
                messagebox.showerror("Validation Error", "Start Date cannot be after Due Date.")
                return

        self.result = {
            "description": desc,
            "start_date": start_date,
            "due_date": due_date
        }
        self.destroy()

    def cancel_event(self):
        self.result = None
        self.destroy()


# --- Time Entry Edit Dialog ---
class TimeEntryEditDialog(ctk.CTkToplevel):
    def __init__(self, parent, entry_data):
        super().__init__(parent)
        self.title("Edit Time Entry")
        self.geometry("400x400")
        self.transient(parent)
        self.grab_set()
        self.result = None

        self.entry_data = entry_data
        
        parent_x = parent.winfo_x(); parent_y = parent.winfo_y()
        parent_width = parent.winfo_width(); parent_height = parent.winfo_height()
        dialog_width = 400; dialog_height = 400
        x = parent_x + (parent_width - dialog_width) // 2
        y = parent_y + (parent_height - dialog_height) // 2
        self.geometry(f"{dialog_width}x{dialog_height}+{x}+{y}")
        
        self.grid_columnconfigure(0, weight=1)
        
        ctk.CTkLabel(self, text="Start Time (YYYY-MM-DD HH:MM:SS):").grid(row=0, column=0, padx=20, pady=(10, 0), sticky="w")
        self.entry_start = ctk.CTkEntry(self, width=360)
        self.entry_start.insert(0, entry_data[1])
        self.entry_start.grid(row=1, column=0, padx=20, pady=(0, 10), sticky="ew")

        ctk.CTkLabel(self, text="End Time (YYYY-MM-DD HH:MM:SS, optional):").grid(row=2, column=0, padx=20, pady=(10, 0), sticky="w")
        self.entry_end = ctk.CTkEntry(self, width=360)
        self.entry_end.insert(0, entry_data[2] if entry_data[2] else "")
        self.entry_end.grid(row=3, column=0, padx=20, pady=(0, 10), sticky="ew")
        
        pauses_td = calculate_total_pause_time(entry_data[5])
        ctk.CTkLabel(self, text=f"Total Pause Time: {format_timedelta(pauses_td)}").grid(row=4, column=0, padx=20, pady=(5, 5), sticky="w")
        
        ctk.CTkLabel(self, text="Notes:").grid(row=5, column=0, padx=20, pady=(10, 0), sticky="w")
        self.txt_notes = ctk.CTkTextbox(self, height=80)
        self.txt_notes.insert("1.0", entry_data[4])
        self.txt_notes.grid(row=6, column=0, padx=20, pady=(0, 10), sticky="ew")

        button_frame = ctk.CTkFrame(self, fg_color="transparent")
        button_frame.grid(row=7, column=0, pady=20)
        
        self.save_button = ctk.CTkButton(button_frame, text="Save Changes", command=self.save_event)
        self.save_button.pack(side="left", padx=10)
        self.cancel_button = ctk.CTkButton(button_frame, text="Cancel", command=self.cancel_event)
        self.cancel_button.pack(side="right", padx=10)
        
        self.entry_start.focus_set()
        self.wait_window()

    def save_event(self, event=None):
        start_str = self.entry_start.get().strip()
        end_str = self.entry_end.get().strip()
        notes = self.txt_notes.get("1.0", "end-1c").strip()
        
        try:
            start_dt = datetime.strptime(start_str, "%Y-%m-%d %H:%M:%S")
        except ValueError:
            messagebox.showerror("Validation Error", "Invalid Start Time format. Use YYYY-MM-DD HH:MM:SS.")
            return

        end_dt = None
        duration_seconds = 0
        pauses_json = self.entry_data[5]

        if end_str:
            try:
                end_dt = datetime.strptime(end_str, "%Y-%m-%d %H:%M:%S")
                if end_dt < start_dt:
                    messagebox.showerror("Validation Error", "End Time cannot be before Start Time.")
                    return
                duration_seconds = calculate_actual_duration(start_str, end_str, pauses_json)
            except ValueError:
                messagebox.showerror("Validation Error", "Invalid End Time format. Use YYYY-MM-DD HH:MM:SS.")
                return
        
        start_time_str = start_dt.strftime("%Y-%m-%d %H:%M:%S")
        end_time_str = end_dt.strftime("%Y-%m-%d %H:%M:%S") if end_dt else None
        
        self.result = {
            "id": self.entry_data[0],
            "start_time": start_time_str,
            "end_time": end_time_str,
            "duration_seconds": duration_seconds,
            "notes": notes,
            "pauses_json": pauses_json
        }
        self.destroy()

    def cancel_event(self):
        self.result = None
        self.destroy()


# --- Report Display Window ---
class ReportDisplayDialog(ctk.CTkToplevel):
    def __init__(self, parent, report_text, report_html):
        super().__init__(parent)
        self.title("Executive Weekly Summary")
        self.geometry("820x680")
        self.transient(parent)
        self.grab_set()

        self.report_text = report_text
        self.report_html = report_html

        parent_x = parent.winfo_x(); parent_y = parent.winfo_y()
        parent_width = parent.winfo_width(); parent_height = parent.winfo_height()
        dialog_width = 820; dialog_height = 680
        x = parent_x + (parent_width - dialog_width) // 2
        y = parent_y + (parent_height - dialog_height) // 2
        self.geometry(f"{dialog_width}x{dialog_height}+{x}+{y}")

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=0)

        self.txt_report = ctk.CTkTextbox(self, font=ctk.CTkFont(family="Segoe UI", size=13))
        self.txt_report.insert("1.0", self.report_text)
        self.txt_report.grid(row=0, column=0, sticky="nsew", padx=15, pady=(15, 10))

        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.grid(row=1, column=0, sticky="ew", padx=15, pady=(0, 15))

        btn_copy = ctk.CTkButton(btn_frame, text="Copy for Email (Rich Format)", fg_color="#2E8B57", hover_color="#20603C",
                                 command=self.copy_to_clipboard, height=32, font=ctk.CTkFont(weight="bold"))
        btn_copy.pack(side="left", padx=(0, 10))

        btn_close = ctk.CTkButton(btn_frame, text="Close", command=self.destroy, height=32)
        btn_close.pack(side="right")

    def copy_to_clipboard(self):
        edited_text = self.txt_report.get("1.0", "end-1c")
        success = copy_html_to_clipboard(self.report_html, edited_text)
        if not success:
            self.clipboard_clear()
            self.clipboard_append(edited_text)
        messagebox.showinfo("Copied", "Formatted summary copied. You can now paste directly into Outlook or your email client.")


class ProjectApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Projectant - Project Support App")
        
        self.geometry("1300x800")
        self.minsize(1100, 700)
        self.db_path = None
        self.current_note_id = None
        self.current_table_id = None
        self.table_widgets = []
        self.original_save_button_color = None

        self.note_save_timer = None
        self.table_save_timer = None
        self.auto_save_delay_ms = 1500
        
        self.projects_dir = os.path.join(_get_documents_folder(), "ProjectantProjects")
        os.makedirs(self.projects_dir, exist_ok=True)
        
        # --- Time Tracking Variables ---
        self.timer_running = False
        self.timer_paused = False 
        self.start_time = None
        self.pauses = [] 
        self.timer_update_id = None
        self.timer_label = None
        self.btn_start_stop = None
        self.btn_pause_resume = None 
        self.current_running_entry_id = None

        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)
        
        self.left_frame = ctk.CTkFrame(self, width=260, corner_radius=0)
        self.left_frame.grid(row=0, column=0, rowspan=2, sticky="nsew")
        self.left_frame.grid_rowconfigure(2, weight=1)
        self.left_frame.grid_rowconfigure(5, weight=1)
        
        self.main_frame = ctk.CTkFrame(self)
        self.main_frame.grid(row=0, column=1, sticky="nsew", padx=20, pady=20)
        self.main_frame.grid_columnconfigure(0, weight=1)
        self.main_frame.grid_rowconfigure(1, weight=1)
        
        self.right_frame = ctk.CTkFrame(self, width=280)
        self.right_frame.grid(row=0, column=2, sticky="nsew", padx=(0, 20), pady=20)
        self.right_frame.grid_rowconfigure(0, weight=1)
        self.right_frame.grid_rowconfigure(1, weight=1)
        self.right_frame.grid_rowconfigure(2, weight=1)
        
        self.menu_bar = ctk.CTkFrame(self, height=30, corner_radius=0)
        self.menu_bar.grid(row=2, column=0, columnspan=3, sticky="ew")
        self.btn_new_project = ctk.CTkButton(self.menu_bar, text="New Project", command=self.new_project)
        self.btn_new_project.pack(side="left", padx=5, pady=5)
        self.btn_open_project = ctk.CTkButton(self.menu_bar, text="Open Project", command=self.open_project)
        self.btn_open_project.pack(side="left", padx=5, pady=5)
        
        self.project_variable = ctk.StringVar(value="Select a Project...")
        self.project_dropdown = ctk.CTkOptionMenu(self.menu_bar, variable=self.project_variable, 
                                                 command=self._load_project_from_dropdown)
        self.project_dropdown.pack(side="left", padx=5, pady=5)
        self._refresh_project_dropdown()

        self.btn_weekly_summary = ctk.CTkButton(self.menu_bar, text="Weekly Summary", fg_color="#6366F1", hover_color="#4F46E5",
                                                command=self.generate_weekly_summary)
        self.btn_weekly_summary.pack(side="left", padx=5, pady=5)
        
        self.project_name_label = ctk.CTkLabel(self.menu_bar, text="No Project Loaded")
        self.project_name_label.pack(side="right", padx=10)

        self.setup_ui_components()
        self.switch_to_notes_view()

    def _refresh_project_dropdown(self):
        try:
            project_files = [f for f in os.listdir(self.projects_dir) if f.endswith(".db")]
            if project_files:
                self.project_dropdown.configure(values=project_files, state="normal")
                if not self.db_path:
                    self.project_variable.set("Select a Project...")
            else:
                self.project_variable.set("No projects found")
                self.project_dropdown.configure(values=["No projects found"], state="disabled")
        except Exception:
            self.project_variable.set("Error listing projects")
            self.project_dropdown.configure(values=[], state="disabled")

    def _load_project_from_dropdown(self, selected_project_name: str):
        if selected_project_name in ["Select a Project...", "No projects found", "Error listing projects"]:
            return
        full_path = os.path.join(self.projects_dir, selected_project_name)
        if os.path.exists(full_path):
            self.init_db(full_path)
        else:
            messagebox.showerror("Error", f"Could not find project file: {selected_project_name}")
            self._refresh_project_dropdown()

    def new_project(self):
        path = filedialog.asksaveasfilename(defaultextension=".db", 
                                             filetypes=[("Project Files", "*.db")],
                                             initialdir=self.projects_dir)
        if path:    
            self.init_db(path)
            self._refresh_project_dropdown()
            self.project_variable.set(os.path.basename(path))

    def open_project(self):
        path = filedialog.askopenfilename(filetypes=[("Project Files", "*.db")],
                                             initialdir=self.projects_dir)
        if path:    
            self.init_db(path)

    def init_db(self, db_path):
        self.db_path = db_path
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        cursor.execute('''CREATE TABLE IF NOT EXISTS notes (
                            id INTEGER PRIMARY KEY, 
                            title TEXT NOT NULL, 
                            content TEXT, 
                            created_date TEXT, 
                            last_modified_date TEXT,
                            is_pinned INTEGER DEFAULT 0,
                            folder TEXT DEFAULT 'General'
                          )''')
        cursor.execute('''CREATE TABLE IF NOT EXISTS tasks (
                            id INTEGER PRIMARY KEY, 
                            task_description TEXT NOT NULL, 
                            start_date TEXT,
                            due_date TEXT, 
                            is_completed INTEGER DEFAULT 0,
                            completed_date TEXT,
                            completion_notes TEXT
                          )''')
        cursor.execute('''CREATE TABLE IF NOT EXISTS questions (
                            id INTEGER PRIMARY KEY, 
                            question_text TEXT NOT NULL, 
                            is_answered INTEGER DEFAULT 0,
                            answer_text TEXT,
                            answered_date TEXT
                          )''')
        cursor.execute('''CREATE TABLE IF NOT EXISTS links (id INTEGER PRIMARY KEY, link_name TEXT NOT NULL, url TEXT NOT NULL)''')
        cursor.execute('''CREATE TABLE IF NOT EXISTS custom_tables (id INTEGER PRIMARY KEY, name TEXT NOT NULL, data TEXT)''')
        cursor.execute('''CREATE TABLE IF NOT EXISTS events (id INTEGER PRIMARY KEY, event_date TEXT NOT NULL, description TEXT NOT NULL)''')
        cursor.execute('''CREATE TABLE IF NOT EXISTS attachments (
                             id INTEGER PRIMARY KEY,
                             note_id INTEGER NOT NULL,
                             filename TEXT NOT NULL,
                             file_data BLOB NOT NULL,
                             FOREIGN KEY (note_id) REFERENCES notes (id) ON DELETE CASCADE
                           )''')
        cursor.execute('''CREATE TABLE IF NOT EXISTS time_entries (
                             id INTEGER PRIMARY KEY,
                             start_time TEXT NOT NULL,
                             end_time TEXT,
                             duration_seconds INTEGER DEFAULT 0,
                             notes TEXT,
                             pauses TEXT DEFAULT '[]'
                           )''')
        
        # MIGRATION STEPS
        try:
            cursor.execute("SELECT is_pinned FROM notes LIMIT 1")
        except sqlite3.OperationalError:
            cursor.execute("ALTER TABLE notes ADD COLUMN is_pinned INTEGER DEFAULT 0")

        try:
            cursor.execute("SELECT folder FROM notes LIMIT 1")
        except sqlite3.OperationalError:
            cursor.execute("ALTER TABLE notes ADD COLUMN folder TEXT DEFAULT 'General'")

        try:
            cursor.execute("SELECT pauses FROM time_entries LIMIT 1")
        except sqlite3.OperationalError:
            cursor.execute("ALTER TABLE time_entries ADD COLUMN pauses TEXT DEFAULT '[]'")
            
        try:
            cursor.execute("SELECT start_date FROM tasks LIMIT 1")
        except sqlite3.OperationalError:
            cursor.execute("ALTER TABLE tasks ADD COLUMN start_date TEXT")

        try:
            cursor.execute("SELECT completed_date FROM tasks LIMIT 1")
        except sqlite3.OperationalError:
            cursor.execute("ALTER TABLE tasks ADD COLUMN completed_date TEXT")

        try:
            cursor.execute("SELECT completion_notes FROM tasks LIMIT 1")
        except sqlite3.OperationalError:
            cursor.execute("ALTER TABLE tasks ADD COLUMN completion_notes TEXT")
            
        conn.commit()
        conn.close()
        
        project_filename = os.path.basename(db_path)
        self.project_name_label.configure(text=f"Project: {project_filename}")
        
        self._refresh_project_dropdown()
        self.project_variable.set(project_filename)
        
        self.load_project_data()
        self.check_for_running_timer() 

    def setup_ui_components(self):
        action_color = "#2E8B57"; hover_color = "#20603C"

        # Notes Header & Folder Selector
        notes_header = ctk.CTkFrame(self.left_frame, fg_color="transparent")
        notes_header.grid(row=0, column=0, padx=10, pady=(10, 2), sticky="ew")
        notes_header.grid_columnconfigure(0, weight=1)

        self.lbl_notes = ctk.CTkLabel(notes_header, text="Notes", font=ctk.CTkFont(size=16, weight="bold"))
        self.lbl_notes.grid(row=0, column=0, sticky="w")

        self.note_folder_var = ctk.StringVar(value="All Folders")
        self.opt_note_folder = ctk.CTkOptionMenu(
            self.left_frame, 
            variable=self.note_folder_var, 
            values=["All Folders", "General", "Status", "+ New Folder..."],
            command=self._on_folder_filter_change,
            height=24,
            font=ctk.CTkFont(size=11)
        )
        self.opt_note_folder.grid(row=1, column=0, padx=10, pady=(0, 5), sticky="ew")

        self.notes_frame = ctk.CTkScrollableFrame(self.left_frame)
        self.notes_frame.grid(row=2, column=0, padx=10, pady=(0, 5), sticky="nsew")
        self.btn_new_note = ctk.CTkButton(self.left_frame, text="New Note", command=self.switch_to_notes_view, fg_color=action_color, hover_color=hover_color)
        self.btn_new_note.grid(row=3, column=0, padx=10, pady=(5, 10), sticky="ew")

        # Tables Section
        self.lbl_tables = ctk.CTkLabel(self.left_frame, text="Tables", font=ctk.CTkFont(size=16, weight="bold"))
        self.lbl_tables.grid(row=4, column=0, padx=10, pady=(5, 0), sticky="ew")
        self.tables_frame = ctk.CTkScrollableFrame(self.left_frame)
        self.tables_frame.grid(row=5, column=0, padx=10, pady=(0, 5), sticky="nsew")
        self.btn_new_table = ctk.CTkButton(self.left_frame, text="New Table", command=self.create_new_table, fg_color=action_color, hover_color=hover_color)
        self.btn_new_table.grid(row=6, column=0, padx=10, pady=(5, 10), sticky="ew")

        # Navigation Buttons
        self.btn_todo = ctk.CTkButton(self.left_frame, text="To-Do List", command=self.switch_to_todo_view)
        self.btn_todo.grid(row=7, column=0, padx=10, pady=3, sticky="ew")

        self.btn_roadmap = ctk.CTkButton(self.left_frame, text="Roadmap", fg_color="#3B8ED0", hover_color="#2563EB", command=self.switch_to_roadmap_view)
        self.btn_roadmap.grid(row=8, column=0, padx=10, pady=3, sticky="ew")

        self.btn_questions = ctk.CTkButton(self.left_frame, text="Questions", command=self.switch_to_questions_view)
        self.btn_questions.grid(row=9, column=0, padx=10, pady=3, sticky="ew")

        self.btn_calendar = ctk.CTkButton(self.left_frame, text="Calendar", command=self.switch_to_calendar_view)
        self.btn_calendar.grid(row=10, column=0, padx=10, pady=3, sticky="ew")

        self.btn_timetracker = ctk.CTkButton(self.left_frame, text="Time Tracker", command=self.switch_to_timetracker_view)
        self.btn_timetracker.grid(row=11, column=0, padx=10, pady=(3, 10), sticky="ew")

        # Right Panel
        tasks_container = ctk.CTkFrame(self.right_frame); tasks_container.grid(row=0, column=0, sticky="nsew"); tasks_container.grid_rowconfigure(1, weight=1); tasks_container.grid_columnconfigure(0, weight=1)
        self.lbl_upcoming_tasks = ctk.CTkLabel(tasks_container, text="Upcoming Tasks", font=ctk.CTkFont(size=16, weight="bold")); self.lbl_upcoming_tasks.grid(row=0, column=0, sticky="ew", padx=5, pady=(10,5))
        self.upcoming_tasks_frame = ctk.CTkScrollableFrame(tasks_container); self.upcoming_tasks_frame.grid(row=1, column=0, sticky="nsew", padx=5, pady=5)
        
        links_container = ctk.CTkFrame(self.right_frame); links_container.grid(row=1, column=0, sticky="nsew"); links_container.grid_rowconfigure(1, weight=1); links_container.grid_columnconfigure(0, weight=1)
        self.lbl_links = ctk.CTkLabel(links_container, text="Project Links", font=ctk.CTkFont(size=16, weight="bold")); self.lbl_links.grid(row=0, column=0, sticky="ew", padx=5, pady=(10,5))
        self.links_frame = ctk.CTkScrollableFrame(links_container); self.links_frame.grid(row=1, column=0, sticky="nsew", padx=5, pady=5)
        self.entry_link_name = ctk.CTkEntry(links_container, placeholder_text="Link Name"); self.entry_link_name.grid(row=2, column=0, sticky="ew", padx=5, pady=5)
        self.entry_link_url = ctk.CTkEntry(links_container, placeholder_text="URL / Link Address"); self.entry_link_url.grid(row=3, column=0, sticky="ew", padx=5, pady=5)
        self.btn_add_link = ctk.CTkButton(links_container, text="Add Link", command=self.add_link, fg_color=action_color, hover_color=hover_color); self.btn_add_link.grid(row=4, column=0, sticky="ew", padx=5, pady=5)
        
        events_container = ctk.CTkFrame(self.right_frame); events_container.grid(row=2, column=0, sticky="nsew"); events_container.grid_rowconfigure(1, weight=1); events_container.grid_columnconfigure(0, weight=1)
        self.lbl_events = ctk.CTkLabel(events_container, text="Upcoming Events", font=ctk.CTkFont(size=16, weight="bold")); self.lbl_events.grid(row=0, column=0, sticky="ew", padx=5, pady=(10,5))
        self.upcoming_events_frame = ctk.CTkScrollableFrame(events_container); self.upcoming_events_frame.grid(row=1, column=0, sticky="nsew", padx=5, pady=5)

    def _get_known_folders(self):
        default_folders = ["General", "Status"]
        if not self.db_path:
            return default_folders
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        try:
            cursor.execute("SELECT DISTINCT folder FROM notes WHERE folder IS NOT NULL AND folder != ''")
            db_folders = [r[0] for r in cursor.fetchall()]
        except sqlite3.OperationalError:
            db_folders = []
        conn.close()
        combined = sorted(list(set(default_folders + db_folders)))
        return combined

    def _update_folder_dropdown_options(self):
        folders = self._get_known_folders()
        sidebar_options = ["All Folders"] + folders + ["+ New Folder..."]
        self.opt_note_folder.configure(values=sidebar_options)
        if hasattr(self, 'opt_current_note_folder') and self.opt_current_note_folder.winfo_exists():
            editor_options = folders + ["+ New Folder..."]
            self.opt_current_note_folder.configure(values=editor_options)

    def _on_folder_filter_change(self, selected_value: str):
        if selected_value == "+ New Folder...":
            new_f = CustomDialog(self, title="New Folder", prompt="Enter folder name:").result
            if new_f and new_f.strip():
                folder_clean = new_f.strip()
                self.note_folder_var.set(folder_clean)
                self._update_folder_dropdown_options()
            else:
                self.note_folder_var.set("All Folders")
        self.load_notes_list()

    def _on_editor_folder_change(self, selected_value: str):
        if selected_value == "+ New Folder...":
            new_f = CustomDialog(self, title="New Folder", prompt="Enter folder name:").result
            if new_f and new_f.strip():
                folder_clean = new_f.strip()
                self.current_note_folder_var.set(folder_clean)
                self._update_folder_dropdown_options()
            else:
                self.current_note_folder_var.set("General")
        self._schedule_note_save()

    def _create_note_click_handler(self, note_id):
        def handler(event):
            self.switch_to_notes_view(note_id)
        return handler

    def _create_link_click_handler(self, url):
        def handler(event):
            self.open_link(url)
        return handler

    def _create_attachment_open_handler(self, attachment_id, filename):
        def handler(event):
            self._open_attachment(attachment_id, filename)
        return handler

    def _schedule_note_save(self, event=None):
        if self.note_save_timer is not None:
            self.after_cancel(self.note_save_timer)
        self.note_save_timer = self.after(self.auto_save_delay_ms, self.save_note)

    def _schedule_table_save(self, event=None):
        if self.table_save_timer is not None:
            self.after_cancel(self.table_save_timer)
        self.table_save_timer = self.after(self.auto_save_delay_ms, lambda: self.save_table_data(show_feedback=False))

    # --- Rich Text Formatting Helpers for Notes ---
    def _setup_note_tags(self):
        text_widget = self.txt_note_content._textbox
        font_family = "Arial"
        font_size = 13

        text_widget.tag_configure("bold", font=(font_family, font_size, "bold"))
        text_widget.tag_configure("italic", font=(font_family, font_size, "italic"))
        text_widget.tag_configure("underline", font=(font_family, font_size, "underline"))
        text_widget.tag_configure("bold_italic", font=(font_family, font_size, "bold italic"))
        text_widget.tag_configure("bullet", lmargin1=25, lmargin2=40)

    def toggle_format(self, tag_name):
        text_widget = self.txt_note_content._textbox
        try:
            sel_start = text_widget.index("sel.first")
            sel_end = text_widget.index("sel.last")
        except Exception:
            return

        if tag_name in text_widget.tag_names(sel_start):
            text_widget.tag_remove(tag_name, sel_start, sel_end)
        else:
            text_widget.tag_add(tag_name, sel_start, sel_end)
            
        self._schedule_note_save()

    def toggle_bullet_line(self):
        text_widget = self.txt_note_content._textbox
        line_start = text_widget.index("insert linestart")
        line_end = text_widget.index("insert lineend")
        line_text = text_widget.get(line_start, line_end)

        if line_text.startswith("• "):
            text_widget.delete(line_start, f"{line_start} + 2 chars")
            text_widget.tag_remove("bullet", line_start, f"{line_start} lineend")
        else:
            text_widget.insert(line_start, "• ")
            text_widget.tag_add("bullet", line_start, f"{line_start} lineend")

        self._schedule_note_save()

    def _handle_note_return(self, event):
        text_widget = self.txt_note_content._textbox
        line_start = text_widget.index("insert linestart")
        line_end = text_widget.index("insert lineend")
        line_text = text_widget.get(line_start, line_end)

        if line_text.startswith("• "):
            if line_text.strip() == "•":
                text_widget.delete(line_start, line_end)
                text_widget.tag_remove("bullet", line_start, f"{line_start} lineend")
                return "break"
            else:
                text_widget.insert("insert", "\n• ")
                new_line_start = text_widget.index("insert linestart")
                text_widget.tag_add("bullet", new_line_start, f"{new_line_start} lineend")
                text_widget.see("insert")
                self._schedule_note_save()
                return "break"
        return None

    def _serialize_note_content(self) -> str:
        text_widget = self.txt_note_content._textbox
        plain_text = self.txt_note_content.get("1.0", "end-1c")
        
        formatted_tags = {}
        for tag in ["bold", "italic", "underline", "bullet"]:
            ranges = text_widget.tag_ranges(tag)
            tag_list = []
            for i in range(0, len(ranges), 2):
                tag_list.append((str(ranges[i]), str(ranges[i+1])))
            if tag_list:
                formatted_tags[tag] = tag_list

        data = {
            "version": "rich_v1",
            "text": plain_text,
            "tags": formatted_tags
        }
        return json.dumps(data)

    def _deserialize_note_content(self, stored_content: str):
        self.txt_note_content.delete("1.0", "end")
        if not stored_content:
            return

        text_widget = self.txt_note_content._textbox
        try:
            data = json.loads(stored_content)
            if isinstance(data, dict) and data.get("version") == "rich_v1":
                text_widget.insert("1.0", data.get("text", ""))
                tags = data.get("tags", {})
                for tag_name, ranges in tags.items():
                    for start_idx, end_idx in ranges:
                        try:
                            text_widget.tag_add(tag_name, start_idx, end_idx)
                        except Exception:
                            continue
                return
        except Exception:
            pass

        self.txt_note_content.insert("1.0", stored_content)
        lines = stored_content.split("\n")
        for i, line in enumerate(lines, 1):
            if line.startswith("• "):
                text_widget.tag_add("bullet", f"{i}.0", f"{i}.end")

    def switch_to_notes_view(self, note_id=None):
        if self.current_note_id == note_id and note_id is not None:
            return

        self.current_table_id = None
        self.clear_main_frame()
        self.main_frame.grid_columnconfigure(0, weight=1)
        self.main_frame.grid_rowconfigure(0, weight=0)
        self.main_frame.grid_rowconfigure(1, weight=0)
        self.main_frame.grid_rowconfigure(2, weight=1)
        self.main_frame.grid_rowconfigure(3, weight=0)
        self.main_frame.grid_rowconfigure(4, weight=1)

        title_folder_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        title_folder_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=(10, 2))
        title_folder_frame.grid_columnconfigure(0, weight=1)

        self.entry_note_title = ctk.CTkEntry(title_folder_frame, placeholder_text="Note Title")
        self.entry_note_title.grid(row=0, column=0, sticky="ew", padx=(0, 10))

        folder_init = self.note_folder_var.get()
        if folder_init in ["All Folders", "+ New Folder..."]:
            folder_init = "General"

        self.current_note_folder_var = ctk.StringVar(value=folder_init)
        known_folders = self._get_known_folders()
        self.opt_current_note_folder = ctk.CTkOptionMenu(
            title_folder_frame,
            variable=self.current_note_folder_var,
            values=known_folders + ["+ New Folder..."],
            command=self._on_editor_folder_change,
            width=130
        )
        self.opt_current_note_folder.grid(row=0, column=1, sticky="e")
        
        toolbar = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        toolbar.grid(row=1, column=0, sticky="ew", padx=10, pady=2)

        btn_bold = ctk.CTkButton(toolbar, text="B", width=36, font=ctk.CTkFont(weight="bold"), 
                                 command=lambda: self.toggle_format("bold"))
        btn_bold.pack(side="left", padx=(0, 4))

        btn_italic = ctk.CTkButton(toolbar, text="I", width=36, font=ctk.CTkFont(slant="italic"), 
                                   command=lambda: self.toggle_format("italic"))
        btn_italic.pack(side="left", padx=4)

        btn_underline = ctk.CTkButton(toolbar, text="U", width=36, font=ctk.CTkFont(underline=True), 
                                      command=lambda: self.toggle_format("underline"))
        btn_underline.pack(side="left", padx=4)

        btn_bullet = ctk.CTkButton(toolbar, text="• Bullet", width=70, 
                                   command=self.toggle_bullet_line)
        btn_bullet.pack(side="left", padx=4)

        self.txt_note_content = ctk.CTkTextbox(self.main_frame, font=ctk.CTkFont(size=13))
        self.txt_note_content.grid(row=2, column=0, sticky="nsew", padx=10, pady=5)
        self._setup_note_tags()

        self.entry_note_title.bind("<KeyRelease>", self._schedule_note_save)
        self.txt_note_content.bind("<KeyRelease>", self._schedule_note_save)
        self.txt_note_content.bind("<Return>", self._handle_note_return)
        self.txt_note_content.bind("<Control-b>", lambda e: [self.toggle_format("bold"), "break"][1])
        self.txt_note_content.bind("<Control-i>", lambda e: [self.toggle_format("italic"), "break"][1])
        self.txt_note_content.bind("<Control-u>", lambda e: [self.toggle_format("underline"), "break"][1])

        attachment_header_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        attachment_header_frame.grid(row=3, column=0, sticky="ew", padx=10, pady=5)
        
        self.lbl_attachments = ctk.CTkLabel(attachment_header_frame, text="Attachments", font=ctk.CTkFont(weight="bold"))
        self.lbl_attachments.pack(side="left")
        
        self.btn_add_attachment = ctk.CTkButton(attachment_header_frame, text="Attach File...", command=self._add_attachment)
        self.btn_add_attachment.pack(side="left", padx=10)

        self.lbl_timestamp = ctk.CTkLabel(attachment_header_frame, text="")
        self.lbl_timestamp.pack(side="left", padx=20)
        
        self.attachments_frame = ctk.CTkScrollableFrame(self.main_frame, label_text="")
        self.attachments_frame.grid(row=4, column=0, sticky="nsew", padx=10, pady=(0, 10))

        if note_id:
            self.load_note_content(note_id)
        else:
            self.current_note_id = None
            self.entry_note_title.delete(0, "end")
            self.txt_note_content.delete("1.0", "end")
            self.lbl_timestamp.configure(text="")
            self.btn_add_attachment.configure(state="disabled")  
            self.entry_note_title.focus()
            
    # --- Table Management Methods ---
    def create_new_table(self):
        if not self.db_path: messagebox.showerror("Error", "No project loaded."); return
        name = CustomDialog(self, title="New Table", prompt="Enter a name for the new table:").result
        if not name: return
        headers_str = CustomDialog(self, title="Table Headers", prompt="Enter column headers, separated by commas:").result
        if not headers_str: return
        headers = [h.strip() for h in headers_str.split(',')]; data = json.dumps([headers])
        conn = sqlite3.connect(self.db_path); cursor = conn.cursor(); cursor.execute("INSERT INTO custom_tables (name, data) VALUES (?, ?)", (name, data)); conn.commit(); conn.close()
        self.load_tables_list()

    def load_tables_list(self):
        for widget in self.tables_frame.winfo_children(): widget.destroy()
        if not self.db_path: return
        self.tables_frame.grid_columnconfigure(0, weight=1)
        conn = sqlite3.connect(self.db_path); cursor = conn.cursor(); cursor.execute("SELECT id, name FROM custom_tables ORDER BY name"); tables = cursor.fetchall(); conn.close()
        for i, (table_id, name) in enumerate(tables):
            btn = ctk.CTkButton(self.tables_frame, text=name, command=lambda t_id=table_id: self.switch_to_table_view(t_id)); btn.grid(row=i, column=0, sticky="ew", padx=(5,2), pady=2)
            del_btn = ctk.CTkButton(self.tables_frame, text="X", width=30, fg_color="red", hover_color="#C00000", command=lambda t_id=table_id: self.delete_table(t_id)); del_btn.grid(row=i, column=1, padx=(0,5), pady=2)

    def delete_table(self, table_id):
        if messagebox.askyesno("Confirm Delete", "Are you sure you want to permanently delete this table and all its data?"):
            conn = sqlite3.connect(self.db_path); cursor = conn.cursor(); cursor.execute("DELETE FROM custom_tables WHERE id = ?", (table_id,)); conn.commit(); conn.close()
            if self.current_table_id == table_id: self.switch_to_notes_view()
            self.load_tables_list()

    def switch_to_table_view(self, table_id):
        if self.current_table_id == table_id: return
        self.current_note_id = None; self.current_table_id = table_id; self.clear_main_frame()
        controls_frame = ctk.CTkFrame(self.main_frame); controls_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=10)
        self.btn_save_table = ctk.CTkButton(controls_frame, text="Save Table", command=self.save_table_data); self.btn_save_table.pack(side="left", padx=5)
        self.btn_add_row = ctk.CTkButton(controls_frame, text="Add Row", command=self.add_table_row); self.btn_add_row.pack(side="left", padx=5)
        self.btn_add_col = ctk.CTkButton(controls_frame, text="Add Column", command=self.add_table_column); self.btn_add_col.pack(side="left", padx=5)
        self.btn_import_csv = ctk.CTkButton(controls_frame, text="Import from CSV", command=self.import_csv); self.btn_import_csv.pack(side="left", padx=5)
        self.btn_export_csv = ctk.CTkButton(controls_frame, text="Export to CSV", command=self.export_csv); self.btn_export_csv.pack(side="left", padx=5)
        self.main_frame.grid_rowconfigure(1, weight=1)
        self.table_canvas = ctk.CTkScrollableFrame(self.main_frame); self.table_canvas.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0,10))
        self.load_table_data(table_id)

    def get_table_data_from_ui(self):
        if not self.table_widgets: return []
        ui_data = []; header_frames = self.table_widgets[0]
        headers = [frame.winfo_children()[0].cget("text") for frame in header_frames]; ui_data.append(headers)
        for row_widgets in self.table_widgets[1:]: ui_data.append([entry.get() for entry in row_widgets[:-1]])
        return ui_data

    def save_table_data(self, show_feedback=True):
        if not self.current_table_id: return
        ui_data = self.get_table_data_from_ui()
        json_data = json.dumps(ui_data)
        conn = sqlite3.connect(self.db_path); cursor = conn.cursor(); cursor.execute("UPDATE custom_tables SET data = ? WHERE id = ?", (json_data, self.current_table_id)); conn.commit(); conn.close()
        if show_feedback and hasattr(self, 'btn_save_table') and self.btn_save_table.winfo_exists():
            self.original_save_button_color = self.btn_save_table.cget("fg_color")
            self.btn_save_table.configure(text="✔ Saved", fg_color="green")
            self.after(2000, lambda: self.btn_save_table.configure(text="Save Table", fg_color=self.original_save_button_color))

    def load_table_data(self, table_id):
        for widget in self.table_canvas.winfo_children(): widget.destroy()
        self.table_widgets = []; conn = sqlite3.connect(self.db_path); cursor = conn.cursor(); cursor.execute("SELECT data FROM custom_tables WHERE id = ?", (table_id,)); result = cursor.fetchone(); conn.close()
        if not result: return
        data = json.loads(result[0])
        for r, row_data in enumerate(data):
            row_widgets = []
            for c, cell_data in enumerate(row_data):
                if r == 0:
                    header_frame = ctk.CTkFrame(self.table_canvas); header_frame.grid(row=r, column=c, padx=1, pady=1, sticky="nsew")
                    label = ctk.CTkLabel(header_frame, text=cell_data, font=ctk.CTkFont(weight="bold")); label.pack(side="left", expand=True, fill="x", padx=5)
                    del_btn = ctk.CTkButton(header_frame, text="X", width=25, fg_color="red", hover_color="#C00000", command=lambda col=c: self.delete_table_column(col)); del_btn.pack(side="right", padx=(0,2))
                    widget = header_frame
                else:
                    widget = ctk.CTkEntry(self.table_canvas); widget.insert(0, cell_data); widget.grid(row=r, column=c, padx=1, pady=1, sticky="ew")
                    widget.bind("<KeyRelease>", self._schedule_table_save)
                row_widgets.append(widget)
            if r > 0:
                del_row_btn = ctk.CTkButton(self.table_canvas, text="X", width=25, fg_color="red", hover_color="#C00000", command=lambda row=r: self.delete_table_row(row)); del_row_btn.grid(row=r, column=len(row_data), padx=5, pady=1)
                row_widgets.append(del_row_btn)
            self.table_widgets.append(row_widgets)

    def add_table_row(self):
        if not self.table_widgets: return
        num_cols = len(self.table_widgets[0]); new_row_num = len(self.table_widgets); row_widgets = []
        for c in range(num_cols):  
            entry = ctk.CTkEntry(self.table_canvas)
            entry.grid(row=new_row_num, column=c, padx=1, pady=1, sticky="ew")
            entry.bind("<KeyRelease>", self._schedule_table_save)
            row_widgets.append(entry)
        del_row_btn = ctk.CTkButton(self.table_canvas, text="X", width=25, fg_color="red", hover_color="#C00000", command=lambda row=new_row_num: self.delete_table_row(row)); del_row_btn.grid(row=new_row_num, column=num_cols, padx=5, pady=1)
        row_widgets.append(del_row_btn); self.table_widgets.append(row_widgets)

    def add_table_column(self):
        if not self.current_table_id: return
        header_name = CustomDialog(self, title="New Column", prompt="Enter the header name for the new column:").result
        if not header_name: return
        self.save_table_data(show_feedback=False); conn = sqlite3.connect(self.db_path); cursor = conn.cursor(); cursor.execute("SELECT data FROM custom_tables WHERE id = ?", (self.current_table_id,)); data = json.loads(cursor.fetchone()[0])
        data[0].append(header_name)
        for i in range(1, len(data)): data[i].append("")
        json_data = json.dumps(data); cursor.execute("UPDATE custom_tables SET data = ? WHERE id = ?", (json_data, self.current_table_id)); conn.commit(); conn.close()
        self.load_table_data(self.current_table_id)

    def delete_table_row(self, row_index):
        if messagebox.askyesno("Confirm Delete", f"Are you sure you want to delete row {row_index}?"):
            self.save_table_data(show_feedback=False); conn = sqlite3.connect(self.db_path); cursor = conn.cursor(); cursor.execute("SELECT data FROM custom_tables WHERE id = ?", (self.current_table_id,)); data = json.loads(cursor.fetchone()[0])
            data.pop(row_index)
            json_data = json.dumps(data); cursor.execute("UPDATE custom_tables SET data = ? WHERE id = ?", (json_data, self.current_table_id)); conn.commit(); conn.close()
            self.load_table_data(self.current_table_id)

    def delete_table_column(self, col_index):
        if messagebox.askyesno("Confirm Delete", f"Are you sure you want to delete column {col_index + 1}?"):
            self.save_table_data(show_feedback=False); conn = sqlite3.connect(self.db_path); cursor = conn.cursor()
            cursor.execute("SELECT data FROM custom_tables WHERE id = ?", (self.current_table_id,)); data = json.loads(cursor.fetchone()[0])
            for row in data: row.pop(col_index)
            json_data = json.dumps(data); cursor.execute("UPDATE custom_tables SET data = ? WHERE id = ?", (json_data, self.current_table_id)); conn.commit(); conn.close()
            self.load_table_data(self.current_table_id)

    def export_csv(self):
        if not self.current_table_id: return
        path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
        if not path: return
        self.save_table_data(show_feedback=False)
        conn = sqlite3.connect(self.db_path); cursor = conn.cursor(); cursor.execute("SELECT data FROM custom_tables WHERE id = ?", (self.current_table_id,)); result = cursor.fetchone(); conn.close()
        if not result: return
        data = json.loads(result[0])
        with open(path, 'w', newline='', encoding='utf-8') as f: writer = csv.writer(f); writer.writerows(data)
        messagebox.showinfo("Success", "Table exported to CSV.")

    def import_csv(self):
        if not self.current_table_id: return
        if not messagebox.askyesno("Confirm Import", "This will overwrite all data in the current table. Are you sure?"): return
        path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        if not path: return
        new_data = []
        with open(path, 'r', encoding='utf-8') as f:
            try: reader = csv.reader(f); new_data.extend(iter(reader))
            except Exception as e: messagebox.showerror("Error", f"Could not read CSV file. Error: {e}"); return
        if not new_data: messagebox.showerror("Error", "CSV file is empty or invalid."); return
        json_data = json.dumps(new_data)
        conn = sqlite3.connect(self.db_path); cursor = conn.cursor(); cursor.execute("UPDATE custom_tables SET data = ? WHERE id = ?", (json_data, self.current_table_id)); conn.commit(); conn.close()
        self.load_table_data(self.current_table_id)
        messagebox.showinfo("Success", "Data imported from CSV.")

    # --- To-Do View ---
    def switch_to_todo_view(self):
        self.current_note_id = None
        self.current_table_id = None
        self.clear_main_frame()

        self.main_frame.grid_columnconfigure(0, weight=1)
        self.main_frame.grid_rowconfigure(0, weight=0)
        self.main_frame.grid_rowconfigure(1, weight=3)
        self.main_frame.grid_rowconfigure(2, weight=2)

        input_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        input_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=(10, 5))
        input_frame.grid_columnconfigure(0, weight=1)

        self.main_btn_add_task = ctk.CTkButton(
            input_frame, 
            text="+ New Task...", 
            font=ctk.CTkFont(weight="bold"), 
            command=self.open_add_task_dialog, 
            fg_color="#2E8B57", 
            hover_color="#20603C",
            height=32
        )
        self.main_btn_add_task.grid(row=0, column=0, sticky="w")

        active_container = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        active_container.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 5))
        active_container.grid_columnconfigure(0, weight=1)
        active_container.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(active_container, text="Active Tasks", font=ctk.CTkFont(size=16, weight="bold")).grid(row=0, column=0, sticky="w", padx=5, pady=(0, 5))
        self.main_todo_frame = ctk.CTkScrollableFrame(active_container)
        self.main_todo_frame.grid(row=1, column=0, sticky="nsew")

        completed_container = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        completed_container.grid(row=2, column=0, sticky="nsew", padx=10, pady=(5, 10))
        completed_container.grid_columnconfigure(0, weight=1)
        completed_container.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(completed_container, text="Completed Tasks (Past Activities)", font=ctk.CTkFont(size=16, weight="bold")).grid(row=0, column=0, sticky="w", padx=5, pady=(0, 5))
        self.completed_todo_frame = ctk.CTkScrollableFrame(completed_container)
        self.completed_todo_frame.grid(row=1, column=0, sticky="nsew")

        self.load_tasks()

    def open_add_task_dialog(self):
        if not self.db_path:
            messagebox.showerror("Error", "No project loaded.")
            return

        dialog = TaskCreateDialog(self)
        task_data = dialog.result
        if not task_data:
            return

        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute(
            "INSERT INTO tasks (task_description, start_date, due_date) VALUES (?, ?, ?)",
            (task_data["description"], task_data["start_date"], task_data["due_date"])
        )
        conn.commit()
        conn.close()

        self.load_tasks()
        self.load_upcoming_tasks()

    def _format_date_badge(self, start_date: str, due_date: str):
        today_str = datetime.now().strftime("%Y-%m-%d")
        if start_date and due_date:
            try:
                s_dt = datetime.strptime(start_date, "%Y-%m-%d").strftime("%d/%m/%Y")
                d_dt = datetime.strptime(due_date, "%Y-%m-%d").strftime("%d/%m/%Y")
                is_overdue = due_date < today_str
                return f"🗓 {s_dt} → {d_dt}", is_overdue
            except ValueError:
                return f"🗓 {start_date} → {due_date}", False
        elif due_date:
            try:
                d_dt = datetime.strptime(due_date, "%Y-%m-%d").strftime("%d/%m/%Y")
                is_overdue = due_date < today_str
                return f"📅 Due: {d_dt}", is_overdue
            except ValueError:
                return f"📅 Due: {due_date}", False
        return "", False

    def load_tasks(self):
        if hasattr(self, 'main_todo_frame') and self.main_todo_frame.winfo_exists():
            for widget in self.main_todo_frame.winfo_children():
                widget.destroy()

        if hasattr(self, 'completed_todo_frame') and self.completed_todo_frame.winfo_exists():
            for widget in self.completed_todo_frame.winfo_children():
                widget.destroy()

        if not self.db_path:
            return

        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()

        cursor.execute("SELECT id, task_description, start_date, due_date FROM tasks WHERE is_completed = 0 ORDER BY id ASC")
        active_tasks = cursor.fetchall()

        cursor.execute("SELECT id, task_description, completed_date, completion_notes, start_date, due_date FROM tasks WHERE is_completed = 1 ORDER BY id DESC")
        completed_tasks = cursor.fetchall()
        conn.close()

        if hasattr(self, 'main_todo_frame') and self.main_todo_frame.winfo_exists():
            self.main_todo_frame.grid_columnconfigure(1, weight=1)
            if not active_tasks:
                ctk.CTkLabel(self.main_todo_frame, text="No active tasks. You're all caught up!").grid(row=0, column=0, padx=10, pady=10, sticky="w")
            else:
                for i, (task_id, desc, s_date, d_date) in enumerate(active_tasks):
                    card = ctk.CTkFrame(self.main_todo_frame, fg_color="#24292E")
                    card.grid(row=i, column=0, columnspan=3, sticky="ew", padx=5, pady=4)
                    card.grid_columnconfigure(1, weight=1)

                    check = ctk.CTkCheckBox(card, text="", width=24, command=lambda t_id=task_id: self.toggle_task(t_id))
                    check.grid(row=0, column=0, padx=(8, 4), pady=8, sticky="n")

                    text_frame = ctk.CTkFrame(card, fg_color="transparent")
                    text_frame.grid(row=0, column=1, sticky="ew", padx=4, pady=6)
                    text_frame.grid_columnconfigure(0, weight=1)

                    label = ctk.CTkLabel(text_frame, text=desc, wraplength=650, justify="left", anchor="w", font=ctk.CTkFont(size=13))
                    label.grid(row=0, column=0, sticky="w")

                    badge_text, is_overdue = self._format_date_badge(s_date, d_date)
                    if badge_text:
                        badge_color = "#F87171" if is_overdue else "#38BDF8"
                        overdue_tag = " [OVERDUE]" if is_overdue else ""
                        meta_lbl = ctk.CTkLabel(text_frame, text=f"{badge_text}{overdue_tag}", text_color=badge_color, font=ctk.CTkFont(size=11, weight="bold"))
                        meta_lbl.grid(row=1, column=0, sticky="w", pady=(2, 0))

                    del_btn = ctk.CTkButton(card, text="X", width=28, height=26, fg_color="red", hover_color="#C00000", command=lambda t_id=task_id: self.delete_task(t_id))
                    del_btn.grid(row=0, column=2, padx=(5, 8), pady=8, sticky="n")

        if hasattr(self, 'completed_todo_frame') and self.completed_todo_frame.winfo_exists():
            self.completed_todo_frame.grid_columnconfigure(0, weight=1)
            if not completed_tasks:
                ctk.CTkLabel(self.completed_todo_frame, text="No completed tasks recorded yet.").grid(row=0, column=0, padx=10, pady=10, sticky="w")
            else:
                for i, (task_id, desc, completed_date, completion_notes, s_date, d_date) in enumerate(completed_tasks):
                    card = ctk.CTkFrame(self.completed_todo_frame, fg_color="#2B2B2B")
                    card.grid(row=i, column=0, sticky="ew", padx=5, pady=4)
                    card.grid_columnconfigure(0, weight=1)

                    comp_date_disp = "N/A"
                    if completed_date:
                        try:
                            dt = datetime.strptime(completed_date, "%Y-%m-%d %H:%M:%S")
                            comp_date_disp = dt.strftime("%d/%m/%Y %H:%M")
                        except ValueError:
                            comp_date_disp = completed_date

                    header_frame = ctk.CTkFrame(card, fg_color="transparent")
                    header_frame.grid(row=0, column=0, sticky="ew", padx=8, pady=(6, 2))
                    header_frame.grid_columnconfigure(0, weight=1)

                    desc_lbl = ctk.CTkLabel(
                        header_frame, 
                        text=f"✔ {desc}", 
                        font=ctk.CTkFont(weight="bold"), 
                        text_color="#94A3B8", 
                        anchor="w", 
                        justify="left",
                        wraplength=600
                    )
                    desc_lbl.grid(row=0, column=0, sticky="w")

                    meta_lbl = ctk.CTkLabel(header_frame, text=f"Completed: {comp_date_disp}", text_color="#38BDF8", font=ctk.CTkFont(size=11))
                    meta_lbl.grid(row=0, column=1, padx=(10, 5), sticky="e")

                    del_btn = ctk.CTkButton(header_frame, text="X", width=25, height=22, fg_color="red", hover_color="#C00000", command=lambda t_id=task_id: self.delete_task(t_id))
                    del_btn.grid(row=0, column=2, padx=(5, 0), sticky="e")

                    if completion_notes and completion_notes.strip():
                        notes_lbl = ctk.CTkLabel(
                            card, 
                            text=f"Note: {completion_notes}", 
                            text_color="#CBD5E1", 
                            font=ctk.CTkFont(size=12, slant="italic"), 
                            anchor="w", 
                            justify="left",
                            wraplength=750
                        )
                        notes_lbl.grid(row=1, column=0, sticky="ew", padx=12, pady=(0, 6))

    # --- Roadmap / Gantt Timeline View ---
    def switch_to_roadmap_view(self):
        self.current_note_id = None
        self.current_table_id = None
        self.clear_main_frame()

        self.main_frame.grid_columnconfigure(0, weight=1)
        self.main_frame.grid_rowconfigure(0, weight=0)
        self.main_frame.grid_rowconfigure(1, weight=1)

        top_controls = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        top_controls.grid(row=0, column=0, sticky="ew", padx=10, pady=(10, 5))
        top_controls.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(top_controls, text="Project Roadmap & Timelines", font=ctk.CTkFont(size=18, weight="bold")).grid(row=0, column=0, sticky="w")

        self.roadmap_scope_var = ctk.StringVar(value="Current Project")
        scope_seg = ctk.CTkSegmentedButton(
            top_controls,
            values=["Current Project", "All Projects (Global Conflict View)"],
            command=lambda val: self.render_roadmap(),
            variable=self.roadmap_scope_var
        )
        scope_seg.grid(row=0, column=1, sticky="e")

        canvas_container = ctk.CTkFrame(self.main_frame)
        canvas_container.grid(row=1, column=0, sticky="nsew", padx=10, pady=(5, 10))
        canvas_container.grid_columnconfigure(0, weight=1)
        canvas_container.grid_rowconfigure(0, weight=1)

        self.roadmap_canvas = tk.Canvas(canvas_container, bg="#1E1E1E", highlightthickness=0)
        self.roadmap_canvas.grid(row=0, column=0, sticky="nsew")

        v_scroll = ctk.CTkScrollbar(canvas_container, orientation="vertical", command=self.roadmap_canvas.yview)
        v_scroll.grid(row=0, column=1, sticky="ns")
        h_scroll = ctk.CTkScrollbar(canvas_container, orientation="horizontal", command=self.roadmap_canvas.xview)
        h_scroll.grid(row=1, column=0, sticky="ew")

        self.roadmap_canvas.configure(xscrollcommand=h_scroll.set, yscrollcommand=v_scroll.set)

        self.render_roadmap()

    def _gather_roadmap_items(self, global_mode=False):
        items = []

        target_files = []
        if global_mode:
            target_files = [os.path.join(self.projects_dir, f) for f in os.listdir(self.projects_dir) if f.endswith(".db")]
        elif self.db_path:
            target_files = [self.db_path]

        for db_file in target_files:
            p_name = os.path.splitext(os.path.basename(db_file))[0]
            try:
                conn = sqlite3.connect(db_file)
                cursor = conn.cursor()

                cursor.execute("SELECT task_description, start_date, due_date FROM tasks WHERE is_completed = 0 AND (start_date IS NOT NULL OR due_date IS NOT NULL)")
                for desc, s_date, d_date in cursor.fetchall():
                    try:
                        start_dt = datetime.strptime(s_date, "%Y-%m-%d").date() if s_date else None
                    except (ValueError, TypeError):
                        start_dt = None
                    try:
                        due_dt = datetime.strptime(d_date, "%Y-%m-%d").date() if d_date else None
                    except (ValueError, TypeError):
                        due_dt = None

                    if start_dt and not due_dt:
                        due_dt = start_dt
                    elif due_dt and not start_dt:
                        start_dt = due_dt

                    if start_dt and due_dt:
                        items.append({
                            "project": p_name,
                            "type": "task",
                            "title": desc,
                            "start": start_dt,
                            "end": due_dt
                        })

                cursor.execute("SELECT event_date, description FROM events")
                for e_date, desc in cursor.fetchall():
                    try:
                        e_dt = datetime.strptime(e_date, "%Y-%m-%d").date()
                        items.append({
                            "project": p_name,
                            "type": "event",
                            "title": desc,
                            "start": e_dt,
                            "end": e_dt
                        })
                    except (ValueError, TypeError):
                        pass

                conn.close()
            except Exception:
                continue

        return sorted(items, key=lambda x: (x["start"], x["end"]))

    def render_roadmap(self):
        self.roadmap_canvas.delete("all")
        scope = self.roadmap_scope_var.get()
        is_global = "All Projects" in scope

        items = self._gather_roadmap_items(global_mode=is_global)

        if not items:
            self.roadmap_canvas.create_text(
                350, 150, 
                text="No scheduled tasks or events found with dates.\nAdd start/due dates to tasks or calendar events to populate this roadmap.",
                fill="#94A3B8", font=("Segoe UI", 12), justify="center"
            )
            return

        today = datetime.now().date()
        min_date = min([it["start"] for it in items] + [today]) - timedelta(days=2)
        max_date = max([it["end"] for it in items] + [today]) + timedelta(days=14)

        total_days = (max_date - min_date).days + 1
        day_width = 38
        row_height = 36
        header_height = 50
        label_column_width = 340

        canvas_width = label_column_width + (total_days * day_width) + 50
        canvas_height = header_height + (len(items) * row_height) + 50

        date_project_map = {}
        if is_global:
            for it in items:
                curr = it["start"]
                while curr <= it["end"]:
                    date_project_map.setdefault(curr, set()).add(it["project"])
                    curr += timedelta(days=1)

        for d in range(total_days):
            current_day = min_date + timedelta(days=d)
            x = label_column_width + (d * day_width)

            is_weekend = current_day.weekday() >= 5
            col_bg = "#26292D" if is_weekend else "#1E1E1E"

            self.roadmap_canvas.create_rectangle(x, header_height, x + day_width, canvas_height, fill=col_bg, outline="#2D3136")

            day_num = current_day.strftime("%d")
            day_name = current_day.strftime("%a")[:2]
            month_name = current_day.strftime("%b")

            self.roadmap_canvas.create_text(x + (day_width / 2), 15, text=month_name if current_day.day == 1 or d == 0 else "", fill="#94A3B8", font=("Segoe UI", 9, "bold"))
            self.roadmap_canvas.create_text(x + (day_width / 2), 30, text=day_num, fill="#FFFFFF" if not is_weekend else "#64748B", font=("Segoe UI", 10, "bold"))
            self.roadmap_canvas.create_text(x + (day_width / 2), 43, text=day_name, fill="#38BDF8" if not is_weekend else "#475569", font=("Segoe UI", 8))

            if current_day == today:
                today_x = x + (day_width / 2)
                self.roadmap_canvas.create_line(today_x, header_height - 10, today_x, canvas_height, fill="#EF4444", width=2, dash=(4, 2))
                self.roadmap_canvas.create_text(today_x, 8, text="TODAY", fill="#EF4444", font=("Segoe UI", 8, "bold"))

        self.roadmap_canvas.create_rectangle(0, 0, label_column_width, canvas_height, fill="#24292E", outline="#2D3136")
        self.roadmap_canvas.create_text(15, 25, text="Task / Milestone", fill="#FFFFFF", font=("Segoe UI", 11, "bold"), anchor="w")

        title_font = tkfont.Font(family="Segoe UI", size=10)
        max_allowed_text_width = label_column_width - 25

        for i, item in enumerate(items):
            y = header_height + (i * row_height)

            has_conflict = False
            if is_global:
                curr = item["start"]
                while curr <= item["end"]:
                    if len(date_project_map.get(curr, set())) > 1:
                        has_conflict = True
                        break
                    curr += timedelta(days=1)

            row_bg = "#2D1D1D" if has_conflict else ("#282C34" if i % 2 == 0 else "#21252B")
            self.roadmap_canvas.create_rectangle(0, y, label_column_width, y + row_height, fill=row_bg, outline="#2D3136")

            conflict_tag = "⚠ " if has_conflict else ""
            prefix = f"[{item['project']}] " if is_global else ""
            icon = "🔷 " if item["type"] == "event" else "◼ "
            display_title = f"{conflict_tag}{prefix}{icon}{item['title']}"

            if title_font.measure(display_title) > max_allowed_text_width:
                while display_title and title_font.measure(display_title + "...") > max_allowed_text_width:
                    display_title = display_title[:-1]
                display_title = display_title.rstrip() + "..."

            title_color = "#FCA5A5" if has_conflict else "#E2E8F0"
            self.roadmap_canvas.create_text(10, y + (row_height / 2), text=display_title, fill=title_color, font=title_font, anchor="w")

            start_offset = (item["start"] - min_date).days
            end_offset = (item["end"] - min_date).days

            bar_x1 = label_column_width + (start_offset * day_width) + 4
            bar_x2 = label_column_width + ((end_offset + 1) * day_width) - 4
            bar_y1 = y + 8
            bar_y2 = y + row_height - 8

            if item["start"] == item["end"]:
                mid_x = (bar_x1 + bar_x2) / 2
                mid_y = y + (row_height / 2)
                d_color = "#A855F7" if item["type"] == "event" else "#38BDF8"
                points = [mid_x, mid_y - 8, mid_x + 8, mid_y, mid_x, mid_y + 8, mid_x - 8, mid_y]
                self.roadmap_canvas.create_polygon(points, fill=d_color, outline="#FFFFFF")
            else:
                bar_color = "#F59E0B" if has_conflict else "#2563EB"
                self.roadmap_canvas.create_rectangle(bar_x1, bar_y1, bar_x2, bar_y2, fill=bar_color, outline="#FFFFFF" if has_conflict else "#1D4ED8", width=1)
                
                duration_days = (item["end"] - item["start"]).days + 1
                if (bar_x2 - bar_x1) > 60:
                    self.roadmap_canvas.create_text(bar_x1 + 8, y + (row_height / 2), text=f"{duration_days}d", fill="#FFFFFF", font=("Segoe UI", 9, "bold"), anchor="w")

        self.roadmap_canvas.config(scrollregion=(0, 0, canvas_width, canvas_height))

    # --- Questions View ---
    def switch_to_questions_view(self):
        self.current_note_id = None
        self.current_table_id = None
        self.clear_main_frame()

        self.main_frame.grid_columnconfigure(0, weight=1)
        self.main_frame.grid_rowconfigure(0, weight=0)
        self.main_frame.grid_rowconfigure(1, weight=3)
        self.main_frame.grid_rowconfigure(2, weight=2)

        input_frame = ctk.CTkFrame(self.main_frame)
        input_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=10)
        input_frame.grid_columnconfigure(0, weight=1)

        self.main_entry_new_question = ctk.CTkEntry(input_frame, placeholder_text="Enter an outstanding question / blocker...")
        self.main_entry_new_question.grid(row=0, column=0, sticky="ew", padx=(0, 10))
        self.main_entry_new_question.bind("<Return>", lambda event: self.add_question())

        self.main_btn_add_question = ctk.CTkButton(input_frame, text="Add Question", command=self.add_question, fg_color="#2E8B57", hover_color="#20603C")
        self.main_btn_add_question.grid(row=0, column=1)

        unanswered_container = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        unanswered_container.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 5))
        unanswered_container.grid_columnconfigure(0, weight=1)
        unanswered_container.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(unanswered_container, text="Outstanding / Unanswered Questions", font=ctk.CTkFont(size=16, weight="bold")).grid(row=0, column=0, sticky="w", padx=5, pady=(0, 5))
        self.unanswered_frame = ctk.CTkScrollableFrame(unanswered_container)
        self.unanswered_frame.grid(row=1, column=0, sticky="nsew")

        answered_container = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        answered_container.grid(row=2, column=0, sticky="nsew", padx=10, pady=(5, 10))
        answered_container.grid_columnconfigure(0, weight=1)
        answered_container.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(answered_container, text="Answered Questions", font=ctk.CTkFont(size=16, weight="bold")).grid(row=0, column=0, sticky="w", padx=5, pady=(0, 5))
        self.answered_frame = ctk.CTkScrollableFrame(answered_container)
        self.answered_frame.grid(row=1, column=0, sticky="nsew")

        self.load_questions()

    def load_questions(self):
        if hasattr(self, 'unanswered_frame') and self.unanswered_frame.winfo_exists():
            for widget in self.unanswered_frame.winfo_children():
                widget.destroy()

        if hasattr(self, 'answered_frame') and self.answered_frame.winfo_exists():
            for widget in self.answered_frame.winfo_children():
                widget.destroy()

        if not self.db_path:
            return

        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute("SELECT id, question_text FROM questions WHERE is_answered = 0 ORDER BY id ASC")
        unanswered = cursor.fetchall()

        cursor.execute("SELECT id, question_text, answer_text, answered_date FROM questions WHERE is_answered = 1 ORDER BY id DESC")
        answered = cursor.fetchall()
        conn.close()

        if hasattr(self, 'unanswered_frame') and self.unanswered_frame.winfo_exists():
            self.unanswered_frame.grid_columnconfigure(0, weight=1)
            if not unanswered:
                ctk.CTkLabel(self.unanswered_frame, text="No outstanding questions right now.").grid(row=0, column=0, padx=10, pady=10, sticky="w")
            else:
                for i, (q_id, q_text) in enumerate(unanswered):
                    card = ctk.CTkFrame(self.unanswered_frame, fg_color="#24292E")
                    card.grid(row=i, column=0, sticky="ew", padx=5, pady=4)
                    card.grid_columnconfigure(0, weight=1)

                    q_lbl = ctk.CTkLabel(card, text=f"?  {q_text}", font=ctk.CTkFont(weight="bold"), wraplength=650, anchor="w", justify="left")
                    q_lbl.grid(row=0, column=0, sticky="w", padx=10, pady=8)

                    btn_ans = ctk.CTkButton(card, text="Answer", width=70, height=26, fg_color="#2563EB", hover_color="#1D4ED8",
                                            command=lambda target_id=q_id: self.answer_question(target_id))
                    btn_ans.grid(row=0, column=1, padx=(5, 5), pady=8)

                    btn_del = ctk.CTkButton(card, text="X", width=26, height=26, fg_color="red", hover_color="#C00000",
                                            command=lambda target_id=q_id: self.delete_question(target_id))
                    btn_del.grid(row=0, column=2, padx=(0, 8), pady=8)

        if hasattr(self, 'answered_frame') and self.answered_frame.winfo_exists():
            self.answered_frame.grid_columnconfigure(0, weight=1)
            if not answered:
                ctk.CTkLabel(self.answered_frame, text="No answered questions yet.").grid(row=0, column=0, padx=10, pady=10, sticky="w")
            else:
                for i, (q_id, q_text, ans_text, ans_date) in enumerate(answered):
                    card = ctk.CTkFrame(self.answered_frame, fg_color="#2B2B2B")
                    card.grid(row=i, column=0, sticky="ew", padx=5, pady=4)
                    card.grid_columnconfigure(0, weight=1)

                    ans_date_disp = "N/A"
                    if ans_date:
                        try:
                            dt = datetime.strptime(ans_date, "%Y-%m-%d %H:%M:%S")
                            ans_date_disp = dt.strftime("%d/%m/%Y %H:%M")
                        except ValueError:
                            ans_date_disp = ans_date

                    top_bar = ctk.CTkFrame(card, fg_color="transparent")
                    top_bar.grid(row=0, column=0, sticky="ew", padx=10, pady=(6, 2))
                    top_bar.grid_columnconfigure(0, weight=1)

                    q_lbl = ctk.CTkLabel(top_bar, text=f"Q: {q_text}", font=ctk.CTkFont(weight="bold"), wraplength=600, anchor="w", justify="left", text_color="#E2E8F0")
                    q_lbl.grid(row=0, column=0, sticky="w")

                    date_lbl = ctk.CTkLabel(top_bar, text=f"Answered: {ans_date_disp}", font=ctk.CTkFont(size=11), text_color="#38BDF8")
                    date_lbl.grid(row=0, column=1, padx=5, sticky="e")

                    del_btn = ctk.CTkButton(top_bar, text="X", width=24, height=22, fg_color="red", hover_color="#C00000",
                                            command=lambda target_id=q_id: self.delete_question(target_id))
                    del_btn.grid(row=0, column=2, padx=(5, 0), sticky="e")

                    ans_lbl = ctk.CTkLabel(card, text=f"A: {ans_text}", wraplength=750, anchor="w", justify="left", text_color="#94A3B8", font=ctk.CTkFont(size=12))
                    ans_lbl.grid(row=1, column=0, sticky="ew", padx=15, pady=(0, 6))

    def add_question(self):
        q_text = self.main_entry_new_question.get().strip()
        if not q_text: return
        if not self.db_path:
            messagebox.showerror("Error", "No project loaded.")
            return

        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute("INSERT INTO questions (question_text, is_answered) VALUES (?, 0)", (q_text,))
        conn.commit()
        conn.close()

        self.main_entry_new_question.delete(0, 'end')
        self.load_questions()

    def answer_question(self, q_id):
        if not self.db_path: return
        dialog = CustomDialog(self, title="Record Answer", prompt="Enter the answer:")
        ans = dialog.result
        if ans is None:
            return

        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute("UPDATE questions SET is_answered = 1, answer_text = ?, answered_date = ? WHERE id = ?", (ans.strip(), now_str, q_id))
        conn.commit()
        conn.close()

        self.load_questions()

    def delete_question(self, q_id):
        if messagebox.askyesno("Confirm Delete", "Are you sure you want to delete this question?"):
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute("DELETE FROM questions WHERE id = ?", (q_id,))
            conn.commit()
            conn.close()
            self.load_questions()

    def toggle_task(self, task_id):
        if not self.db_path: return

        dialog = CustomDialog(self, title="Complete Task", prompt="Add an optional completion note:")
        note = dialog.result
        if note is None:
            note = ""

        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute(
            "UPDATE tasks SET is_completed = 1, completed_date = ?, completion_notes = ? WHERE id = ?",
            (now_str, note.strip(), task_id)
        )
        conn.commit()
        conn.close()
        self.load_tasks()
        self.load_upcoming_tasks()
    
    def delete_task(self, task_id):
        if messagebox.askyesno("Confirm Delete", "Are you sure you want to delete this task?"):
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute("DELETE FROM tasks WHERE id = ?", (task_id,))
            conn.commit()
            conn.close()
            self.load_tasks()
            self.load_upcoming_tasks()
                
    def load_upcoming_tasks(self):
        for widget in self.upcoming_tasks_frame.winfo_children(): widget.destroy()
        if not self.db_path: return
        self.upcoming_tasks_frame.grid_columnconfigure(0, weight=1)
        conn = sqlite3.connect(self.db_path); cursor = conn.cursor()
        cursor.execute("SELECT task_description, due_date FROM tasks WHERE is_completed = 0 ORDER BY id ASC LIMIT 10")
        tasks = cursor.fetchall(); conn.close()
        if not tasks:   
            ctk.CTkLabel(self.upcoming_tasks_frame, text="No upcoming tasks.").grid(row=0, column=0, padx=5, pady=5)
        else:
            for i, (desc, due_date) in enumerate(tasks):
                due_tag = ""
                if due_date:
                    try:
                        due_tag = f" ({datetime.strptime(due_date, '%Y-%m-%d').strftime('%d/%m')})"
                    except ValueError:
                        pass
                lbl_desc = ctk.CTkLabel(self.upcoming_tasks_frame, text=f"• {desc}{due_tag}", wraplength=240, justify="left", anchor="w")
                lbl_desc.grid(row=i, column=0, sticky="ew", padx=5, pady=2)
                
    def toggle_pin_note(self, note_id, current_pinned_status):
        if not self.db_path: return
        new_status = 0 if current_pinned_status == 1 else 1
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute("UPDATE notes SET is_pinned = ? WHERE id = ?", (new_status, note_id))
        conn.commit()
        conn.close()
        self.load_notes_list()

    def load_notes_list(self):
        for widget in self.notes_frame.winfo_children(): widget.destroy()
        if not self.db_path: return
        self.notes_frame.grid_columnconfigure(0, weight=1)
        
        selected_folder = self.note_folder_var.get()
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        if selected_folder in ["All Folders", "+ New Folder..."]:
            cursor.execute("SELECT id, title, is_pinned, folder FROM notes ORDER BY is_pinned DESC, last_modified_date DESC")
        else:
            cursor.execute("SELECT id, title, is_pinned, folder FROM notes WHERE folder = ? ORDER BY is_pinned DESC, last_modified_date DESC", (selected_folder,))
        
        notes = cursor.fetchall()
        conn.close()
        
        for i, (note_id, title, is_pinned, folder) in enumerate(notes):
            bg_color = "#1D4ED8" if is_pinned else "#3B8ED0"
            frame = ctk.CTkFrame(self.notes_frame, fg_color=bg_color, corner_radius=6)
            frame.grid(row=i, column=0, sticky="ew", padx=(5, 2), pady=2)
            frame.grid_columnconfigure(0, weight=1)
            
            display_title = f"📌 {title}" if is_pinned else title
            label = ctk.CTkLabel(frame, text=display_title, wraplength=140, justify="left", anchor="w", fg_color="transparent", text_color="white")
            label.grid(row=0, column=0, sticky="ew", padx=8, pady=5)
            
            click_handler = self._create_note_click_handler(note_id)
            frame.bind("<Button-1>", click_handler)
            label.bind("<Button-1>", click_handler)
            
            pin_btn = ctk.CTkButton(
                self.notes_frame, 
                text="📌", 
                width=26, 
                fg_color="#475569" if not is_pinned else "#F59E0B", 
                hover_color="#334155" if not is_pinned else "#D97706",
                command=lambda n_id=note_id, p_val=is_pinned: self.toggle_pin_note(n_id, p_val)
            )
            pin_btn.grid(row=i, column=1, padx=(2, 2), pady=2)

            del_btn = ctk.CTkButton(self.notes_frame, text="X", width=26, fg_color="red", hover_color="#C00000", command=lambda n_id=note_id: self.delete_note(n_id))
            del_btn.grid(row=i, column=2, padx=(0, 5), pady=2)

    # --- Links Management Methods ---
    def add_link(self):
        if not self.db_path:
            messagebox.showerror("Error", "No project loaded.")
            return

        name = self.entry_link_name.get().strip()
        url = self.entry_link_url.get().strip()

        if not name or not url:
            messagebox.showwarning("Validation Error", "Please provide both a link name and URL.")
            return

        if not (url.startswith("http://") or url.startswith("https://") or os.path.exists(url)):
            url = "https://" + url

        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute("INSERT INTO links (link_name, url) VALUES (?, ?)", (name, url))
        conn.commit()
        conn.close()

        self.entry_link_name.delete(0, "end")
        self.entry_link_url.delete(0, "end")
        self.load_links()

    def open_link(self, url):
        try:
            webbrowser.open(url)
        except Exception as e:
            messagebox.showerror("Error", f"Could not open link: {e}")

    def delete_link(self, link_id):
        if messagebox.askyesno("Confirm Delete", "Are you sure you want to delete this link?"):
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute("DELETE FROM links WHERE id = ?", (link_id,))
            conn.commit()
            conn.close()
            self.load_links()
            
    def load_links(self):
        for widget in self.links_frame.winfo_children():widget.destroy()
        if not self.db_path:return
        self.links_frame.grid_columnconfigure(0, weight=1)
        conn=sqlite3.connect(self.db_path);cursor=conn.cursor();cursor.execute("SELECT id, link_name, url FROM links ORDER BY id");links=cursor.fetchall();conn.close()
        for i, (link_id,name,url) in enumerate(links):
            frame = ctk.CTkFrame(self.links_frame, fg_color="#3B8ED0", corner_radius=6); frame.grid(row=i, column=0, sticky="ew", padx=(5,2), pady=2); frame.grid_columnconfigure(0, weight=1)
            label = ctk.CTkLabel(frame, text=name, wraplength=180, justify="left", anchor="w", fg_color="transparent", text_color="white"); label.grid(row=0, column=0, sticky="ew", padx=8, pady=5)
            click_handler = self._create_link_click_handler(url)
            frame.bind("<Button-1>", click_handler)
            label.bind("<Button-1>", click_handler)
            del_btn=ctk.CTkButton(self.links_frame,text="X",width=30,fg_color="red",hover_color="#C00000",command=lambda l_id=link_id:self.delete_link(l_id)); del_btn.grid(row=i, column=1, padx=(0,5), pady=2)
            
    def load_note_content(self, note_id):
        self.current_note_id = note_id
        conn = sqlite3.connect(self.db_path); cursor = conn.cursor()
        cursor.execute("SELECT title, content, last_modified_date, folder FROM notes WHERE id = ?", (note_id,))
        note = cursor.fetchone(); conn.close()
        if note:
            title, content, last_mod, folder = note
            self.entry_note_title.delete(0, "end"); self.entry_note_title.insert(0, title)
            self._deserialize_note_content(content)
            
            note_folder = folder if folder else "General"
            self.current_note_folder_var.set(note_folder)
            self._update_folder_dropdown_options()
            
            self.lbl_timestamp.configure(text=f"Last Modified: {last_mod}" if last_mod else "")
            self.btn_add_attachment.configure(state="normal")
            self._load_attachments_for_note(note_id)
            
    def save_note(self, event=None):
        if not self.db_path or not hasattr(self, 'entry_note_title') or not self.entry_note_title.winfo_exists():
            return
        title = self.entry_note_title.get()
        raw_text = self.txt_note_content.get("1.0", "end-1c")
        folder = self.current_note_folder_var.get()
        if not folder or folder == "+ New Folder...":
            folder = "General"
        
        if not title.strip() and not raw_text.strip(): return
        
        serialized_content = self._serialize_note_content()
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        conn = sqlite3.connect(self.db_path); cursor = conn.cursor()
        if self.current_note_id is None:
            cursor.execute("INSERT INTO notes (title, content, created_date, last_modified_date, is_pinned, folder) VALUES (?, ?, ?, ?, 0, ?)", 
                           (title, serialized_content, now_str, now_str, folder))
            self.current_note_id = cursor.lastrowid
            self.btn_add_attachment.configure(state="normal")
        else:
            cursor.execute("UPDATE notes SET title = ?, content = ?, last_modified_date = ?, folder = ? WHERE id = ?", 
                           (title, serialized_content, now_str, folder, self.current_note_id))
        conn.commit(); conn.close()
        self.lbl_timestamp.configure(text=f"Last Modified: {now_str}")
        self._update_folder_dropdown_options()
        self.load_notes_list()
        
    def delete_note(self, note_id):
        if messagebox.askyesno("Confirm Delete", "Are you sure you want to permanently delete this note and all its attachments?"):
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute("PRAGMA foreign_keys = ON")
            cursor.execute("DELETE FROM notes WHERE id = ?", (note_id,))
            conn.commit()
            conn.close()
            if self.current_note_id == note_id:
                self.switch_to_notes_view()
            self._update_folder_dropdown_options()
            self.load_notes_list()
            
    def _add_attachment(self):
        if self.current_note_id is None:
            messagebox.showerror("Error", "Please save the note before adding attachments.")
            return
        filepath = filedialog.askopenfilename()
        if not filepath:
            return
        filename = os.path.basename(filepath)
        try:
            with open(filepath, 'rb') as f:
                file_data = f.read()
        except Exception as e:
            messagebox.showerror("Error", f"Could not read file: {e}")
            return
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute("INSERT INTO attachments (note_id, filename, file_data) VALUES (?, ?, ?)",
                       (self.current_note_id, filename, file_data))
        conn.commit()
        conn.close()
        self._load_attachments_for_note(self.current_note_id)
        
    def _load_attachments_for_note(self, note_id):
        for widget in self.attachments_frame.winfo_children():
            widget.destroy()

        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute("SELECT id, filename FROM attachments WHERE note_id = ?", (note_id,))
        attachments = cursor.fetchall()
        conn.close()
        
        self.attachments_frame.grid_columnconfigure(0, weight=1)
        for i, (att_id, filename) in enumerate(attachments):
            item_frame = ctk.CTkFrame(self.attachments_frame)
            item_frame.grid(row=i, column=0, sticky="ew", padx=5, pady=2)
            item_frame.grid_columnconfigure(0, weight=1)
            
            lbl = ctk.CTkLabel(item_frame, text=filename, anchor="w")
            lbl.grid(row=0, column=0, sticky="ew", padx=5, pady=5)
            
            save_btn = ctk.CTkButton(item_frame, text="Save As...", width=80,
                                     command=lambda a_id=att_id, f_name=filename: self._save_attachment_as(a_id, f_name))
            save_btn.grid(row=0, column=1, padx=5, pady=5)
            
            del_btn = ctk.CTkButton(item_frame, text="X", width=30, fg_color="red", hover_color="#C00000",
                                     command=lambda a_id=att_id: self._delete_attachment(a_id))
            del_btn.grid(row=0, column=2, padx=5, pady=5)
            
            open_handler = self._create_attachment_open_handler(att_id, filename)
            item_frame.bind("<Double-Button-1>", open_handler)
            lbl.bind("<Double-Button-1>", open_handler)
            
    def _save_attachment_as(self, attachment_id, filename):
        save_path = filedialog.asksaveasfilename(initialfile=filename)
        if not save_path:
            return
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute("SELECT file_data FROM attachments WHERE id = ?", (attachment_id,))
        result = cursor.fetchone()
        conn.close()
        if result:
            try:
                with open(save_path, 'wb') as f:
                    f.write(result[0])
                messagebox.showinfo("Success", f"Saved {filename}")
            except Exception as e:
                messagebox.showerror("Error", f"Could not save file: {e}")
                
    def _delete_attachment(self, attachment_id):
        if messagebox.askyesno("Confirm", "Are you sure you want to delete this attachment?"):
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute("DELETE FROM attachments WHERE id = ?", (attachment_id,))
            conn.commit()
            conn.close()
            self._load_attachments_for_note(self.current_note_id)
            
    def _open_attachment(self, attachment_id, filename):
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute("SELECT file_data FROM attachments WHERE id = ?", (attachment_id,))
        result = cursor.fetchone()
        conn.close()

        if not result:
            messagebox.showerror("Error", "Could not find attachment data.")
            return

        file_data = result[0]
        
        try:
            temp_dir = tempfile.gettempdir()
            temp_path = os.path.join(temp_dir, filename)

            with open(temp_path, "wb") as temp_file:
                temp_file.write(file_data)

            os.startfile(temp_path)
        except Exception as e:
            messagebox.showerror("Error", f"Could not open file: {e}")
            
    def switch_to_calendar_view(self):
        self.current_note_id = None
        self.current_table_id = None
        self.clear_main_frame()
        
        self.main_frame.grid_rowconfigure(0, weight=0)
        self.main_frame.grid_rowconfigure(1, weight=1)
        self.main_frame.grid_columnconfigure(0, weight=1)
        
        self.cal = Calendar(
            self.main_frame,
            selectmode='day',
            font="Arial 11 bold",
            background="#1F1F1F",
            foreground="#FFFFFF",
            headersbackground="#2B2B2B",
            headersforeground="#38BDF8",
            normalbackground="#2A2D2E",
            normalforeground="#F1F5F9",
            weekendbackground="#222527",
            weekendforeground="#A5B4FC",
            othermonthbackground="#18181B",
            othermonthforeground="#64748B",
            othermonthwebackground="#141416",
            othermonthweforeground="#475569",
            selectbackground="#2563EB",
            selectforeground="#FFFFFF",
            bordercolor="#3F3F46"
        )
        self.cal.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        self.cal.bind("<<CalendarSelected>>", self.on_date_select)
        
        event_frame = ctk.CTkFrame(self.main_frame)
        event_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=10)
        event_frame.grid_columnconfigure(0, weight=1)
        
        self.lbl_event_day = ctk.CTkLabel(event_frame, text="Events for:", font=ctk.CTkFont(size=14, weight="bold"))
        self.lbl_event_day.pack(pady=(5, 0))
        
        self.event_list_frame = ctk.CTkScrollableFrame(event_frame)
        self.event_list_frame.pack(fill="both", expand=True, padx=5, pady=5)
        
        btn_add_event = ctk.CTkButton(event_frame, text="Add Event to Selected Date", command=self.add_event)
        btn_add_event.pack(pady=10)
        
        self.on_date_select()
        self.load_all_event_markers()

    def on_date_select(self, event=None):
        if not self.db_path: self.lbl_event_day.configure(text="No project loaded"); return
        selected_date = self.cal.get_date(); db_date = datetime.strptime(selected_date, "%m/%d/%y").strftime("%Y-%m-%d")
        self.lbl_event_day.configure(text=f"Events for: {db_date}")
        for widget in self.event_list_frame.winfo_children(): widget.destroy()
        conn = sqlite3.connect(self.db_path); cursor = conn.cursor(); cursor.execute("SELECT id, description FROM events WHERE event_date = ? ORDER BY id", (db_date,)); events = cursor.fetchall(); conn.close()
        if not events: ctk.CTkLabel(self.event_list_frame, text="No events for this day.").pack()
        else:
            for event_id, desc in events:
                frame = ctk.CTkFrame(self.event_list_frame); frame.pack(fill="x", pady=2)
                lbl = ctk.CTkLabel(frame, text=desc, wraplength=self.main_frame.winfo_width() - 80); lbl.pack(side="left", expand=True, fill="x", padx=5)
                del_btn = ctk.CTkButton(frame, text="X", width=30, fg_color="red", hover_color="#C00000", command=lambda e_id=event_id: self.delete_event(e_id)); del_btn.pack(side="right")

    def load_all_event_markers(self):
        if hasattr(self, 'cal') and self.cal.winfo_exists(): self.cal.calevent_remove("all")
        if not self.db_path: return
        conn = sqlite3.connect(self.db_path); cursor = conn.cursor(); cursor.execute("SELECT DISTINCT event_date FROM events"); dates = cursor.fetchall(); conn.close()
        for date_tuple in dates:
            date_obj = datetime.strptime(date_tuple[0], "%Y-%m-%d")
            self.cal.calevent_create(date_obj, 'Event', 'event')
        self.cal.tag_config('event', background='#6366F1', foreground='#FFFFFF')

    def add_event(self):
        if not self.db_path: messagebox.showerror("Error", "Please open a project first."); return
        selected_date = self.cal.get_date(); db_date = datetime.strptime(selected_date, "%m/%d/%y").strftime("%Y-%m-%d")
        desc = CustomDialog(self, title=f"New Event on {db_date}", prompt="Enter event description:").result
        if not desc: return
        conn = sqlite3.connect(self.db_path); cursor = conn.cursor(); cursor.execute("INSERT INTO events (event_date, description) VALUES (?,?)", (db_date, desc)); conn.commit(); conn.close()
        self.on_date_select(); self.load_all_event_markers(); self.load_upcoming_events()

    def delete_event(self, event_id):
        if messagebox.askyesno("Confirm Delete", "Are you sure you want to delete this event?"):
            conn = sqlite3.connect(self.db_path); cursor = conn.cursor(); cursor.execute("DELETE FROM events WHERE id = ?", (event_id,)); conn.commit(); conn.close()
            self.on_date_select(); self.load_all_event_markers(); self.load_upcoming_events()

    def load_upcoming_events(self):
        for widget in self.upcoming_events_frame.winfo_children(): widget.destroy()
        if not self.db_path: return
        today_str = datetime.now().strftime("%Y-%m-%d")
        conn = sqlite3.connect(self.db_path); cursor = conn.cursor()
        cursor.execute("SELECT event_date, description FROM events WHERE event_date >= ? ORDER BY event_date ASC", (today_str,))
        events = cursor.fetchall(); conn.close()
        if not events: ctk.CTkLabel(self.upcoming_events_frame, text="No upcoming events.").pack(pady=5)
        else:
            for date_str, desc in events:
                frame = ctk.CTkFrame(self.upcoming_events_frame, fg_color="transparent"); frame.pack(fill="x", pady=4)
                lbl_date = ctk.CTkLabel(frame, text=date_str, font=ctk.CTkFont(weight="bold")); lbl_date.pack(anchor="w")
                lbl_desc = ctk.CTkLabel(frame, text=desc, wraplength=240, justify="left"); lbl_desc.pack(anchor="w", padx=(10,0))

    def switch_to_timetracker_view(self):
        self.current_note_id = None; self.current_table_id = None; self.clear_main_frame()
        if not self.db_path:
            ctk.CTkLabel(self.main_frame, text="Please open a project to use the Time Tracker.").pack(pady=50)
            return

        self.main_frame.grid_columnconfigure(0, weight=1)
        self.main_frame.grid_rowconfigure(0, weight=0)
        self.main_frame.grid_rowconfigure(1, weight=0)
        self.main_frame.grid_rowconfigure(2, weight=1)

        controls_frame = ctk.CTkFrame(self.main_frame)
        controls_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=10)
        controls_frame.grid_columnconfigure(0, weight=1)
        
        self.timer_label = ctk.CTkLabel(controls_frame, text="0hrs 0mins 0secs", font=ctk.CTkFont(size=24, weight="bold"))
        self.timer_label.grid(row=0, column=0, padx=20, pady=10, sticky="w")
        
        btn_frame = ctk.CTkFrame(controls_frame, fg_color="transparent")
        btn_frame.grid(row=0, column=1, padx=20, pady=10)

        self.btn_pause_resume = ctk.CTkButton(btn_frame, text="Pause Timer", command=self.toggle_pause_resume, width=150)
        self.btn_pause_resume.pack(side="left", padx=(0, 10))

        self.btn_start_stop = ctk.CTkButton(btn_frame, text="Start Timer", command=self.toggle_timer, width=150)
        self.btn_start_stop.pack(side="right")
        
        self.lbl_total_time = ctk.CTkLabel(controls_frame, text="Total Project Time: Calculating...")
        self.lbl_total_time.grid(row=1, column=0, columnspan=2, padx=20, pady=(0, 10), sticky="w")

        list_label = ctk.CTkLabel(self.main_frame, text="Time Entries History", font=ctk.CTkFont(size=18, weight="bold"))
        list_label.grid(row=1, column=0, sticky="ew", padx=10, pady=(10, 0))
        
        self.entries_frame = ctk.CTkScrollableFrame(self.main_frame, label_text="") 
        self.entries_frame.grid(row=2, column=0, sticky="nsew", padx=10, pady=(0, 10))
        self.entries_frame.grid_columnconfigure(0, weight=1)
        
        self.check_for_running_timer() 
        self.load_time_entries()
        self.update_timer_display()

    def check_for_running_timer(self):
        if not self.db_path: return
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute("SELECT id, start_time, pauses FROM time_entries WHERE end_time IS NULL ORDER BY id DESC LIMIT 1")
        result = cursor.fetchone()
        conn.close()
        
        if not hasattr(self, 'btn_start_stop') or not self.btn_start_stop or not self.btn_start_stop.winfo_exists():
            if result:
                start_time_str = result[1]
                self.timer_running = True
                self.start_time = datetime.strptime(start_time_str, "%Y-%m-%d %H:%M:%S")
                self.current_running_entry_id = result[0]
                self.pauses = json.loads(result[2])
                
                if self.pauses and self.pauses[-1] is not None and 'end' not in self.pauses[-1]:
                    self.timer_paused = True
            else:
                self.timer_running = False
                self.timer_paused = False
            return

        if result:
            entry_id, start_time_str, pauses_json = result
            self.timer_running = True
            self.start_time = datetime.strptime(start_time_str, "%Y-%m-%d %H:%M:%S")
            self.current_running_entry_id = entry_id 
            self.pauses = json.loads(pauses_json)
            
            if self.pauses and self.pauses[-1] is not None and 'end' not in self.pauses[-1]:
                self.timer_paused = True
                self.btn_pause_resume.configure(text="Resume Timer", fg_color="#3B8ED0", hover_color="#307FB0")
                self.timer_label.configure(text="PAUSED")
            else:
                self.timer_paused = False
                self.btn_pause_resume.configure(text="Pause Timer", fg_color="#F8A701", hover_color="#D58E00")
                
            self.btn_start_stop.configure(text="Stop Timer", fg_color="red", hover_color="#C00000", state="normal")
            self.btn_pause_resume.configure(state="normal")
            
        else:
            self.timer_running = False
            self.timer_paused = False
            self.start_time = None
            self.pauses = []
            self.current_running_entry_id = None
            
            self.btn_start_stop.configure(text="Start Timer", fg_color="#2E8B57", hover_color="#20603C", state="normal")
            self.btn_pause_resume.configure(text="Pause Timer", state="disabled")

    def toggle_timer(self):
        if not self.db_path:
            messagebox.showerror("Error", "No project loaded.")
            return

        if self.timer_running:
            self.stop_timer()
        else:
            self.start_new_timer()

    def start_new_timer(self):
        if self.timer_running: return
        
        self.start_time = datetime.now()
        start_time_str = self.start_time.strftime("%Y-%m-%d %H:%M:%S")
        self.pauses = []
        
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute("INSERT INTO time_entries (start_time, notes, pauses) VALUES (?, ?, ?)", (start_time_str, "Working on project", json.dumps(self.pauses)))
        self.current_running_entry_id = cursor.lastrowid 
        conn.commit()
        conn.close()

        self.timer_running = True
        self.timer_paused = False
        self.btn_start_stop.configure(text="Stop Timer", fg_color="red", hover_color="#C00000")
        self.btn_pause_resume.configure(text="Pause Timer", fg_color="#F8A701", hover_color="#D58E00", state="normal")
        self.update_timer_display()

    def toggle_pause_resume(self):
        if not self.timer_running: 
            messagebox.showerror("Error", "Cannot pause: Timer is not running.")
            return

        if not self.timer_paused:
            self.timer_paused = True
            pause_start_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f")
            self.pauses.append({'start': pause_start_time})
            
            if self.timer_update_id:
                self.after_cancel(self.timer_update_id)
                self.timer_update_id = None
                
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute("UPDATE time_entries SET pauses = ? WHERE id = ?", (json.dumps(self.pauses), self.current_running_entry_id))
            conn.commit()
            conn.close()

            self.btn_pause_resume.configure(text="Resume Timer", fg_color="#3B8ED0", hover_color="#307FB0")
            self.timer_label.configure(text="PAUSED")

        else:
            self.timer_paused = False
            pause_end_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f")
            if self.pauses and 'end' not in self.pauses[-1]:
                self.pauses[-1]['end'] = pause_end_time
            
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute("UPDATE time_entries SET pauses = ? WHERE id = ?", (json.dumps(self.pauses), self.current_running_entry_id))
            conn.commit()
            conn.close()

            self.btn_pause_resume.configure(text="Pause Timer", fg_color="#F8A701", hover_color="#D58E00")
            self.update_timer_display()

    def stop_timer(self):
        if not self.timer_running: return
        
        if self.timer_paused:
            self.toggle_pause_resume() 

        end_time = datetime.now()
        end_time_str = end_time.strftime("%Y-%m-%d %H:%M:%S")
        start_time_str = self.start_time.strftime("%Y-%m-%d %H:%M:%S")
        pauses_json = json.dumps(self.pauses)
        
        duration_seconds = calculate_actual_duration(start_time_str, end_time_str, pauses_json)
        
        if self.timer_update_id:
            self.after_cancel(self.timer_update_id)
            self.timer_update_id = None

        dialog = CustomDialog(self, title="Timer Stopped", prompt="Add notes for this time entry:", initial_value="")
        notes = dialog.result
        if notes is None:
            notes = "" 

        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute("UPDATE time_entries SET end_time = ?, duration_seconds = ?, notes = ?, pauses = ? WHERE id = ?", 
                       (end_time_str, duration_seconds, notes, pauses_json, self.current_running_entry_id))
        conn.commit()
        conn.close()

        self.timer_running = False
        self.timer_paused = False
        self.start_time = None
        self.pauses = []
        self.current_running_entry_id = None
        self.btn_start_stop.configure(text="Start Timer", fg_color="#2E8B57", hover_color="#20603C")
        self.btn_pause_resume.configure(text="Pause Timer", state="disabled")
        self.timer_label.configure(text="0hrs 0mins 0secs")
        
        self.load_time_entries()

    def update_timer_display(self):
        if self.timer_running and not self.timer_paused and self.start_time:
            elapsed_total = datetime.now() - self.start_time
            total_pause = calculate_total_pause_time(json.dumps(self.pauses))
            actual_elapsed = elapsed_total - total_pause
            
            self.timer_label.configure(text=format_timedelta(actual_elapsed))
            self.timer_update_id = self.after(1000, self.update_timer_display)
        elif self.timer_label and self.timer_label.winfo_exists() and not self.timer_paused:
            self.timer_label.configure(text="0hrs 0mins 0secs")

    def load_time_entries(self):
        if not self.db_path: return
        
        if hasattr(self, 'entries_frame') and self.entries_frame.winfo_exists():
            for widget in self.entries_frame.winfo_children(): widget.destroy()
        else:
            return

        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute("SELECT id, start_time, end_time, duration_seconds, notes, pauses FROM time_entries ORDER BY start_time DESC")
        entries = cursor.fetchall()
        conn.close()
        
        total_duration = timedelta()
        
        if not entries:
            ctk.CTkLabel(self.entries_frame, text="No time entries recorded yet.").pack(pady=10)
            self.lbl_total_time.configure(text="Total Project Time: 0hrs 0mins")
            return

        for i, (entry_id, start_str, end_str, duration_seconds, notes, pauses_json) in enumerate(entries):
            duration = timedelta(seconds=duration_seconds)
            
            if self.timer_running and entry_id == self.current_running_entry_id:
                continue

            if duration_seconds > 0:
                total_duration += duration

            entry_frame = ctk.CTkFrame(self.entries_frame)
            entry_frame.pack(fill="x", padx=5, pady=5)
            entry_frame.grid_columnconfigure(0, weight=1)
            
            time_info_frame = ctk.CTkFrame(entry_frame, fg_color="transparent")
            time_info_frame.grid(row=0, column=0, sticky="ew", padx=5, pady=(5,0))
            
            try:
                start_dt = datetime.strptime(start_str, "%Y-%m-%d %H:%M:%S")
                start_time_disp = start_dt.strftime("%d/%m/%Y %H:%M")
                end_time_disp = datetime.strptime(end_str, "%Y-%m-%d %H:%M:%S").strftime("%H:%M") if end_str else "N/A"
            except (ValueError, TypeError):
                 start_time_disp = "N/A"
                 end_time_disp = "N/A"

            lbl_timespan_text = f"**{format_timedelta(duration)}** ({start_time_disp} - {end_time_disp})"
            lbl_timespan = ctk.CTkLabel(time_info_frame, text=lbl_timespan_text, justify="left", font=ctk.CTkFont(weight="bold"))
            lbl_timespan.grid(row=0, column=0, sticky="w", padx=5)
            
            entry_data_for_edit = (entry_id, start_str, end_str, duration_seconds, notes, pauses_json)
            btn_edit = ctk.CTkButton(time_info_frame, text="Edit", width=60, command=lambda data=entry_data_for_edit: self.edit_time_entry(data))
            btn_edit.grid(row=0, column=1, padx=(5,2))
            btn_delete = ctk.CTkButton(time_info_frame, text="X", width=30, fg_color="red", hover_color="#C00000", command=lambda e_id=entry_id: self.delete_time_entry(e_id))
            btn_delete.grid(row=0, column=2, padx=(0,5))
            
            lbl_notes = ctk.CTkLabel(entry_frame, text=f"Notes: {notes}", wraplength=900, justify="left", anchor="w")
            lbl_notes.grid(row=1, column=0, sticky="ew", padx=5, pady=(0, 5))
            
        self.lbl_total_time.configure(text=f"Total Project Time: {format_timedelta(total_duration)}")

    def edit_time_entry(self, entry_data):
        if not self.db_path: return
        
        dialog = TimeEntryEditDialog(self, entry_data)
        updated_data = dialog.result

        if updated_data:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute("UPDATE time_entries SET start_time = ?, end_time = ?, duration_seconds = ?, notes = ?, pauses = ? WHERE id = ?", 
                           (updated_data['start_time'], updated_data['end_time'], updated_data['duration_seconds'], updated_data['notes'], updated_data['pauses_json'], updated_data['id']))
            conn.commit()
            conn.close()
            self.load_time_entries()
            messagebox.showinfo("Success", "Time entry updated.")

    def delete_time_entry(self, entry_id):
        if not self.db_path: return
        if messagebox.askyesno("Confirm Delete", "Are you sure you want to delete this time entry?"):
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute("DELETE FROM time_entries WHERE id = ?", (entry_id,))
            conn.commit()
            conn.close()
            self.load_time_entries()
            
    def clear_main_frame(self):
        if self.note_save_timer is not None: self.after_cancel(self.note_save_timer); self.note_save_timer = None
        if self.table_save_timer is not None: self.after_cancel(self.table_save_timer); self.table_save_timer = None
        if self.timer_update_id: self.after_cancel(self.timer_update_id); self.timer_update_id = None
        for widget in self.main_frame.winfo_children(): widget.destroy()
        
    def load_project_data(self):
        if not self.db_path: return
        self._update_folder_dropdown_options()
        self.load_notes_list(); self.load_upcoming_tasks(); self.load_links(); self.load_tables_list(); self.load_upcoming_events()
        self.switch_to_notes_view()

    # --- Cross-Project Weekly Summary Generator ---
    def generate_weekly_summary(self):
        days_input = CustomDialog(self, title="Report Period", prompt="Enter number of days back to summarize:", initial_value="7").result
        if not days_input:
            return
        
        try:
            days_back = int(days_input)
        except ValueError:
            messagebox.showerror("Error", "Please enter a valid whole number of days.")
            return

        cutoff_date = datetime.now() - timedelta(days=days_back)
        cutoff_str = cutoff_date.strftime("%Y-%m-%d %H:%M:%S")

        project_files = [f for f in os.listdir(self.projects_dir) if f.endswith(".db")]
        if not project_files:
            messagebox.showinfo("No Projects", "No project database files found to summarize.")
            return

        date_range_disp = f"{cutoff_date.strftime('%d/%m/%Y')} – {datetime.now().strftime('%d/%m/%Y')}"

        plain_lines = [
            f"WEEKLY ACTIVITY SUMMARY ({date_range_disp})",
            "═" * 70,
            ""
        ]

        html_blocks = [
            f'<div style="font-family: \'Segoe UI\', Calibri, Arial, sans-serif; color: #1E293B; font-size: 14px; line-height: 1.5; max-width: 800px;">',
            f'<div style="border-bottom: 2px solid #0284C7; padding-bottom: 8px; margin-bottom: 20px;">'
            f'<h2 style="margin: 0; color: #0F172A; font-size: 20px;">Weekly Activity Summary</h2>'
            f'<span style="color: #64748B; font-size: 13px;">Reporting Period: <strong>{date_range_disp}</strong> (Past {days_back} Days)</span>'
            f'</div>'
        ]

        total_tracked_seconds = 0
        total_completed_tasks_count = 0
        total_open_tasks_count = 0
        total_open_questions_count = 0

        for db_file in sorted(project_files):
            project_name = os.path.splitext(db_file)[0]
            db_full_path = os.path.join(self.projects_dir, db_file)
            
            try:
                conn = sqlite3.connect(db_full_path)
                cursor = conn.cursor()

                # 1. Latest Status Note
                status_notes = []
                try:
                    cursor.execute(
                        "SELECT title, content, last_modified_date FROM notes WHERE folder = 'Status' ORDER BY last_modified_date DESC LIMIT 1"
                    )
                    row = cursor.fetchone()
                    if row:
                        title, content, last_mod = row
                        if last_mod and last_mod >= cutoff_str:
                            plain_text = content
                            try:
                                note_data = json.loads(content)
                                if isinstance(note_data, dict) and "text" in note_data:
                                    plain_text = note_data["text"]
                            except Exception:
                                pass
                            
                            cleaned_lines = []
                            for line in plain_text.splitlines():
                                line_s = line.strip()
                                if line_s:
                                    for prefix in ["•", "-", "*"]:
                                        if line_s.startswith(prefix):
                                            line_s = line_s[len(prefix):].strip()
                                            break
                                    cleaned_lines.append(line_s)
                            if cleaned_lines:
                                status_notes.append((title, cleaned_lines))
                except sqlite3.OperationalError:
                    pass

                # 2. Completed Tasks
                completed_tasks = []
                try:
                    cursor.execute(
                        "SELECT task_description, completed_date, completion_notes FROM tasks WHERE is_completed = 1 AND completed_date >= ? ORDER BY completed_date ASC",
                        (cutoff_str,)
                    )
                    completed_tasks = cursor.fetchall()
                except sqlite3.OperationalError:
                    pass

                # 3. Open Tasks with Dates
                open_tasks = []
                try:
                    cursor.execute("SELECT task_description, start_date, due_date FROM tasks WHERE is_completed = 0 ORDER BY id ASC")
                    open_tasks = cursor.fetchall()
                except sqlite3.OperationalError:
                    pass

                # 4. Open Questions / Blockers
                open_questions = []
                try:
                    cursor.execute("SELECT question_text FROM questions WHERE is_answered = 0 ORDER BY id ASC")
                    open_questions = cursor.fetchall()
                except sqlite3.OperationalError:
                    pass

                # 5. Logged Time Entries
                logged_seconds = 0
                time_work_notes = []
                try:
                    cursor.execute(
                        "SELECT duration_seconds, notes FROM time_entries WHERE start_time >= ? ORDER BY start_time ASC",
                        (cutoff_str,)
                    )
                    time_entries = cursor.fetchall()
                    for dur, notes in time_entries:
                        if dur: logged_seconds += dur
                        if notes and notes.strip() and notes.strip() != "Working on project":
                            time_work_notes.append(f"{notes.strip()} ({format_timedelta(timedelta(seconds=dur))})")
                except sqlite3.OperationalError:
                    pass

                conn.close()

                if not status_notes and not completed_tasks and logged_seconds == 0 and not open_tasks and not open_questions:
                    continue

                plain_lines.append(f"PROJECT: {project_name.upper()}")
                plain_lines.append("─" * 45)

                html_blocks.append(
                    f'<div style="background-color: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 6px; padding: 14px; margin-bottom: 16px;">'
                    f'<div style="display: flex; justify-content: space-between; align-items: baseline; margin-bottom: 10px;">'
                    f'<span style="font-size: 16px; font-weight: bold; color: #0F172A;">{project_name}</span>'
                )

                if logged_seconds > 0:
                    total_tracked_seconds += logged_seconds
                    time_str = format_timedelta(timedelta(seconds=logged_seconds))
                    plain_lines.append(f"  Tracked Time: {time_str}")
                    html_blocks.append(f'<span style="font-size: 12px; background-color: #E0E7FF; color: #3730A3; padding: 2px 8px; border-radius: 4px; font-weight: 600;">⏱ {time_str}</span>')

                html_blocks.append('</div>')

                if status_notes:
                    for n_title, lines in status_notes:
                        plain_lines.append(f"  Status ({n_title}):")
                        html_blocks.append(f'<div style="margin-top: 8px;"><strong style="font-size: 13px; color: #0369A1;">Key Highlights ({n_title}):</strong><ul style="margin: 4px 0 10px 20px; padding: 0;">')
                        for l in lines:
                            plain_lines.append(f"    • {l}")
                            html_blocks.append(f'<li style="margin-bottom: 3px; color: #334155;">{l}</li>')
                        html_blocks.append('</ul></div>')

                if completed_tasks:
                    total_completed_tasks_count += len(completed_tasks)
                    plain_lines.append("  Completed Work:")
                    html_blocks.append('<div style="margin-top: 8px;"><strong style="font-size: 13px; color: #15803D;">Completed Actions:</strong><ul style="margin: 4px 0 10px 20px; padding: 0;">')
                    for desc, c_date, c_note in completed_tasks:
                        date_tag = ""
                        if c_date:
                            try:
                                dt = datetime.strptime(c_date, "%Y-%m-%d %H:%M:%S")
                                date_tag = f" ({dt.strftime('%d/%m')})"
                            except ValueError:
                                pass
                        plain_lines.append(f"    ✔ {desc}{date_tag}")
                        html_blocks.append(f'<li style="margin-bottom: 3px; color: #1E293B;"><strong>✔ {desc}</strong><span style="color: #64748B; font-size: 12px;">{date_tag}</span>')
                        if c_note and c_note.strip():
                            plain_lines.append(f"       Note: {c_note.strip()}")
                            html_blocks.append(f'<div style="color: #64748B; font-size: 12px; margin-left: 5px; font-style: italic;">↳ {c_note.strip()}</div>')
                        html_blocks.append('</li>')
                    html_blocks.append('</ul></div>')

                if open_tasks:
                    total_open_tasks_count += len(open_tasks)
                    plain_lines.append("  Outstanding / In Progress:")
                    html_blocks.append('<div style="margin-top: 8px;"><strong style="font-size: 13px; color: #B45309;">Outstanding Tasks / Next Steps:</strong><ul style="margin: 4px 0 8px 20px; padding: 0;">')
                    for (task_desc, s_date, d_date) in open_tasks:
                        date_badge, is_overdue = self._format_date_badge(s_date, d_date)
                        date_str_plain = f" [{date_badge}]" if date_badge else ""
                        plain_lines.append(f"    ◻ {task_desc}{date_str_plain}")
                        
                        html_date = ""
                        if date_badge:
                            color = "#DC2626" if is_overdue else "#0284C7"
                            html_date = f' <span style="color: {color}; font-size: 11px; font-weight: 600;">{date_badge}</span>'
                        html_blocks.append(f'<li style="margin-bottom: 3px; color: #475569;">◻ {task_desc}{html_date}</li>')
                    html_blocks.append('</ul></div>')

                if open_questions:
                    total_open_questions_count += len(open_questions)
                    plain_lines.append("  Open Questions / Blockers:")
                    html_blocks.append('<div style="margin-top: 8px;"><strong style="font-size: 13px; color: #DC2626;">Open Questions / Blockers:</strong><ul style="margin: 4px 0 8px 20px; padding: 0;">')
                    for (q_desc,) in open_questions:
                        plain_lines.append(f"    ? {q_desc}")
                        html_blocks.append(f'<li style="margin-bottom: 3px; color: #DC2626;">? {q_desc}</li>')
                    html_blocks.append('</ul></div>')

                plain_lines.append("")
                html_blocks.append('</div>')

            except Exception:
                continue

        plain_lines.append("═" * 70)
        plain_lines.append(f"Total Completed Items:    {total_completed_tasks_count}")
        plain_lines.append(f"Total Outstanding Tasks:  {total_open_tasks_count}")
        plain_lines.append(f"Total Open Questions:     {total_open_questions_count}")
        if total_tracked_seconds > 0:
            plain_lines.append(f"Total Project Time Log:   {format_timedelta(timedelta(seconds=total_tracked_seconds))}")

        html_blocks.append(
            f'<div style="border-top: 1px solid #CBD5E1; padding-top: 12px; margin-top: 20px; font-size: 13px; color: #475569;">'
            f'<strong>Summary Totals:</strong> {total_completed_tasks_count} Completed Tasks &nbsp;|&nbsp; '
            f'{total_open_tasks_count} Outstanding Tasks &nbsp;|&nbsp; '
            f'<span style="color: #DC2626;"><strong>{total_open_questions_count} Open Questions</strong></span>'
        )
        if total_tracked_seconds > 0:
            html_blocks.append(f' &nbsp;|&nbsp; <strong>Time Logged:</strong> {format_timedelta(timedelta(seconds=total_tracked_seconds))}')
        html_blocks.append('</div></div>')

        full_plain_text = "\n".join(plain_lines)
        full_html_text = "".join(html_blocks)
        ReportDisplayDialog(self, full_plain_text, full_html_text)


if __name__ == "__main__":
    app = ProjectApp()
    app.mainloop()