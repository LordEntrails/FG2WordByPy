# gui_app.py
import os
import sys
import json
import importlib
import customtkinter as ctk
from tkinter import filedialog
import main as pipeline_engine

CONFIG_FILE = ".config.json"

class ConsoleRedirector:
    def __init__(self, text_widget):
        self.text_widget = text_widget
    def write(self, string):
        self.text_widget.configure(state="normal")
        self.text_widget.insert("end", string)
        self.text_widget.see("end")
        self.text_widget.configure(state="disabled")
    def flush(self):
        pass

class FGWorkbookGUI(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("FG2WordByPy")
        self.geometry("720x720")
        ctk.set_appearance_mode("dark")
        
        self.last_mod_path = ""
        self.load_cached_config()
        self.create_widgets()

        if self.last_mod_path and os.path.exists(self.last_mod_path):
            self.mod_entry.insert(0, self.last_mod_path)
            self.auto_generate_output_name(self.last_mod_path)

        sys.stdout = ConsoleRedirector(self.console_output)
        sys.stderr = ConsoleRedirector(self.console_output)

    def create_widgets(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(6, weight=1)

        # --- MODULE FILE LAYER ---
        self.mod_label = ctk.CTkLabel(self, text="Fantasy Grounds Module Path (.mod):", font=("Helvetica", 13, "bold"))
        self.mod_label.grid(row=0, column=0, padx=20, pady=(15, 5), sticky="w")

        self.mod_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.mod_frame.grid(row=1, column=0, padx=20, pady=5, sticky="ew")
        self.mod_frame.grid_columnconfigure(0, weight=1)

        self.mod_entry = ctk.CTkEntry(self.mod_frame, placeholder_text="Browse or paste .mod source archive pathway...")
        self.mod_entry.grid(row=0, column=0, padx=(0, 10), sticky="ew")

        self.browse_btn = ctk.CTkButton(self.mod_frame, text="Browse", width=100, command=self.browse_mod_file)
        self.browse_btn.grid(row=0, column=1, sticky="e")

        # --- OUTPUT CHRONICLE LAYER ---
        self.out_label = ctk.CTkLabel(self, text="Output Document Name (.docx):", font=("Helvetica", 13, "bold"))
        self.out_label.grid(row=2, column=0, padx=20, pady=(10, 5), sticky="w")
        self.out_entry = ctk.CTkEntry(self, placeholder_text="Calculated matching output name target frame...")
        self.out_entry.grid(row=3, column=0, padx=20, pady=5, sticky="ew")

        # --- CONTROLS FRAME ---
        self.control_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.control_frame.grid(row=4, column=0, padx=20, pady=10, sticky="ew")
        self.control_frame.columnconfigure(0, weight=3)
        self.control_frame.columnconfigure(1, weight=1)

        self.compile_btn = ctk.CTkButton(self.control_frame, text="⚡ Compile Campaign Workbook", font=("Helvetica", 13, "bold"), height=40, fg_color="#1f538d", hover_color="#14375e", command=self.execute_extraction_pipeline)
        self.compile_btn.grid(row=0, column=0, padx=(0, 10), sticky="ew")

        self.reload_btn = ctk.CTkButton(self.control_frame, text="🔄 Reload Engine", font=("Helvetica", 13, "bold"), height=40, fg_color="#2b2b2b", hover_color="#3e3e3e", command=self.reload_engine_modules)
        self.reload_btn.grid(row=0, column=1, sticky="ew")

        # --- MONITORING CONSOLE FRAME ---
        self.console_label = ctk.CTkLabel(self, text="System Compile Telemetry Log Stream:", font=("Helvetica", 12, "bold"))
        self.console_label.grid(row=5, column=0, padx=20, pady=(5, 0), sticky="w")

        self.console_output = ctk.CTkTextbox(self, font=("Consolas", 11), fg_color="#1e1e1e", border_color="#333333", border_width=1, text_color="#d4d4d4", state="disabled")
        self.console_output.grid(row=6, column=0, padx=20, pady=(5, 15), sticky="nsew")

    def reload_engine_modules(self):
        """Purges cached script instances from system memory instantly."""
        print("\n[SYSTEM] Dropping cached files and resetting parsing layers...")
        try:
            importlib.reload(pipeline_engine)
            # Cascade reloads across dependent modules safely
            if "story_parser" in sys.modules: importlib.reload(sys.modules["story_parser"])
            if "npc_parser" in sys.modules: importlib.reload(sys.modules["npc_parser"])
            if "item_parser" in sys.modules: importlib.reload(sys.modules["item_parser"])
            if "vehicle_parser" in sys.modules: importlib.reload(sys.modules["vehicle_parser"])
            if "npc_renderer" in sys.modules: importlib.reload(sys.modules["npc_renderer"])
            if "word_renderer" in sys.modules: importlib.reload(sys.modules["word_renderer"])
            print("[SUCCESS] Core pipeline components refreshed. Ready to test.")
        except Exception as e:
            print(f"[!] Engine reload failed: {e}")

    def browse_mod_file(self):
        initial_dir = os.path.dirname(self.last_mod_path) if self.last_mod_path else os.getcwd()
        selected_file = filedialog.askopenfilename(
            initialdir=initial_dir,
            title="Select Fantasy Grounds Module Package Archive",
            filetypes=[("Fantasy Grounds Module", "*.mod"), ("All Files", "*.*")]
        )
        if selected_file:
            self.mod_entry.delete(0, "end")
            self.mod_entry.insert(0, selected_file)
            self.auto_generate_output_name(selected_file)
            self.save_cached_config(selected_file)

    def auto_generate_output_name(self, filepath):
        if not filepath or not os.path.exists(filepath): return
        try:
            internal_name = pipeline_engine.main(custom_mod_file=filepath, custom_output_file="__GET_INTERNAL_NAME__")
            for char in ['/', '\\', '?', '%', '*', ':', '|', '"', '<', '>']:
                internal_name = internal_name.replace(char, "_")
            self.out_entry.delete(0, "end")
            self.out_entry.insert(0, f"{internal_name}.docx")
        except Exception:
            name_part = os.path.splitext(os.path.basename(filepath))[0]
            self.out_entry.delete(0, "end")
            self.out_entry.insert(0, f"{name_part}.docx")

    def load_cached_config(self):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r") as f:
                    self.last_mod_path = json.load(f).get("last_mod_path", "")
            except Exception: pass

    def save_cached_config(self, filepath):
        self.last_mod_path = filepath
        try:
            with open(CONFIG_FILE, "w") as f: json.dump({"last_mod_path": filepath}, f)
        except Exception: pass

    def execute_extraction_pipeline(self):
        mod_target = self.mod_entry.get().strip()
        out_target = self.out_entry.get().strip()
        if not mod_target or not out_target: return
        
        self.save_cached_config(mod_target)
        self.console_output.configure(state="normal")
        self.console_output.delete("1.0", "end")
        self.console_output.configure(state="disabled")

        self.compile_btn.configure(state="disabled", text="Compiling Modules...")
        self.update_idletasks()

        pipeline_engine.main(custom_mod_file=mod_target, custom_output_file=out_target)
        self.compile_btn.configure(state="normal", text="⚡ Compile Campaign Workbook")

if __name__ == "__main__":
    app = FGWorkbookGUI()
    app.mainloop()