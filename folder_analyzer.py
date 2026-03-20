import os
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import datetime
import logging
import threading
import pythoncom
import win32com.client
import pandas as pd
import subprocess

log_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "folder_analyzer_error.log")
logging.basicConfig(filename=log_file, level=logging.ERROR, 
                    format='%(asctime)s - %(levelname)s - %(message)s')

class FolderAnalyzerLogic:
    def __init__(self, target_folder, output_excel_path, selected_columns, progress_callback, finish_callback):
        self.target_folder = target_folder
        self.output_excel_path = output_excel_path
        self.selected_columns = selected_columns
        self.progress_callback = progress_callback
        self.finish_callback = finish_callback

    def run(self):
        pythoncom.CoInitialize()
        try:
            self._analyze()
        except Exception as e:
            logging.error("Fatal error during analysis", exc_info=True)
            self.finish_callback(False, str(e), None)
        finally:
            pythoncom.CoUninitialize()
            
    def _analyze(self):
        data = []
        
        self.progress_callback("Counting files to set up progress bar...", 0, 0)
        total_items = 0
        for root, dirs, files in os.walk(self.target_folder):
            total_items += len(dirs) + len(files)
            
        if total_items == 0:
            self.finish_callback(False, "The selected folder is empty or not accessible.", None)
            return

        shell = win32com.client.Dispatch("Shell.Application")
        processed_items = 0
        
        for root, dirs, files in os.walk(self.target_folder):
            
            namespace = None
            author_idx = 20
            last_saved_by_idx = 105
            try:
                namespace = shell.NameSpace(os.path.abspath(root))
                if namespace:
                    for i in range(300):
                        prop_name = str(namespace.GetDetailsOf(None, i)).strip().lower()
                        if prop_name in ["authors", "autoren", "author", "autor"]:
                            author_idx = i
                        elif prop_name in ["last saved by", "zuletzt gespeichert von"]:
                            last_saved_by_idx = i
            except Exception as e:
                logging.error(f"Failed to get properties namespace for {root}", exc_info=True)
                
            for dir_name in dirs:
                dir_path = os.path.join(root, dir_name)
                try:
                    stat = os.stat(dir_path)
                    item_data = {
                        "Name": dir_name,
                        "Extension": "Folder",
                        "Date accessed": datetime.datetime.fromtimestamp(stat.st_atime).strftime('%Y-%m-%d %H:%M:%S'),
                        "Date modified": datetime.datetime.fromtimestamp(stat.st_mtime).strftime('%Y-%m-%d %H:%M:%S'),
                        "Date created": datetime.datetime.fromtimestamp(stat.st_ctime).strftime('%Y-%m-%d %H:%M:%S'),
                        "Author": "",
                        "Last time saved by": "",
                        "Folder Path": root
                    }
                    data.append({k: v for k, v in item_data.items() if k in self.selected_columns})
                except Exception as e:
                    logging.error(f"Error reading folder stat: {dir_path}", exc_info=True)
                
                processed_items += 1
                if processed_items % 50 == 0:
                    self.progress_callback(f"Scanning: {processed_items}/{total_items} items", processed_items, total_items)

            for file_name in files:
                file_path = os.path.join(root, file_name)
                try:
                    stat = os.stat(file_path)
                    ext = os.path.splitext(file_name)[1]
                    
                    author = ""
                    last_saved = ""
                    
                    if namespace and ("Author" in self.selected_columns or "Last time saved by" in self.selected_columns):
                        item = namespace.ParseName(file_name)
                        if item:
                            if "Author" in self.selected_columns:
                                author = namespace.GetDetailsOf(item, author_idx)
                            if "Last time saved by" in self.selected_columns:
                                last_saved = namespace.GetDetailsOf(item, last_saved_by_idx)
                            
                    item_data = {
                        "Name": file_name,
                        "Extension": ext,
                        "Date accessed": datetime.datetime.fromtimestamp(stat.st_atime).strftime('%Y-%m-%d %H:%M:%S'),
                        "Date modified": datetime.datetime.fromtimestamp(stat.st_mtime).strftime('%Y-%m-%d %H:%M:%S'),
                        "Date created": datetime.datetime.fromtimestamp(stat.st_ctime).strftime('%Y-%m-%d %H:%M:%S'),
                        "Author": author,
                        "Last time saved by": last_saved,
                        "Folder Path": root
                    }
                    data.append({k: v for k, v in item_data.items() if k in self.selected_columns})
                except Exception as e:
                    logging.error(f"Error reading file info: {file_path}", exc_info=True)

                processed_items += 1
                if processed_items % 50 == 0:
                    self.progress_callback(f"Scanning: {processed_items}/{total_items} items", processed_items, total_items)
        
        self.progress_callback(f"Saving {processed_items} items to Excel. Please wait...", total_items, total_items)
        try:
            if not data:
                self.finish_callback(False, "No data found to save.", None)
                return
                
            df = pd.DataFrame(data)
            # Ensure columns are in the correct order (as selected)
            ordered_cols = [col for col in self.selected_columns if col in df.columns]
            df = df[ordered_cols]
            df.to_excel(self.output_excel_path, index=False)
            self.finish_callback(True, "Success", self.output_excel_path)
        except Exception as e:
            logging.error("Failed writing Excel file", exc_info=True)
            self.finish_callback(False, f"Target file might be open or read-only:\n{str(e)}", None)


class FolderAnalyzerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Advanced Folder Content Analyzer")
        self.root.geometry("550x300")
        self.root.resizable(False, False)
        
        style = ttk.Style(self.root)
        if "vista" in style.theme_names():
            style.theme_use("vista")
        elif "clam" in style.theme_names():
            style.theme_use("clam")
            
        self.all_columns = [
            "Name", "Extension", "Date accessed", 
            "Date modified", "Date created", "Author", 
            "Last time saved by", "Folder Path"
        ]
        
        # State variables for checkboxes
        self.column_vars = {col: tk.BooleanVar(value=True) for col in self.all_columns}
        
        self.logic = None
        self.current_output_path = None
        
        # Top Frame with label and settings button
        self.top_frame = ttk.Frame(self.root)
        self.top_frame.pack(fill="x", padx=20, pady=(15, 5))
        
        self.lbl_folder = ttk.Label(self.top_frame, text="Select Folder to analyze:")
        self.lbl_folder.pack(side="left")
        
        self.btn_settings = ttk.Button(self.top_frame, text="⚙ Settings", width=10, command=self.open_settings)
        self.btn_settings.pack(side="right")
        
        # Folder input frame
        self.frame_folder = ttk.Frame(self.root)
        self.frame_folder.pack(fill="x", padx=20)
        
        self.entry_folder = ttk.Entry(self.frame_folder)
        self.entry_folder.pack(side="left", fill="x", expand=True)
        self.entry_folder.insert(0, "P:/Osnabrück 2023 Baulos 29/Abrechnung/Aufmaße LWL")
        
        self.btn_browse = ttk.Button(self.frame_folder, text="Browse...", command=self.browse_folder)
        self.btn_browse.pack(side="right", padx=(5, 0))
        
        # Action Buttons
        self.btn_analyze = ttk.Button(self.root, text="Start Analysis & Save to Excel", command=self.start_analysis)
        self.btn_analyze.pack(fill="x", padx=20, pady=(15, 5), ipady=5)
        
        # Open Folder link (initially hidden or disabled)
        self.btn_open_folder = ttk.Button(self.root, text="📁 Open Output Folder", command=self.open_output_folder, state="disabled")
        self.btn_open_folder.pack(fill="x", padx=20, pady=(0, 10))
        
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(self.root, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill="x", padx=20, pady=(0, 10))
        
        self.lbl_status = ttk.Label(self.root, text="Ready", foreground="gray")
        self.lbl_status.pack()

    def open_settings(self):
        # Create a Toplevel window for column selection
        settings_win = tk.Toplevel(self.root)
        settings_win.title("Output Column Settings")
        settings_win.geometry("300x320")
        settings_win.grab_set() # Focus lock to this window
        
        ttk.Label(settings_win, text="Select columns to include:", font=("", 10, "bold")).pack(pady=10, anchor="w", padx=20)
        
        for col in self.all_columns:
            chk = ttk.Checkbutton(settings_win, text=col, variable=self.column_vars[col])
            chk.pack(anchor="w", padx=30, pady=2)
            
        ttk.Button(settings_win, text="Save & Close", command=settings_win.destroy).pack(pady=15)

    def browse_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.entry_folder.delete(0, tk.END)
            self.entry_folder.insert(0, folder)

    def update_progress(self, status_text, current, total):
        def _update():
            self.lbl_status.config(text=status_text)
            if total > 0:
                self.progress_var.set((current / total) * 100)
        self.root.after(0, _update)

    def finish_analysis(self, success, result_msg, file_path):
        def _finish():
            self.btn_analyze.config(state="normal")
            self.progress_var.set(100 if success else 0)
            if success:
                self.current_output_path = file_path
                self.btn_open_folder.config(state="normal")
                self.lbl_status.config(text="Analysis complete!", foreground="green")
                messagebox.showinfo("Success", f"Data successfully saved to:\n{file_path}")
            else:
                self.current_output_path = None
                self.btn_open_folder.config(state="disabled")
                self.lbl_status.config(text="Error occurred.", foreground="red")
                messagebox.showerror("Error", f"An error occurred during analysis:\n{result_msg}\n\nCheck folder_analyzer_error.log for details.")
        self.root.after(0, _finish)

    def open_output_folder(self):
        if self.current_output_path and os.path.exists(self.current_output_path):
            folder_path = os.path.dirname(self.current_output_path)
            try:
                # Use Windows startfile to open explorer and select the file
                subprocess.Popen(f'explorer /select,"{os.path.normpath(self.current_output_path)}"')
            except Exception as e:
                logging.error("Failed to open folder", exc_info=True)
                messagebox.showerror("Error", "Could not open folder.")

    def start_analysis(self):
        folder = self.entry_folder.get().strip()
        if not folder or not os.path.exists(folder):
            messagebox.showerror("Error", "Please select a valid folder.")
            return
            
        selected_cols = [col for col in self.all_columns if self.column_vars[col].get()]
        if not selected_cols:
            messagebox.showerror("Error", "Please select at least one column in Settings.")
            return
            
        output_file = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            title="Save analysis as"
        )
        
        if not output_file: 
            return
            
        self.btn_analyze.config(state="disabled")
        self.btn_open_folder.config(state="disabled")
        self.progress_var.set(0)
        
        self.logic = FolderAnalyzerLogic(folder, output_file, selected_cols, self.update_progress, self.finish_analysis)
        
        thread = threading.Thread(target=self.logic.run)
        thread.daemon = True
        thread.start()

if __name__ == "__main__":
    root = tk.Tk()
    app = FolderAnalyzerApp(root)
    root.mainloop()
