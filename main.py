"""
Main script for the Enrollment App.

This script provides a GUI for updating master Excel spreadsheets
with data from other Excel files in a selected folder.
"""

import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import time
import threading
import openpyxl


def browse_master_file():
    """Opens a file dialog to select the master spreadsheet."""
    filepath = filedialog.askopenfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
    )
    if filepath:
        master_file_entry.delete(0, tk.END)
        master_file_entry.insert(0, filepath)


def browse_folder():
    """Opens a file dialog to select the folder."""
    folder_path = filedialog.askdirectory()
    if folder_path:
        folder_entry.delete(0, tk.END)
        folder_entry.insert(0, folder_path)


def update_spreadsheet(master_file_path, new_file_path):
    """Updates the master spreadsheet with data from a new file."""
    try:
        master_wb = openpyxl.load_workbook(master_file_path)
        master_sheet = master_wb.active

        new_wb = openpyxl.load_workbook(new_file_path)
        new_sheet = new_wb.active

        headers = [cell.value for cell in master_sheet[1]]
        if 'CRN' not in headers:
            raise ValueError("CRN column not found in master file headers.")
        crn_index = headers.index('CRN')

        new_data = {row[crn_index]: row for row in new_sheet.iter_rows(min_row=2, values_only=True)}

        for row_idx, row in enumerate(master_sheet.iter_rows(min_row=2), start=2):
            crn = row[crn_index].value
            if crn in new_data:
                for col_idx, value in enumerate(new_data[crn], start=1):
                    master_sheet.cell(row=row_idx, column=col_idx).value = value

        master_wb.save(master_file_path)
        return True

    except FileNotFoundError as e:
        messagebox.showerror("Error", f"File not found: {e}")
    except openpyxl.utils.exceptions.InvalidFileException as e:
        messagebox.showerror("Error", f"Invalid Excel file: {e}")
    except ValueError as e:
        messagebox.showerror("Error", str(e))
    except (PermissionError, IOError) as e:
        messagebox.showerror("Error", f"File permission or IO error: {e}")
    except Exception as e:  # Fallback for unexpected exceptions (with justification)
        messagebox.showerror("Error", f"An unexpected error occurred: {e}")
        raise  # Optionally re-raise the exception for debugging
    return False


def update_progress(current_file, total_files):
    """Updates the progress bar and status label."""
    progress_bar['value'] = (current_file / total_files) * 100
    status_label.config(text=f"Processing file {current_file} of {total_files}")
    root.update_idletasks()


def start_update():
    """Starts the spreadsheet update process in a separate thread."""
    master_file_path = master_file_entry.get()
    folder_path = folder_entry.get()

    if not master_file_path or not folder_path:
        messagebox.showerror("Error", "Please select both the master file and the folder.")
        return

    def update_thread():
        try:
            files = [
                f for f in os.listdir(folder_path)
                if f.endswith('.xlsx') and f != os.path.basename(master_file_path)
            ]
            total_files = len(files)
            for i, filename in enumerate(files):
                new_file_path = os.path.join(folder_path, filename)
                if update_spreadsheet(master_file_path, new_file_path):
                    update_progress(i + 1, total_files)
                    time.sleep(1)
                else:
                    break
            status_label.config(text="Update complete!")
        except FileNotFoundError as e:
            messagebox.showerror("Error", f"Folder not found: {e}")
        except PermissionError as e:
            messagebox.showerror("Error", f"Permission denied: {e}")
        except OSError as e:
            messagebox.showerror("Error", f"OS error: {e}")
        except Exception as e:  # Fallback for unexpected exceptions (with justification)
            messagebox.showerror("Error", f"An unexpected error occurred: {e}")
            raise  # Optionally re-raise the exception for debugging

    threading.Thread(target=update_thread).start()


root = tk.Tk()
root.title("Spreadsheet Updater")

master_file_label = tk.Label(root, text="Master Spreadsheet:")
master_file_label.grid(row=0, column=0, padx=5, pady=5)
master_file_entry = tk.Entry(root, width=50)
master_file_entry.grid(row=0, column=1, padx=5, pady=5)
master_file_button = tk.Button(root, text="Browse", command=browse_master_file)
master_file_button.grid(row=0, column=2, padx=5, pady=5)

folder_label = tk.Label(root, text="Folder:")
folder_label.grid(row=1, column=0, padx=5, pady=5)
folder_entry = tk.Entry(root, width=50)
folder_entry.grid(row=1, column=1, padx=5, pady=5)
folder_button = tk.Button(root, text="Browse", command=browse_folder)
folder_button.grid(row=1, column=2, padx=5, pady=5)

update_button = tk.Button(root, text="Update Now", command=start_update)
update_button.grid(row=2, column=1, padx=5, pady=10)

progress_bar = ttk.Progressbar(root, orient="horizontal", length=300, mode="determinate")
progress_bar.grid(row=3, column=0, columnspan=3, padx=5, pady=5)

status_label = tk.Label(root, text="")
status_label.grid(row=4, column=0, columnspan=3, padx=5, pady=5)

root.mainloop()
