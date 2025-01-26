import tkinter as tk
from tkinter import filedialog, messagebox
import sys
import threading

class RedirectOutput:
    def __init__(self, text_widget):
        self.text_widget = text_widget

    def write(self, text):
        self.text_widget.insert(tk.END, text)
        self.text_widget.see(tk.END)  # Automatically scroll to the end

    def flush(self):
        pass

def browse_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
    if file_path:
        file_path_entry.delete(0, tk.END)  # Clear the entry field
        file_path_entry.insert(0, file_path)  # Insert the selected file path

def run_script():
    browser_choice = browser_var.get()
    file_path = file_path_entry.get()

    if not browser_choice:
        messagebox.showerror("Error", "Please select a browser (Chrome or Edge).")
        return

    if not file_path:
        messagebox.showerror("Error", "Please select an Employee Name Excel file.")
        return

    # Define the function to simulate your script
    def script_logic():
        print(f"Selected Browser: {browser_choice}")
        print(f"Selected File Path: {file_path}")

        # Simulate processing
        print("Processing file")
        # Your script logic goes here...
        print("Processing complete.")

    # Run the script logic in a separate thread to prevent UI freezing
    threading.Thread(target=script_logic).start()

# Create the main application window
root = tk.Tk()
root.title("Browser and File Selector")
root.geometry("600x400")
root.resizable(False, False)

# Create a frame for the file selection
file_frame = tk.Frame(root, padx=10, pady=10)
file_frame.pack(fill=tk.X)

browse_button = tk.Button(file_frame, text="Browse", command=browse_file)
browse_button.pack(side=tk.LEFT, padx=5)

file_path_entry = tk.Entry(file_frame, width=50)
file_path_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)

# Create a frame for browser selection
browser_frame = tk.Frame(root, padx=10, pady=10)
browser_frame.pack(fill=tk.X)

browser_var = tk.StringVar()

chrome_radio = tk.Radiobutton(browser_frame, text="Chrome", variable=browser_var, value="Chrome")
chrome_radio.pack(side=tk.LEFT, padx=10)

edge_radio = tk.Radiobutton(browser_frame, text="Edge", variable=browser_var, value="Edge")
edge_radio.pack(side=tk.LEFT, padx=10)

# Create a button to run the script
run_button = tk.Button(root, text="Run", command=run_script)
run_button.pack(pady=10)

# Create a log screen
log_frame = tk.Frame(root, padx=10, pady=10)
log_frame.pack(fill=tk.BOTH, expand=True)

log_label = tk.Label(log_frame, text="Log Screen", anchor="w")
log_label.pack(fill=tk.X)

log_text = tk.Text(log_frame, bg="black", fg="white", state=tk.NORMAL)
log_text.pack(fill=tk.BOTH, expand=True)

# Redirect print statements to the log screen
sys.stdout = RedirectOutput(log_text)

# Start the Tkinter event loop
root.mainloop()
