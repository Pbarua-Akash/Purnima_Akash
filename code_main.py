import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
from excel_processor import process_excel  # Import the process_excel function

# Function to handle file selection and processing
def start_processing():
    # Disable the "Convert to CSV" button and show the "Please wait" message
    button_process.config(state=tk.DISABLED)
    label_status.config(text="Please wait, processing...", fg="blue")
    progress_bar.start(10)  # Start the progress bar
    root.update()  # Force update the GUI to reflect the changes

    try:
        # Get the input file path
        input_file = entry_input_file.get()
        if not input_file:
            messagebox.showerror("Error", "Please select an input Excel file!")
            return

        # Get the output directory
        output_dir = entry_output_dir.get()
        if not output_dir:
            messagebox.showerror("Error", "Please select an output directory!")
            return

        # Check if the input file exists
        if not os.path.isfile(input_file):
            messagebox.showerror("Error", "The input file does not exist!")
            return

        # Check if the output directory exists
        if not os.path.isdir(output_dir):
            messagebox.showerror("Error", "The output directory does not exist!")
            return

        # Process the Excel file
        process_excel(input_file, output_dir)
        messagebox.showinfo("Success", "CSV files exported successfully!")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")
    finally:
        # Re-enable the "Convert to CSV" button, clear the status message, and stop the progress bar
        button_process.config(state=tk.NORMAL)
        label_status.config(text="", fg="black")
        progress_bar.stop()

# Function to select the input file
def select_input_file():
    file_path = filedialog.askopenfilename(
        title="Select Input Excel File",
        filetypes=[("Excel Files", "*.xlsx *.xls")]
    )
    if file_path:
        entry_input_file.delete(0, tk.END)
        entry_input_file.insert(0, file_path)

# Function to select the output directory
def select_output_dir():
    dir_path = filedialog.askdirectory(title="Select Output Directory")
    if dir_path:
        entry_output_dir.delete(0, tk.END)
        entry_output_dir.insert(0, dir_path)

# Create the main window
root = tk.Tk()
root.title("Excel to CSV Converter")

# Input file selection
label_input_file = tk.Label(root, text="Input Excel File:")
label_input_file.grid(row=0, column=0, padx=10, pady=10)

entry_input_file = tk.Entry(root, width=50)
entry_input_file.grid(row=0, column=1, padx=10, pady=10)

button_input_file = tk.Button(root, text="Browse", command=select_input_file)
button_input_file.grid(row=0, column=2, padx=10, pady=10)

# Output directory selection
label_output_dir = tk.Label(root, text="Output Directory:")
label_output_dir.grid(row=1, column=0, padx=10, pady=10)

entry_output_dir = tk.Entry(root, width=50)
entry_output_dir.grid(row=1, column=1, padx=10, pady=10)

button_output_dir = tk.Button(root, text="Browse", command=select_output_dir)
button_output_dir.grid(row=1, column=2, padx=10, pady=10)

# Process button
button_process = tk.Button(root, text="Convert to CSV", command=start_processing)
button_process.grid(row=2, column=1, padx=10, pady=20)

# Status label
label_status = tk.Label(root, text="", fg="black")
label_status.grid(row=3, column=1, padx=10, pady=10)

# Progress bar
progress_bar = ttk.Progressbar(root, orient="horizontal", length=300, mode="indeterminate")
progress_bar.grid(row=4, column=1, padx=10, pady=10)

# Run the application
root.mainloop()