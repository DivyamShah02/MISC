import os
import psutil
import subprocess
import tkinter as tk
from tkinter import font, messagebox

def show_summary(success_count, error_count, error_file_path=None):
    """Display a summary window showing the counts of messages sent and errors."""
    root = tk.Tk()
    root.title("WhatsApp Bot Processing Summary")

    def open_error_file():
        """Open the error file with the default application and close the Tkinter window."""
        if error_file_path and os.path.exists(error_file_path):
            try:
                # Open the error file using the default application
                process = subprocess.Popen(['start', '', error_file_path], shell=True)
                # Close Tkinter window after 1 second
                root.after(1000, close_all_cmds_and_tkinter)
            except subprocess.CalledProcessError:
                messagebox.showerror("Error", "Failed to open the error file.")
                root.after(1000, close_all_cmds_and_tkinter)  # Close Tkinter window after 1 second
        else:
            messagebox.showerror("Error", "Error file not found.")
            root.after(1000, close_all_cmds_and_tkinter)  # Close Tkinter window after 1 second

    def close_all_cmds_and_tkinter():
        """Close all CMD processes and the Tkinter window."""
        for proc in psutil.process_iter(['pid', 'name']):
            try:
                # Check if the process name is 'cmd.exe'
                if proc.info['name'] == 'cmd.exe':
                    # Terminate the process
                    process = psutil.Process(proc.info['pid'])
                    process.terminate()
                    print(f"Terminated cmd.exe process with PID: {proc.info['pid']}")
            except (psutil.NoSuchProcess, psutil.AccessDenied) as e:
                print(f"Failed to terminate process with PID: {proc.info['pid']}. Error: {e}")

        root.destroy()

    def on_closing():
        """Handle the Tkinter window closing."""
        close_all_cmds_and_tkinter()

    # Define a font and color scheme
    title_font = font.Font(family="Segoe UI", size=16, weight="bold")
    text_font = font.Font(family="Segoe UI", size=14)
    bg_color = "#333333"  # Dark background color
    text_color = "#FFFFFF"  # Light text color
    button_color = "#0078D4"  # Accent color (blue)
    button_text_color = "#FFFFFF"

    # Center the window
    window_width = 600
    window_height = 350  # Increased height to accommodate all buttons
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x = (screen_width / 2) - (window_width / 2)
    y = (screen_height / 2) - (window_height / 2)
    root.geometry(f"{window_width}x{window_height}+{int(x)}+{int(y)}")
    root.configure(bg=bg_color)
    root.attributes("-topmost", True)

    # Title label
    title_label = tk.Label(root, text="WhatsApp Bot Processing Summary", font=title_font, bg=bg_color, fg=text_color)
    title_label.pack(pady=10)

    # Create a label to display the summary
    summary_message = (
        f"Messages Sent: {success_count}\n"
        f"Errors Encountered: {error_count}"
    )
    summary_label = tk.Label(root, text=summary_message, font=text_font, bg=bg_color, fg=text_color, padx=20, pady=20)
    summary_label.pack(pady=10)

    # Open Error Excel button
    if error_file_path:
        open_error_button = tk.Button(
            root, 
            text="Open Error Excel", 
            command=open_error_file, 
            bg='red', 
            fg=button_text_color, 
            font=text_font, 
            padx=10, 
            pady=5
        )
        open_error_button.pack(pady=10)

    # Close button
    close_button = tk.Button(root, text="Close", command=close_all_cmds_and_tkinter, bg=button_color, fg=button_text_color, font=text_font, padx=10, pady=10)
    close_button.pack(pady=10)

    # Handle the window closing event (cross button)
    root.protocol("WM_DELETE_WINDOW", on_closing)

    # Run the Tkinter main loop
    root.mainloop()
