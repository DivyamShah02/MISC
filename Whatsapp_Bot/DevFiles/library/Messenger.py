import tkinter as tk

def show_success_message():
    root = tk.Tk()
    root.title("Process Status")
    root.configure(bg='green')

    label = tk.Label(root, text="Process Completed!", bg='green', fg='white', font=('Helvetica', 16))
    label.pack(pady=40, padx=40)

    root.mainloop()

def show_danger_message():
    root = tk.Tk()
    root.title("Process Status")
    root.configure(bg='red')

    label = tk.Label(root, text="Bot FATA!!", bg='red', fg='white', font=('Helvetica', 16))
    label.pack(pady=40, padx=40)

    root.mainloop()
