import openpyxl
import tkinter as tk
from tkinter import ttk

root = tk.Tk()
root.title("Login History")
root.geometry("250x300")
root.resizable(False, False)


main_frame = tk.Frame(root)
main_frame.pack(fill=tk.BOTH, expand=1)


header_frame = tk.Frame(main_frame)
header_frame.pack(fill=tk.X)

canvas = tk.Canvas(main_frame)
canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)


scrollbar = tk.Scrollbar(main_frame, orient=tk.VERTICAL, command=canvas.yview)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)


canvas.configure(yscrollcommand=scrollbar.set)
canvas.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))


content_frame = tk.Frame(canvas)
canvas.create_window((0, 0), window=content_frame, anchor="nw")

def on_mouse_scroll(event):
    """Enables scrolling with the mouse wheel."""
    canvas.yview_scroll(-1 * (event.delta // 120), "units")

def create_gui():
    """Loads login history while keeping headers visible."""
    wb = openpyxl.load_workbook("Inventory.xlsx")
    ws = wb["LoginHistory"]

    
    headers = ["Username", "Timestamp"]
    for col, header in enumerate(headers):
        tk.Label(header_frame, text=header, borderwidth=2, relief="groove", width=15, bg="MistyRose3", fg="black").grid(row=0, column=col, padx=5, pady=5)

    for row_idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=1):
        for col_idx, value in enumerate(row):
            tk.Label(content_frame, text=value, borderwidth=2, relief="groove", width=15).grid(row=row_idx, column=col_idx, padx=5, pady=5)

    wb.close()

    # Update scroll region when data is loaded
    content_frame.update_idletasks()
    canvas.configure(scrollregion=canvas.bbox("all"))


canvas.bind_all("<MouseWheel>", on_mouse_scroll)

create_gui()
root.mainloop()
