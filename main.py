import tkinter as tk
from docx2pdf import convert
from tkinter import filedialog,messagebox
import os


def pdfConverter(file_path):
    try:
        convert(file_path)
        pdf_path = os.path.splitext(file_path)[0] + ".pdf"
        print("✅ Your new PDF is stored in:", pdf_path)
        messagebox.showinfo("Success", f"PDF saved at:\n{pdf_path}")
    except Exception as e:
        messagebox.showerror("Error", str(e))
        print("❌ Conversion failed:", e)


def open_file_explorer():
    file_path = filedialog.askopenfilename(title="Select a DOCX file", filetypes=[("Word Documents", "*.docx")])
    if file_path:
        print("Selected file:", file_path)
        pdfConverter(file_path)


def main():
    root = tk.Tk()
    root.title("Niti Bossi Converter")
    root.geometry("300x120")

    btn = tk.Button(root, text="Choose Word File to Convert", command=open_file_explorer)
    btn.pack(pady=30)

    root.mainloop()


if __name__ == "__main__":
    main()