#============================libraries start==================================================
import tkinter as tk
from tkinter import filedialog, messagebox
from pdf2docx import Converter
import customtkinter as ctk
#============================libraries end==================================================


#============================file selection function start==================================================
def select_pdf_file():
    file_path = filedialog.askopenfilename(
        filetypes=[("PDF files", "*.pdf")],
        title="Choose a PDF file"
    )
    pdf_entry.delete(0, tk.END)
    pdf_entry.insert(0, file_path)
#============================file selection function end==================================================



#============================file convert function start==================================================
def convert_to_docx():
    pdf_path = pdf_entry.get()
    if not pdf_path:
        messagebox.showerror("Error", "Please select a PDF file")
        return

    docx_path = filedialog.asksaveasfilename(
        defaultextension=".docx",
        filetypes=[("DOCX files", "*.docx")],
        title="Save DOCX as"
    )
    if not docx_path:
        return

    try:
        cv = Converter(pdf_path)
        cv.convert(docx_path, start=0, end=None)
        cv.close()
        messagebox.showinfo("Success", f"File converted successfully: {docx_path}")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to convert file: {e}")

#============================file convert function end==================================================


# Set up the GUI
root=ctk.CTk()
root.title("Philip Willson PDF to DOCX Converter")
root.geometry("600x600")
root.resizable(0,0)

#============================window apperance start==================================================
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("dark-blue")
#============================window apperance end==================================================


#============================backgroun image start==================================================
bg1=tk.PhotoImage(file="bg.png")
bg2=tk.Label(root,image=bg1,bg="gray").place(y=1,x=10)
#============================background image end==================================================



#============================logo start==================================================
logo1=tk.PhotoImage(file="logo.png")
logo2=tk.Label(root,image=logo1,background="white").pack(pady=30)
#============================logo end====================================================




#============================banner start==================================================
banner1=tk.PhotoImage(file="banner.png")
banner2=tk.Label(root,image=banner1,background="red").pack(pady=30)
#============================banner end====================================================

#============================button image start==================================================
img1=tk.PhotoImage(file="convert.png")
img2=tk.PhotoImage(file="file.png")
#============================button image end==================================================


#============================entry label start==================================================
label1=tk.Label(root, text="PDF File:",bg="red").place(y=576,x=170)
#============================entry label end==================================================


#============================entry start==================================================
pdf_entry = tk.Entry(root, width=50)
pdf_entry.pack(padx=10, pady=10)
#============================entry end==================================================


#============================file selection function start==================================================
btn1=ctk.CTkButton(root, text="Browse file", command=select_pdf_file,fg_color="green",image=img1).place(y=500,x=140)
btn2=ctk.CTkButton(root,text="Convert to DOCX", command=convert_to_docx,fg_color="red",image=img2).place(y=500,x=296)




root.mainloop()

