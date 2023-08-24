import tkinter as tk
from tkinterdnd2 import DND_FILES, TkinterDnD
import os
from pdf2docx import Converter
from tkinter import messagebox

class DragDropWidget(tk.Frame):
    def __init__(self, master=None, **kwargs):
        super().__init__(master, **kwargs)
        self.create_widgets()

    def create_widgets(self):
        self.label = tk.Label(self, text='將PDF文件拖放到此處', font=('Arial', 12), width=30, height=10)
        self.label.pack(pady=50)

        self.label.drop_target_register(DND_FILES)
        self.label.dnd_bind('<<Drop>>', self.on_drop)

    def on_drop(self, event):
        file_path = event.data
        if file_path:
            output_dir = os.path.dirname(file_path)
            convert_pdf_to_word(file_path, output_dir)
            messagebox.showinfo('轉換完成', 'PDF 轉換為 Word 文檔完成！')

def convert_pdf_to_word(pdf_path, output_dir):
    converter = Converter(pdf_path)
    output_filename = os.path.splitext(os.path.basename(pdf_path))[0] + '.docx'
    output_path = os.path.join(output_dir, output_filename)
    converter.convert(output_path, start=0, end=None)
    converter.close()

root = TkinterDnD.Tk()
root.title('PDF 轉 Word 工具')

app = DragDropWidget(root)
app.pack()

root.mainloop()