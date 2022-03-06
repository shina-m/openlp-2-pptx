from tkinter import *
from tkinter.ttk import *
from tkinter.filedialog import askopenfilename, asksaveasfilename
import converter
import os, sys, subprocess


def generate_window():
    ws = Tk()
    ws.title('OpenLP Service File to PowerPoint Converter')
    ws.geometry('400x200')
    ws.columnconfigure(1, weight=1)   # Set weight to row and
    ws.rowconfigure(1, weight=0)      # column where the widget is

    osz_file_path = None
    def open_file():
        global osz_file_path
        file_path = askopenfilename(
            title="Select File...",
            filetypes=[('OpenLP Service Files', '*.osz')],
            defaultextension='.osz',
        )

        if file_path is not None:
            osz_file_path = file_path
            Label(ws, text=os.path.basename(file_path), foreground='green').grid(row=2, column=1, pady=(0,10))

    def upload_osz():
        global osz_file_path
        ws.update_idletasks()
        if osz_file_path:
            dir, filename = os.path.split(os.path.abspath(osz_file_path))
            ppt_file_path = asksaveasfilename(
                title="Select Location to Save Slide",
                initialfile=os.path.splitext(filename)[0] + ".pptx",
                initialdir=dir,
                filetypes=[('PowerPoint Files', '.pptx')],
                defaultextension='.pptx',
            )
            if ppt_file_path:
                err = converter.gui_endpoint(osz_file_path, ppt_file_path)
                if not err:
                    Label(ws, text="File Converted Successfully", foreground='green').grid(row=4, column=1, pady=10)
                    dir, filename = os.path.split(os.path.abspath(ppt_file_path))
                    open_ppt(dir, filename)
                else:
                    Label(ws, text=err, foreground='red').grid(row=6, column=1, pady=10)

    def open_ppt(dir, filename):
        if sys.platform == "win32":
            os.startfile(filename)
        else:
            opener = "open" if sys.platform == "darwin" else "xdg-open"
            subprocess.Popen([opener, filename], cwd=dir)

    osz_btn = Button(
        ws,
        text='Choose Service File',
        command=lambda: open_file()
    )
    osz_btn.grid(row=1, column=1, pady=(50,10))
    # osz_btn.place(relx=0.5, rely=0.5, anchor=CENTER)
    upld = Button(
        ws,
        text='Convert',
        command=upload_osz
    )
    upld.grid(row=3, column=1,pady=(10,0))
    # upld.place(relx=1.5, rely=1.5, anchor=CENTER)

    return ws
