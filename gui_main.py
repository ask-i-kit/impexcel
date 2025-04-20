# coding:utf-8
import tkinter as tk
import tkinter.font as font
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox
import cntl_excel
import os 

class Application(tk.Frame):
    def __init__(self, root):
        super().__init__(root)
        self.pack()
        TREE_EXCELFILES_COLUMN = ["ファイル名","更新日時"]
        FONT = font.Font(
            root,
            family = 'Yu Gothic UI',
            size=10
            )
        self.folder_name = tk.StringVar()
        self.folder_name.set("")
        
        frame_input_directory = tk.Frame(root)
        
        label_input_directory = ttk.Label(frame_input_directory,
                        text='フォルダ参照',
                        # bg='white',
                        font=FONT,
                        border=1,
                        padding=(5,2),
                        justify=tk.LEFT,
                        # relief='ridge',
                        )
        label_input_directory.pack(side=tk.LEFT)

        enrty_input_directory = ttk.Entry(frame_input_directory,
                        foreground='black',
                        font=FONT,
                        background='white',
                        width = 30,
                        justify=tk.LEFT,
                        textvariable=self.folder_name
                        )
        
        button_input_directory = ttk.Button(frame_input_directory,
                        text="参照",
                        padding=(5,2),
                        
                        # textvariable=some_StringVar,
                        )
        
        button_import_excelfiles = ttk.Button(frame_input_directory,text="読込",padding=(5,2))
        button_import_excelfiles.bind('<ButtonPress>',self.import_excels)
        enrty_input_directory.pack(side=tk.LEFT)
        button_input_directory.bind('<ButtonPress>', self.folder_dialog)
        button_input_directory.pack(side=tk.LEFT)
        frame_input_directory.pack(side=tk.TOP)
        button_import_excelfiles.pack(side=tk.LEFT)
        
        self.list_excelfiles = []
        self.listview_excelfile_var = tk.StringVar(value=self.list_excelfiles)
        
        
        self.treeview_excelfile = ttk.Treeview(root,
                                    # font=FONT,
                                    # width=50,
                                    # listvariable=self.listview_excelfile_var,
                                    columns=TREE_EXCELFILES_COLUMN
                                    )
        self.treeview_excelfile.column('#0',width=0,stretch='no')
        self.treeview_excelfile.column('ファイル名',anchor='w',width=150)
        self.treeview_excelfile.column('更新日時',anchor='w',width=50)
        self.treeview_excelfile.heading('#0',text='')
        self.treeview_excelfile.heading('ファイル名',anchor='w',text='ファイル名')
        self.treeview_excelfile.heading('更新日時',anchor='w',text='更新日時')
        
        
        self.treeview_excelfile.pack(side=tk.TOP, fill=tk.BOTH,padx=5,pady=2)
        

    def folder_dialog(self, event):
        iDir = os.path.abspath(os.path.dirname(__file__))
        folder_name = filedialog.askdirectory(initialdir=iDir)
        if len(folder_name) == 0:
            self.folder_name.set('')
        else:
            self.folder_name.set(folder_name)

    def import_excels(self, event):
        print("DEBUG: import_excels START")
        if self.folder_name.get() == "":
            messagebox.showwarning("注意","保存先フォルダを指定してください")
        else:
            self.list_excelfiles = cntl_excel.list_excel_files_with_timestamps(self.folder_name.get())
            if len(self.list_excelfiles) < 1:
                messagebox.showinfo("情報","指定されたフォルダに\nExcelファイルが存在しません")
        self.listview_excelfile_var.set(self.list_excelfiles)
        
        for i in self.treeview_excelfile.get_children():
            self.treeview_excelfile.delete(i)
        for excelfile in self.list_excelfiles:
            self.treeview_excelfile.insert(parent='',index='end',values=excelfile)
            
        print("DEBUG: import_excels END")
        
def main():
    root = tk.Tk()

    root.geometry('500x200')
    root.title('tkinter template')

    app = Application(root=root)
    app.mainloop()


if __name__ == "__main__":
    main()