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
        
        
        # フォルダ名取得エリア
        self.folder_name = tk.StringVar()
        self.folder_name.set("C:\\TEMP")
        
        # 各要素の定義
        frame_input_directory = tk.Frame(root)
        
        label_input_directory = ttk.Label(frame_input_directory,
                        text='フォルダ参照',
                        # bg='white',
                        font=FONT,
                        border=1,
                        padding=(5,2),
                        justify=tk.LEFT,
                        )
        enrty_input_directory = ttk.Entry(frame_input_directory,
                        font=FONT,
                        width = 30,
                        justify=tk.LEFT,
                        textvariable=self.folder_name
                        )
        
        button_input_directory = ttk.Button(frame_input_directory,
                        text="参照",
                        padding=(5,2),
                        )
        
        button_import_excelfiles = ttk.Button(frame_input_directory,text="読込",padding=(5,2))
        
        # イベントバインド
        button_import_excelfiles.bind('<ButtonPress>',self.import_excels)
        button_input_directory.bind('<ButtonPress>', self.folder_dialog)
        
        # 各要素をパック
        label_input_directory.pack(side=tk.LEFT)
        enrty_input_directory.pack(side=tk.LEFT)
        button_input_directory.pack(side=tk.LEFT)
        button_import_excelfiles.pack(side=tk.LEFT)
        frame_input_directory.pack(side=tk.TOP)
        
        # 取得したファイル一覧を表示するエリア
        self.treeview_excelfile = ttk.Treeview(root,
                                    columns=TREE_EXCELFILES_COLUMN,
                                    show='headings'
                                    )
        self.treeview_excelfile.column('#0',width=0,stretch='no')
        self.treeview_excelfile.column('ファイル名',anchor='w',width=250)
        self.treeview_excelfile.column('更新日時',anchor='w',width=130)
        self.treeview_excelfile.heading('#0',text='')
        self.treeview_excelfile.heading('ファイル名', text='ファイル名',
            command=lambda: self.treeview_sort_column(self.treeview_excelfile, 'ファイル名', False))
        self.treeview_excelfile.heading('更新日時', text='更新日時',
            command=lambda: self.treeview_sort_column(self.treeview_excelfile, '更新日時', False))
        
        
        self.treeview_excelfile.bind("<Double-1>", self.open_excel_file_on_double_click)

        self.treeview_excelfile.pack(side=tk.TOP, fill=tk.BOTH,padx=5,pady=2)
        

    def folder_dialog(self, event):
        if self.folder_name.get() == "":
            initial_dir = "C:\\TEMP"
        else:
            initial_dir = self.folder_name.get()
        folder_name = filedialog.askdirectory(initialdir=initial_dir)
        if len(folder_name) == 0:
            self.folder_name.set(initial_dir)
        else:
            self.folder_name.set(folder_name)

    def import_excels(self, event):
        print("DEBUG: import_excels START")

        # 保存先フォルダが指定されていない場合
        if self.folder_name.get() == "":
            messagebox.showwarning("警告", "保存先フォルダを指定してください")
            return  # フォルダがない場合はここで処理を中断

        # Excelファイル一覧を取得
        self.list_excelfiles = cntl_excel.list_excel_files_with_timestamps(self.folder_name.get())

        if len(self.list_excelfiles) < 1:
            messagebox.showwarning("警告", "指定されたフォルダに\nExcelファイルが存在しません")
            # ツリービューをクリアして終了
            for i in self.treeview_excelfile.get_children():
                self.treeview_excelfile.delete(i)
            return

        # ツリービューの要素をすべて削除
        for i in self.treeview_excelfile.get_children():
            self.treeview_excelfile.delete(i)

        # 取得したExcelファイル情報をツリービューに追加
        for excelfile in self.list_excelfiles:
            # excelfile は (パス, タイムスタンプ) のタプル
            self.treeview_excelfile.insert(parent='', index='end', values=excelfile)

        print("DEBUG: import_excels END")
        
    def treeview_sort_column(self, tv, col, reverse):
    # 現在のデータを取得してソート
        data = [(tv.set(k, col), k) for k in tv.get_children('')]

        if col == '更新日時':
            from datetime import datetime
            data.sort(key=lambda t: datetime.strptime(t[0], "%Y-%m-%d %H:%M:%S"), reverse=reverse)
        else:
            data.sort(reverse=reverse)

        # 並び替え
        for index, (_, k) in enumerate(data):
            tv.move(k, '', index)

        # 次回クリック時は反対順に
        tv.heading(col, command=lambda: self.treeview_sort_column(tv, col, not reverse))
        
    def open_excel_file_on_double_click(self, event):
        selected_item = self.treeview_excelfile.focus()
        if not selected_item:
            return

        values = self.treeview_excelfile.item(selected_item, "values")
        if not values:
            return

        filepath = values[0]  # 最初のカラムがファイルパスの場合

        if os.path.exists(filepath):
            try:
                os.startfile(filepath)  # Windows専用。macOS/Linuxは別処理になる
            except Exception as e:
                messagebox.showerror("エラー", f"ファイルを開けませんでした:\n{e}")
        else:
            messagebox.showwarning("警告", "ファイルが存在しません。")
            
            
def main():
    root = tk.Tk()

    root.geometry('500x300')
    root.title('tkinter template')

    app = Application(root=root)
    app.mainloop()


if __name__ == "__main__":
    main()