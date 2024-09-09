import tkinter as tk
from tkinter import filedialog, messagebox, ttk, font
import pandas as pd
import os


class DuplicateDataIdentifierApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Duplicate Data Identifier")
        self.custom_font = font.Font(family="Helvetica", size=12)

        self.style = ttk.Style()
        self.style.configure('TLabel', font=('Helvetica', 12))
        self.style.configure('TButton', font=('Helvetica', 12), background="#333")
        self.style.configure('TEntry', font=('Helvetica', 12), background="#fff")

        self.show_main_menu()

    def show_main_menu(self):
        self.clear_window()
        ttk.Button(self.root, text="处理单个文件", command=self.show_single_file_interface).grid(column=1, row=0,
                                                                                                 padx=10, pady=10)
        ttk.Button(self.root, text="处理批量文件", command=self.show_batch_file_interface).grid(column=1, row=1,
                                                                                                padx=10, pady=10)

    def show_single_file_interface(self):
        self.clear_window()
        ttk.Button(self.root, text="返回", command=self.show_main_menu).grid(column=0, row=0, padx=10, pady=5,
                                                                             sticky='W')

        ttk.Button(self.root, text="选择主文件", command=lambda: self.select_file('primary')).grid(column=0, row=1,
                                                                                                   sticky='W', padx=10,
                                                                                                   pady=5)
        self.primary_file_path = tk.StringVar()
        self.primary_file_entry = ttk.Entry(self.root, textvariable=self.primary_file_path, font=self.custom_font,
                                            state='readonly', width=75)
        self.primary_file_entry.grid(column=1, row=1, padx=10, pady=5)

        ttk.Button(self.root, text="选择支持文件", command=lambda: self.select_file('supporting')).grid(column=0, row=2,
                                                                                                        sticky='W',
                                                                                                        padx=10, pady=5)
        self.supporting_file_path = tk.StringVar()
        self.supporting_file_entry = ttk.Entry(self.root, textvariable=self.supporting_file_path, font=self.custom_font,
                                               state='readonly', width=75)
        self.supporting_file_entry.grid(column=1, row=2, padx=10, pady=5)

        ttk.Button(self.root, text="选择输出文件夹路径", command=self.select_output_folder).grid(column=0, row=3,
                                                                                                 sticky='W', padx=10,
                                                                                                 pady=5)
        self.output_folder = tk.StringVar()
        self.output_folder_entry = ttk.Entry(self.root, textvariable=self.output_folder, font=self.custom_font,
                                             state='readonly', width=75)
        self.output_folder_entry.grid(column=1, row=3, padx=10, pady=5)

        ttk.Label(self.root, text="标题行数字:").grid(column=0, row=4, sticky='W', padx=10, pady=5)
        self.header_row = tk.IntVar(value=1)
        ttk.Entry(self.root, textvariable=self.header_row, font=self.custom_font, width=30).grid(column=1, row=4,
                                                                                                 padx=10, pady=5)

        ttk.Button(self.root, text="更新列名", command=self.update_identifier_options_single).grid(column=0, row=5,
                                                                                                   sticky='W', padx=10,
                                                                                                   pady=5)

        ttk.Label(self.root, text="唯一标识符列名:").grid(column=0, row=6, sticky='W', padx=10, pady=5)
        self.unique_identifier_combobox = ttk.Combobox(self.root, font=self.custom_font, state='readonly', width=28)
        self.unique_identifier_combobox.grid(column=1, row=6, padx=10, pady=5)

        ttk.Label(self.root, text="输出文件名前缀:").grid(column=0, row=7, sticky='W', padx=10, pady=5)
        self.output_filename = tk.StringVar()
        ttk.Entry(self.root, textvariable=self.output_filename, font=self.custom_font, width=30).grid(column=1, row=7,
                                                                                                      padx=10, pady=5)
        ttk.Label(self.root, text="(输出文件名↑，前缀_repeat.xlsx/前缀_non_repeat.xlsx，需符合excel命名规则)", font=self.custom_font).grid(column=1, row=8, padx=10, pady=0)

        ttk.Button(self.root, text="运行", command=self.process_single_file).grid(column=1, row=9, padx=10, pady=10)

    def show_batch_file_interface(self):
        self.clear_window()
        ttk.Button(self.root, text="返回", command=self.show_main_menu).grid(column=0, row=0, padx=10, pady=5,
                                                                             sticky='W')

        ttk.Button(self.root, text="选择主文件夹", command=lambda: self.select_folder_path('primary')).grid(column=0,
                                                                                                            row=1,
                                                                                                            sticky='W',
                                                                                                            padx=10,
                                                                                                            pady=5)
        primary_frame = ttk.Frame(self.root)
        primary_frame.grid(column=1, row=1, columnspan=3, sticky='EW', padx=10, pady=5)
        self.primary_folder_path = tk.StringVar()
        ttk.Entry(primary_frame, textvariable=self.primary_folder_path, font=self.custom_font, state='readonly',
                  width=75).pack(fill='x', expand=True)

        ttk.Button(self.root, text="选择支持文件夹", command=lambda: self.select_folder_path('supporting')).grid(
            column=0, row=2, sticky='W', padx=10, pady=5)
        supporting_frame = ttk.Frame(self.root)
        supporting_frame.grid(column=1, row=2, columnspan=3, sticky='EW', padx=10, pady=5)
        self.supporting_folder_path = tk.StringVar()
        ttk.Entry(supporting_frame, textvariable=self.supporting_folder_path, font=self.custom_font, state='readonly',
                  width=75).pack(fill='x', expand=True)

        ttk.Button(self.root, text="选择输出文件夹路径", command=self.select_output_folder).grid(column=0, row=3,
                                                                                                 sticky='W', padx=10,
                                                                                                 pady=5)
        output_frame = ttk.Frame(self.root)
        output_frame.grid(column=1, row=3, columnspan=3, sticky='EW', padx=10, pady=5)
        self.output_folder = tk.StringVar()
        ttk.Entry(output_frame, textvariable=self.output_folder, font=self.custom_font, state='readonly',
                  width=75).pack(fill='x', expand=True)

        ttk.Button(self.root, text="选择示例文件", command=lambda: self.select_file('sample')).grid(column=0, row=4,
                                                                                                    sticky='W', padx=10,
                                                                                                    pady=5)
        sample_frame = ttk.Frame(self.root)
        sample_frame.grid(column=1, row=4, columnspan=3, sticky='EW', padx=10, pady=5)
        self.sample_file_path = tk.StringVar()
        ttk.Entry(sample_frame, textvariable=self.sample_file_path, font=self.custom_font, state='readonly',
                  width=75).pack(fill='x', expand=True)

        ttk.Label(self.root, text="标题行数字:").grid(column=0, row=5, sticky='W', padx=10, pady=5)
        self.header_row = tk.IntVar(value=1)
        ttk.Entry(self.root, textvariable=self.header_row, font=self.custom_font).grid(column=1, row=5, padx=10, pady=5)

        ttk.Button(self.root, text="更新列名", command=self.update_identifier_options_batch).grid(column=0, row=6,
                                                                                                  sticky='W', padx=10,
                                                                                                  pady=5)

        ttk.Label(self.root, text="唯一标识符列名:").grid(column=0, row=7, sticky='W', padx=10, pady=5)
        self.unique_identifier_combobox = ttk.Combobox(self.root, font=self.custom_font, state='readonly', width=18)
        self.unique_identifier_combobox.grid(column=1, row=7, padx=10, pady=5)

        ttk.Label(self.root, text="文件名列表(以逗号分隔):").grid(column=0, row=8, sticky='W', padx=10, pady=5)
        self.file_list = tk.StringVar()
        ttk.Entry(self.root, textvariable=self.file_list, font=self.custom_font).grid(column=1, row=8, padx=10, pady=5)

        ttk.Label(self.root, text="重复文件名前缀:").grid(column=0, row=9, sticky='W', padx=10, pady=5)
        self.dup_prefix = tk.StringVar()
        ttk.Entry(self.root, textvariable=self.dup_prefix, font=self.custom_font).grid(column=1, row=9, padx=10, pady=5)

        ttk.Label(self.root, text="重复文件名后缀:").grid(column=2, row=9, sticky='W', padx=10, pady=5)
        self.dup_suffix = tk.StringVar()
        ttk.Entry(self.root, textvariable=self.dup_suffix, font=self.custom_font).grid(column=3, row=9, padx=10, pady=5)

        ttk.Label(self.root, text="非重复文件名前缀:").grid(column=0, row=10, sticky='W', padx=10, pady=5)
        self.nondup_prefix = tk.StringVar()
        ttk.Entry(self.root, textvariable=self.nondup_prefix, font=self.custom_font).grid(column=1, row=10, padx=10,
                                                                                          pady=5)

        ttk.Label(self.root, text="非重复文件名后缀:").grid(column=2, row=10, sticky='W', padx=10, pady=5)
        self.nondup_suffix = tk.StringVar()
        ttk.Entry(self.root, textvariable=self.nondup_suffix, font=self.custom_font).grid(column=3, row=10, padx=10,
                                                                                          pady=5)

        ttk.Label(self.root, text="重复数据汇总文件名:").grid(column=0, row=11, sticky='W', padx=10, pady=5)
        self.dup_summary = tk.StringVar()
        ttk.Entry(self.root, textvariable=self.dup_summary, font=self.custom_font).grid(column=1, row=11, padx=10,
                                                                                        pady=5)

        ttk.Label(self.root, text="非重复数据汇总文件名:").grid(column=0, row=12, sticky='W', padx=10, pady=5)
        self.nondup_summary = tk.StringVar()
        ttk.Entry(self.root, textvariable=self.nondup_summary, font=self.custom_font).grid(column=1, row=12, padx=10,
                                                                                           pady=5)

        ttk.Button(self.root, text="运行", command=self.process_batch_files).grid(column=1, row=13, padx=10, pady=10)

    def clear_window(self):
        for widget in self.root.winfo_children():
            widget.destroy()

    def select_file(self, file_type):
        file_selected = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_type == 'primary':
            self.primary_file_path.set(file_selected)
            self.root.update_idletasks()  # 更新GUI
            self.primary_file_entry.xview_moveto(1)  # 移动到结尾
        elif file_type == 'supporting':
            self.supporting_file_path.set(file_selected)
            self.root.update_idletasks()  # 更新GUI
            self.supporting_file_entry.xview_moveto(1)  # 移动到结尾
        elif file_type == 'sample':
            self.sample_file_path.set(file_selected)
            self.root.update_idletasks()  # 更新GUI
            self.sample_file_entry.xview_moveto(1)  # 移动到结尾
        self.update_identifier_options_single()

    def select_folder_path(self, folder_type):
        folder_selected = filedialog.askdirectory()
        if folder_type == 'primary':
            self.primary_folder_path.set(folder_selected)
        elif folder_type == 'supporting':
            self.supporting_folder_path.set(folder_selected)

    def select_output_folder(self):
        folder_selected = filedialog.askdirectory()
        self.output_folder.set(folder_selected)

    def update_identifier_options_single(self):
        primary_path = self.primary_file_path.get()
        if not primary_path:
            messagebox.showerror("错误", "请先选择一个主文件。")
            return

        columns_set = self.extract_columns_from_file(primary_path)
        self.unique_identifier_combobox['values'] = list(columns_set)
        if columns_set:
            self.unique_identifier_combobox.set(next(iter(columns_set)))

    def update_identifier_options_batch(self):
        sample_path = self.sample_file_path.get()
        if not sample_path:
            messagebox.showerror("错误", "请先选择一个示例文件。")
            return

        columns_set = self.extract_columns_from_file(sample_path)
        self.unique_identifier_combobox['values'] = list(columns_set)
        if columns_set:
            self.unique_identifier_combobox.set(next(iter(columns_set)))

    def extract_columns_from_file(self, file_path):
        columns_set = set()
        try:
            header_row_index = self.header_row.get() - 1
            if header_row_index < 0:
                messagebox.showerror("错误", "标题行不能为负。请检查您的输入。")
                return columns_set
            df_example = pd.read_excel(file_path, header=header_row_index)
            columns = df_example.columns.tolist()
            columns_set.update(columns)
        except Exception as e:
            messagebox.showerror("错误", f"读取Excel列名时发生错误: {e}")
        return columns_set

    def process_single_file(self):
        if not all([self.primary_file_path.get(), self.supporting_file_path.get(), self.output_folder.get(),
                    self.output_filename.get()]):
            messagebox.showerror("错误", "请确保所有字段都已填写。")
            return

        try:
            df_primary = pd.read_excel(self.primary_file_path.get(), header=self.header_row.get() - 1)
            df_supporting = pd.read_excel(self.supporting_file_path.get(), header=self.header_row.get() - 1)

            unique_identifier = self.unique_identifier_combobox.get()
            repeat_students = df_primary[df_primary[unique_identifier].isin(df_supporting[unique_identifier])]
            non_repeat_students = df_primary[~df_primary[unique_identifier].isin(df_supporting[unique_identifier])]

            # Add file name column
            repeat_students['file'] = os.path.basename(self.primary_file_path.get())
            non_repeat_students['file'] = os.path.basename(self.primary_file_path.get())

            output_folder = self.output_folder.get()
            repeat_students_file = os.path.join(output_folder, f"{self.output_filename.get()}_repeat.xlsx")
            non_repeat_students_file = os.path.join(output_folder, f"{self.output_filename.get()}_non_repeat.xlsx")

            repeat_students.to_excel(repeat_students_file, index=False)
            non_repeat_students.to_excel(non_repeat_students_file, index=False)

            messagebox.showinfo("成功", "处理完成！请检查输出文件夹。")
        except Exception as e:
            messagebox.showerror("错误", f"处理过程中出现错误：{e}")

    def process_batch_files(self):
        if not all([self.primary_folder_path.get(), self.supporting_folder_path.get(), self.output_folder.get(),
                    self.sample_file_path.get(), self.file_list.get(), self.dup_summary.get(),
                    self.nondup_summary.get()]):
            messagebox.showerror("错误", "请确保所有字段都已填写。")
            return

        try:
            files = self.file_list.get().split(',')
            dup_summary_list = []
            nondup_summary_list = []

            for file in files:
                file = file.strip()
                if not file:
                    continue

                path_primary = os.path.join(self.primary_folder_path.get(), f'{file}.xlsx')
                path_supporting = os.path.join(self.supporting_folder_path.get(), f'{file}.xlsx')

                df_primary = pd.read_excel(path_primary, header=self.header_row.get() - 1)
                df_supporting = pd.read_excel(path_supporting, header=self.header_row.get() - 1)

                unique_identifier = self.unique_identifier_combobox.get()
                repeat_students = df_primary[df_primary[unique_identifier].isin(df_supporting[unique_identifier])]
                non_repeat_students = df_primary[~df_primary[unique_identifier].isin(df_supporting[unique_identifier])]

                # Add file name column
                repeat_students['file'] = file
                non_repeat_students['file'] = file

                output_folder = self.output_folder.get()
                repeat_prefix = self.dup_prefix.get() if self.dup_prefix.get() else ""
                repeat_suffix = self.dup_suffix.get() if self.dup_suffix.get() else ""
                non_repeat_prefix = self.nondup_prefix.get() if self.nondup_prefix.get() else ""
                non_repeat_suffix = self.nondup_suffix.get() if self.nondup_suffix.get() else ""

                repeat_students_file = os.path.join(output_folder, f"{repeat_prefix}{file}{repeat_suffix}.xlsx")
                non_repeat_students_file = os.path.join(output_folder,
                                                        f"{non_repeat_prefix}{file}{non_repeat_suffix}.xlsx")

                repeat_students.to_excel(repeat_students_file, index=False)
                non_repeat_students.to_excel(non_repeat_students_file, index=False)

                dup_summary_list.append(repeat_students)
                nondup_summary_list.append(non_repeat_students)

            # Merge all duplicates and non-duplicates into summary files
            if dup_summary_list:
                dup_summary_df = pd.concat(dup_summary_list, ignore_index=True)
                dup_summary_file = os.path.join(self.output_folder.get(), f"{self.dup_summary.get()}.xlsx")
                dup_summary_df.to_excel(dup_summary_file, index=False)

            if nondup_summary_list:
                nondup_summary_df = pd.concat(nondup_summary_list, ignore_index=True)
                nondup_summary_file = os.path.join(self.output_folder.get(), f"{self.nondup_summary.get()}.xlsx")
                nondup_summary_df.to_excel(nondup_summary_file, index=False)

            messagebox.showinfo("成功", "处理完成！请检查输出文件夹。")
        except Exception as e:
            messagebox.showerror("错误", f"处理过程中出现错误：{e}")


if __name__ == "__main__":
    root = tk.Tk()
    app = DuplicateDataIdentifierApp(root)
    root.geometry("900x600")  # Adjust window size as needed
    root.mainloop()
