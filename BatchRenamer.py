import os
import tkinter as tk
import ttkbootstrap as ttk
from tkinter import filedialog, messagebox
from datetime import datetime


class BulkFileRenamer:
    """
       A GUI application for bulk renaming files in a selected folder using various methods.

       Attributes:
           master (tk.Tk): The main window of the application.
           folder_label (ttk.Label): Label displaying the selected folder path.
           select_folder_button (ttk.Button): Button to open the folder selection dialog.
           list_frame (ttk.Frame): Frame containing the Treeview for displaying files.
           file_treeview (ttk.Treeview): Treeview for displaying the list of files in the selected folder.
           sort_column (str): Column name by which the file list is currently sorted.
           scrollbar (ttk.Scrollbar): Scrollbar for the Treeview.
           select_all_button (ttk.Button): Button to select all files in the Treeview.
           deselect_all_button (ttk.Button): Button to deselect all files in the Treeview.
           replace_section_label (ttk.Label): Label for the replace section.
           replace_label (ttk.Label): Label for the replace field.
           replace_entry (ttk.Entry): Entry for the text to be replaced.
           new_label (ttk.Label): Label for the new text field.
           new_entry (ttk.Entry): Entry for the new text.
           rename_button (ttk.Button): Button to apply the replace rename operation.
           case_rename_button (ttk.Button): Button to apply the case conversion rename operation.
           titlecase_radio (ttk.Radiobutton): Radiobutton to select title case conversion.
           uppercase_radio (ttk.Radiobutton): Radiobutton to select upper case conversion.
           lowercase_radio (ttk.Radiobutton): Radiobutton to select lower case conversion.
           case_label (ttk.Label): Label for the case conversion section.
           case_conversion_var (tk.StringVar): Variable to store the selected case conversion option.
           none_radio (ttk.Radiobutton): Radiobutton to select no case conversion.
           prefix_suffix_label (ttk.Label): Label for the prefix/suffix section.
           suffix_label (ttk.Label): Label for the suffix field.
           prefix_entry (ttk.Entry): Entry for the prefix text.
           suffix_entry (ttk.Entry): Entry for the suffix text.
           prefix_label (ttk.Label): Label for the prefix field.
           prefix_suffix_button (ttk.Button): Button to apply the prefix/suffix rename operation.
           sort_ascending (bool): Boolean indicating if the file list is sorted in ascending order.
           selected_files (list): List of selected files for renaming.
       """

    def __init__(self, master):
        """
        Initialize the BulkFileRenamer class.

        Args:
            master (tk.Tk): The root window or parent frame.
        """
        self.master = master
        self.master.title("Bulk File Renamer")
        self.folder_label = None
        self.select_folder_button = None
        self.list_frame = None
        self.file_treeview = None
        self.sort_column = None
        self.scrollbar = None
        self.select_all_button = None
        self.deselect_all_button = None
        self.replace_section_label = None
        self.replace_label = None
        self.replace_entry = None
        self.new_label = None
        self.new_entry = None
        self.rename_button = None
        self.case_rename_button = None
        self.titlecase_radio = None
        self.uppercase_radio = None
        self.lowercase_radio = None
        self.case_label = None
        self.case_conversion_var = None
        self.none_radio = None
        self.prefix_suffix_label = None
        self.suffix_label = None
        self.prefix_entry = None
        self.suffix_entry = None
        self.prefix_label = None
        self.prefix_suffix_button = None
        self.sort_ascending = None
        self.selected_files = []

        self.app_interface()

    def app_interface(self):
        """
        Set up the user interface for the application.

        This includes creating buttons, labels, entries, and the file list Treeview.
        """
        # Folder Selection
        self.select_folder_button = ttk.Button(self.master, text="Open Folder", command=self.select_folder,
                                               bootstyle="success")
        self.select_folder_button.grid(row=0, column=1, columnspan=2, padx=10, pady=10, sticky="ew")

        self.folder_label = ttk.Label(self.master, text="Selected Folder: ")
        self.folder_label.grid(row=1, column=0, columnspan=5, padx=10, pady=10, sticky="w")

        # File List
        self.list_frame = ttk.Frame(self.master)
        self.list_frame.grid(row=2, column=0, columnspan=5, padx=10, pady=10, sticky="nsew")

        self.file_treeview = ttk.Treeview(self.list_frame, columns=("Name", "Preview", "Date Modified"),
                                          show="headings")
        self.file_treeview.heading("Name", text="Name", command=lambda: self.on_column_click("Name"))
        self.file_treeview.heading("Preview", text="Preview")
        self.file_treeview.heading("Date Modified", text="Date Modified",
                                   command=lambda: self.on_column_click("Date Modified"))
        self.file_treeview.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.scrollbar = ttk.Scrollbar(self.list_frame, orient=tk.VERTICAL, command=self.file_treeview.yview,
                                       bootstyle="round")
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.file_treeview.configure(yscrollcommand=self.scrollbar.set)

        # Mouse select
        self.file_treeview.bind("<Control-Button-1>", self.on_mouse_press)
        self.file_treeview.bind("<ButtonPress-1>", self.on_mouse_press)
        self.file_treeview.bind("<B1-Motion>", self.on_mouse_drag)
        self.file_treeview.bind("<<TreeviewSelect>>", self.update_preview_on_selection_change)

        # Select all / Deselect all
        self.select_all_button = ttk.Button(self.master, text="Select All", command=self.select_all,
                                            bootstyle="secondary")
        self.select_all_button.grid(row=3, column=0, columnspan=2, padx=(15, 10), pady=10, sticky="ew")

        self.deselect_all_button = ttk.Button(self.master, text="Deselect All", command=self.deselect_all,
                                              bootstyle="secondary")
        self.deselect_all_button.grid(row=3, column=2, columnspan=3, padx=(10, 15), pady=10, sticky="ew")

        ttk.Separator(self.master, orient='horizontal').grid(row=4, column=0, columnspan=5, padx=10, pady=10,
                                                             sticky="ew")
        # Replace Section
        self.replace_section_label = ttk.Label(self.master, text="Replace character(s)")
        self.replace_section_label.grid(row=5, column=0, padx=10, sticky="w")
        self.replace_label = ttk.Label(self.master, text="Replace:")
        self.replace_label.grid(row=6, column=0, padx=10, sticky="w")
        self.replace_entry = ttk.Entry(self.master)
        self.replace_entry.grid(row=6, column=1, columnspan=2, pady=5, sticky="ew")
        self.replace_entry.bind("<FocusIn>", self.save_file_selection)
        self.replace_entry.bind("<FocusOut>", self.restore_file_selection)
        self.replace_entry.bind("<KeyRelease>", self.update_preview)

        self.new_label = ttk.Label(self.master, text="With:")
        self.new_label.grid(row=7, column=0, padx=10, sticky="w")
        self.new_entry = ttk.Entry(self.master)
        self.new_entry.grid(row=7, column=1, columnspan=2, pady=5, sticky="ew")
        self.new_entry.bind("<FocusIn>", self.save_file_selection)
        self.new_entry.bind("<FocusOut>", self.restore_file_selection)
        self.new_entry.bind("<KeyRelease>", self.update_preview)

        self.rename_button = ttk.Button(self.master, text="Replace character(s)", command=self.rename_files,
                                        bootstyle="primary")
        self.rename_button.grid(row=6, column=3, rowspan=2, padx=5)

        ttk.Separator(self.master, orient='horizontal').grid(row=8, column=0, columnspan=5, padx=10, pady=10,
                                                             sticky="ew")

        # Case Conversion Section
        self.case_label = ttk.Label(self.master, text="Case Conversion:")
        self.case_label.grid(row=10, column=0, rowspan=2, padx=10, sticky="w")
        self.case_conversion_var = tk.StringVar(value="none")
        self.none_radio = ttk.Radiobutton(self.master, text="None", variable=self.case_conversion_var,
                                          value="none", command=self.update_preview)
        self.none_radio.grid(row=10, column=1, pady=5, sticky="w")
        self.lowercase_radio = ttk.Radiobutton(self.master, text="Lowercase", variable=self.case_conversion_var,
                                               value="lowercase", command=self.update_preview)
        self.lowercase_radio.grid(row=10, column=2, pady=5, sticky="w")
        self.uppercase_radio = ttk.Radiobutton(self.master, text="Uppercase", variable=self.case_conversion_var,
                                               value="uppercase", command=self.update_preview)
        self.uppercase_radio.grid(row=11, column=1, pady=5, sticky="w")
        self.titlecase_radio = ttk.Radiobutton(self.master, text="Titlecase", variable=self.case_conversion_var,
                                               value="titlecase", command=self.update_preview)
        self.titlecase_radio.grid(row=11, column=2, pady=5, sticky="w")
        self.case_rename_button = ttk.Button(self.master, text="Change Case", command=self.rename_case)
        self.case_rename_button.grid(row=10, rowspan=2, column=3, sticky="w")

        # Prefix/Suffix Section
        ttk.Separator(self.master, orient='horizontal').grid(row=12, column=0, columnspan=5, padx=10, pady=10,
                                                             sticky="ew")
        self.prefix_suffix_label = ttk.Label(self.master, text="Add Prefix/Suffix:")
        self.prefix_suffix_label.grid(row=13, column=0, padx=10, sticky="w")

        self.prefix_label = ttk.Label(self.master, text="Prefix:")
        self.prefix_label.grid(row=14, column=0, padx=10, pady=5, sticky="w")
        self.prefix_entry = ttk.Entry(self.master)
        self.prefix_entry.grid(row=14, column=1, pady=5, sticky="ew")
        self.prefix_entry.bind("<KeyRelease>", self.update_preview)

        self.suffix_label = ttk.Label(self.master, text="Suffix:")
        self.suffix_label.grid(row=15, column=0, padx=10, pady=5, sticky="w")
        self.suffix_entry = ttk.Entry(self.master)
        self.suffix_entry.grid(row=15, column=1, pady=5, sticky="ew")
        self.suffix_entry.bind("<KeyRelease>", self.update_preview)

        self.prefix_suffix_button = ttk.Button(self.master, text="Rename Prefix/Suffix",
                                               command=self.rename_prefix_suffix)
        self.prefix_suffix_button.grid(row=14, column=2, rowspan=2, pady=5)

        self.sort_column = None
        self.sort_ascending = True

    def select_folder(self):
        """
        Open a dialog to select a folder and list its files in the Treeview.
        """
        folder_path = filedialog.askdirectory()
        if folder_path:
            self.folder_label.config(text="Selected Folder: " + folder_path)
            self.load_file_list(folder_path)

    def load_file_list(self, folder_path):
        """
       List all files in the selected folder in the Treeview.

       Args:
           folder_path (str): The path to the selected folder.
       """
        self.file_treeview.delete(*self.file_treeview.get_children())
        files = os.listdir(folder_path)
        for file_name in files:
            file_path = os.path.join(folder_path, file_name)
            if os.path.isfile(file_path):
                last_modified = os.path.getmtime(file_path)
                last_modified_date = datetime.fromtimestamp(last_modified).strftime('%Y-%m-%d %H:%M:%S')
                self.file_treeview.insert("", tk.END, values=(file_name, "", last_modified_date))

    def get_file_info(self, item):
        """
        Fetch file information from Treeview item.

        Args:
            item: The Treeview item.

        Returns:
            original_file_name (str): The original file name.
            original_file_path (str): The full file path.
            name_part (str): The name part of the file.
            extension_part (str): The extension part of the file.
        """
        file_info = self.file_treeview.item(item, 'values')
        original_file_name = file_info[0]
        original_file_path = os.path.join(self.folder_label.cget("text").replace("Selected Folder: ", ""),
                                          original_file_name)
        name_part, extension_part = os.path.splitext(original_file_name)

        return original_file_name, original_file_path, name_part, extension_part

    def sort_column(self, column, reverse):
        data = [(self.file_treeview.set(item, column), item) for item in self.file_treeview.get_children()]
        data.sort(reverse=reverse)

        for index, (value, item) in enumerate(data):
            self.file_treeview.move(item, "", index)

    def on_column_click(self, column):
        """
        Handle the column header click for sorting the files.

        Args:
            column (str): The column name by which to sort.
        """
        if self.sort_column == column:
            self.sort_ascending = not self.sort_ascending
        else:
            self.sort_column = column
            self.sort_ascending = True

        items = [(self.file_treeview.set(k, column), k) for k in self.file_treeview.get_children('')]

        if column == "Name":
            items.sort(key=lambda t: t[0].lower(), reverse=not self.sort_ascending)
        elif column == "Date Modified":
            items.sort(key=lambda t: datetime.strptime(t[0], '%Y-%m-%d %H:%M:%S'), reverse=not self.sort_ascending)

        for index, (val, k) in enumerate(items):
            self.file_treeview.move(k, '', index)

        self.file_treeview.heading(column, text=column + (" ↑" if self.sort_ascending else " ↓"))

    def select_all(self):
        """
        Select all files in the Treeview.
        """
        self.file_treeview.selection_set(self.file_treeview.get_children())

    def deselect_all(self):
        """
        Deselect all files in the Treeview.
        """
        self.file_treeview.selection_remove(self.file_treeview.get_children())

    def save_file_selection(self, event):
        self.selected_files = self.file_treeview.selection()

    def restore_file_selection(self, event):
        for item in self.selected_files:
            if self.file_treeview.exists(item):
                self.file_treeview.selection_add(item)

    def on_mouse_press(self, event):
        """
        Handle mouse press events to select or deselect files.

        Args:
            event (tk.Event): The mouse press event.
        """
        ctrl_pressed = (event.state & 0x4) != 0
        shift_pressed = (event.state & 0x1) != 0

        item = self.file_treeview.identify_row(event.y)
        if item:
            if ctrl_pressed:
                if item in self.file_treeview.selection():
                    self.file_treeview.selection_remove(item)
                else:
                    self.file_treeview.selection_add(item)
            elif shift_pressed:
                selection = self.file_treeview.selection()
                if selection:
                    last_selected = selection[-1]
                    start_idx = self.file_treeview.index(last_selected)
                    end_idx = self.file_treeview.index(item)
                    if start_idx < end_idx:
                        items_to_select = self.file_treeview.get_children()[start_idx:end_idx + 1]
                    else:
                        items_to_select = self.file_treeview.get_children()[end_idx:start_idx + 1]
                    self.file_treeview.selection_set(items_to_select)
                else:
                    self.file_treeview.selection_set(item)
            else:
                self.file_treeview.selection_set(item)
        else:
            self.file_treeview.selection_remove(self.file_treeview.selection())
        self.master.config(cursor="")

    def on_mouse_drag(self, event):
        item = self.file_treeview.identify_row(event.y)
        if item:
            self.file_treeview.selection_set(self.file_treeview.selection() + (item,))

    @staticmethod
    def check_invalid_characters(text):
        invalid_characters = ["<", ">", ":", "\"", "/", "\\", "|", "?", "*"]
        for character in invalid_characters:
            if character in text:
                return character

    def rename_files(self):
        """
        Perform the renaming of the selected files based on the user's input.
        """
        replace_text = self.replace_entry.get()
        new_text = self.new_entry.get()
        folder_path = self.folder_label.cget("text").replace("Selected Folder: ", "")

        if not self.file_treeview.selection():
            messagebox.showwarning("Nothing selected!", "Please select at least one file.")
            return

        invalid_char = self.check_invalid_characters(new_text)
        if invalid_char:
            messagebox.showwarning("Invalid character", f"You can't use '{invalid_char}' in the file name")
            return

        matched_files = []
        existing_files = set(os.listdir(folder_path))

        for item in self.file_treeview.selection():
            original_file_name = self.file_treeview.item(item, "values")[0]
            name_part, extension_part = os.path.splitext(original_file_name)

            if replace_text in name_part:
                matched_files.append(original_file_name)

        if not matched_files:
            if len(self.file_treeview.selection()) == 1:
                messagebox.showwarning("Not found", "No file name matches your input.")
            else:
                messagebox.showwarning("Not found", "No file names match your input. Nothing was renamed.")
            return

        renamed_count = 0
        for original_file_name in matched_files:
            name_part, extension_part = os.path.splitext(original_file_name)
            new_file_name = name_part.replace(replace_text, new_text) + extension_part

            if new_file_name in existing_files:
                messagebox.showwarning("Name already exists",
                                       f"A file named '{new_file_name}' already exists. Nothing was renamed.")
                return

            original_file_path = os.path.join(folder_path, original_file_name)
            new_file_path = os.path.join(folder_path, new_file_name)
            os.rename(original_file_path, new_file_path)
            renamed_count += 1

            existing_files.remove(original_file_name)
            existing_files.add(new_file_name)

        self.load_file_list(folder_path)
        messagebox.showinfo("Success",
                            f"{renamed_count} file{'s' if renamed_count > 1 else ''} ha{'ve' if renamed_count > 1 else 's'} been renamed successfully.")

    def rename_case(self):
        """
        Apply the selected case conversion to the selected files' names.
        """
        case_conversion = self.case_conversion_var.get()
        if not case_conversion:
            messagebox.showwarning("Input Error", "Please select a case conversion option.")
            return

        selected_items = self.file_treeview.selection()
        if not selected_items:
            messagebox.showwarning("Selection Error", "No files selected.")
            return

        folder_path = self.folder_label.cget("text").replace("Selected Folder: ", "")

        renamed_count = 0
        for item in selected_items:
            original_file_name, original_file_path, name_part, extension_part = self.get_file_info(item)
            new_name_part = name_part

            if case_conversion == "lowercase":
                new_name_part = name_part.lower()
            elif case_conversion == "uppercase":
                new_name_part = name_part.upper()
            elif case_conversion == "titlecase":
                new_name_part = name_part.title()
            elif case_conversion == "none":
                new_name_part = name_part

            new_file_name = f"{new_name_part}{extension_part}"
            new_file_path = os.path.join(folder_path, new_file_name)

            try:
                os.rename(original_file_path, new_file_path)
                renamed_count += 1
            except OSError as e:
                messagebox.showerror("Error", f"Failed to rename {original_file_name} to {new_file_name}. Error: {e}")
                return

        self.load_file_list(folder_path)
        messagebox.showinfo("Success",
                            f"{renamed_count} file{'s' if renamed_count > 1 else ''} ha{'ve' if renamed_count > 1 else 's'} been renamed successfully.")

    def rename_prefix_suffix(self):
        """
        Add specified prefix and/or suffix to the selected files' names and rename them.
        """
        folder_path = self.folder_label.cget("text").replace("Selected Folder: ", "")
        if not self.file_treeview.selection():
            messagebox.showwarning("Nothing selected", "Please select at least one file.")
            return

        prefix_text = self.prefix_entry.get()
        suffix_text = self.suffix_entry.get()
        renamed_count = 0

        for item in self.file_treeview.selection():
            original_file_name, original_file_path, name_part, extension_part = self.get_file_info(item)

            new_file_name = f"{prefix_text}{name_part}{suffix_text}{extension_part}"
            new_file_path = os.path.join(folder_path, new_file_name)

            try:
                os.rename(original_file_path, new_file_path)
                renamed_count += 1
            except OSError as e:
                messagebox.showerror("Error", f"Failed to rename {original_file_name} to {new_file_name}. Error: {e}")
                return

        self.load_file_list(folder_path)
        messagebox.showinfo("Success",
                            f"{renamed_count} file{'s' if renamed_count > 1 else ''} ha{'ve' if renamed_count > 1 else 's'} been renamed successfully.")

    def update_preview(self, event=None):
        """
        Update the preview of the renamed files in the Treeview based on the user's input.

        Parameters:
            event (tk.Event, optional): The event that triggered the update. Defaults to None.
        """
        replace_text = self.replace_entry.get()
        new_text = self.new_entry.get()
        case_conversion = self.case_conversion_var.get()
        prefix = self.prefix_entry.get()
        suffix = self.suffix_entry.get()

        selected_items = self.file_treeview.selection()

        for item in selected_items:
            file_name = self.file_treeview.item(item, 'values')[0]
            name_part, extension_part = os.path.splitext(file_name)
            if case_conversion == "lowercase":
                name_part = name_part.lower()
            elif case_conversion == "uppercase":
                name_part = name_part.upper()
            elif case_conversion == "titlecase":
                name_part = name_part.title()

            if replace_text:
                name_part = name_part.replace(replace_text, new_text)

            new_file_name = f"{prefix}{name_part}{suffix}{extension_part}"

            self.file_treeview.set(item, "Preview", new_file_name)

    def update_preview_on_selection_change(self, event):
        selected_items = self.file_treeview.selection()
        for item in self.file_treeview.get_children():
            if item in selected_items:
                original_file_name = self.file_treeview.item(item, "values")[0]
                replace_text = self.replace_entry.get()
                new_text = self.new_entry.get()
                name_part, extension_part = os.path.splitext(original_file_name)
                if replace_text in name_part:
                    new_name_part = name_part.replace(replace_text, new_text)
                else:
                    new_name_part = name_part
                new_file_name = new_name_part + extension_part
                self.file_treeview.set(item, "Preview", new_file_name)
            else:
                # Clear the Preview column if the item is not selected
                self.file_treeview.set(item, "Preview", "")


def main():
    root = ttk.Window(themename="superhero")

    app = BulkFileRenamer(root)
    root.mainloop()


if __name__ == "__main__":
    main()
