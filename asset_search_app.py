import ttkbootstrap as tb

from ttkbootstrap.constants import *

from tkinter import filedialog, messagebox

import pandas as pd

import tksheet



class AssetApp:

    def __init__(self, root):

        self.root = root

        self.root.title("IT Assets Management")

        self.root.geometry("1200x700")


        self.df = None


        # ===== HEADER =====

        header = tb.Label(root, text="IT Assets Management", font=("Segoe UI", 20, "bold"))

        header.pack(pady=10)


        # ===== CONTROLS =====

        frame = tb.Frame(root)

        frame.pack(pady=10)


        tb.Button(frame, text="Load Excel", bootstyle=SUCCESS, command=self.load_excel).grid(row=0, column=0, padx=10)


        tb.Label(frame, text="Search Column:").grid(row=0, column=1)

        self.column_box = tb.Combobox(frame, width=25, state="readonly")

        self.column_box.grid(row=0, column=2, padx=5)


        tb.Label(frame, text="Search Value:").grid(row=0, column=3)

        self.search_entry = tb.Entry(frame, width=25)

        self.search_entry.grid(row=0, column=4, padx=5)


        tb.Button(frame, text="Search", bootstyle=PRIMARY, command=self.search).grid(row=0, column=5, padx=5)

        tb.Button(frame, text="Show All", bootstyle=SECONDARY, command=self.show_all).grid(row=0, column=6, padx=5)


        # ===== TABLE =====

        table_frame = tb.Frame(root)

        table_frame.pack(fill=BOTH, expand=True, padx=20, pady=20)

        table_frame.rowconfigure(0, weight=1)

        table_frame.columnconfigure(0, weight=1)


        self.sheet = tksheet.Sheet(table_frame,
                                  show_x_scrollbar=True,
                                  show_y_scrollbar=True,
                                  font=("Segoe UI", 10, "normal"),
                                  header_font=("Segoe UI", 10, "bold"))

        self.sheet.grid(row=0, column=0, sticky="nsew")

        self.sheet.enable_bindings("single_select",
                                  "column_select",
                                  "row_select",
                                  "column_width_resize",
                                  "double_click_column_resize",
                                  "arrowkeys",
                                  "copy",
                                  "rc_select")


    # =========================================================

    # HEADER DETECTION (MOST IMPORTANT FIX)

    # =========================================================

    def detect_header_row(self, path, sheet_name=0):

        raw = pd.read_excel(path, sheet_name=sheet_name, header=None)


        for i, row in raw.iterrows():

            for cell in row:

                if isinstance(cell, str):

                    value = cell.strip().lower()

                    if value.startswith(("sl", "s1", "si")):

                        return i


        return None


    # =========================================================

    # LOAD EXCEL

    # =========================================================

    def load_excel(self):

        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])

        if not file_path:

            return


        try:

            # Get all sheet names

            xls = pd.ExcelFile(file_path)

            all_frames = []


            for sheet in xls.sheet_names:

                header_row = self.detect_header_row(file_path, sheet_name=sheet)


                if header_row is None:

                    # Skip sheets where no header is found

                    continue


                df_sheet = pd.read_excel(file_path, sheet_name=sheet, header=header_row)


                # Clean columns

                df_sheet.columns = df_sheet.columns.astype(str).str.strip()


                # Remove unnamed columns

                df_sheet = df_sheet.loc[:, ~df_sheet.columns.str.contains("^Unnamed")]


                # Remove fully empty rows

                df_sheet.dropna(how="all", inplace=True)


                # Tag each row with the source sheet name

                df_sheet.insert(0, "Sheet", sheet)


                all_frames.append(df_sheet)


            if not all_frames:

                raise Exception("❌ Could not find header row containing 'Sl' column in any sheet")


            # Combine all sheets (columns may differ; missing cols become NaN)

            self.df = pd.concat(all_frames, ignore_index=True)


            # Populate dropdown with all available columns

            self.column_box["values"] = list(self.df.columns)

            self.column_box.current(0)


            self.setup_table()

            self.show_all()


            loaded_sheets = [d["Sheet"].iloc[0] for d in all_frames]

            messagebox.showinfo("Success",

                f"✅ Loaded {len(all_frames)} sheet(s): {', '.join(loaded_sheets)}")


        except Exception as e:

            messagebox.showerror("Error", str(e))


    # =========================================================

    # TABLE SETUP

    # =========================================================

    def setup_table(self):

        self.sheet.headers(list(self.df.columns))


    # =========================================================

    # DISPLAY DATA

    # =========================================================

    def display(self, data):

        rows = []

        for _, row in data.iterrows():

            values = ["" if pd.isna(v) else str(v) for v in row]

            rows.append(values)

        self.sheet.set_sheet_data(rows)

        self.sheet.set_all_column_widths(150)


    # =========================================================

    # SEARCH FUNCTION

    # =========================================================

    def search(self):

        if self.df is None:

            return



        column = self.column_box.get()

        text = self.search_entry.get().lower().strip()



        filtered = self.df[self.df[column].astype(str).str.lower().str.contains(text, na=False)]

        self.display(filtered)



    # =========================================================

    # SHOW ALL

    # =========================================================

    def show_all(self):

        if self.df is not None:

            self.display(self.df)





# =========================================================

# RUN APP

# =========================================================

if __name__ == "__main__":

    root = tb.Window(themename="flatly")

    app = AssetApp(root)

    root.mainloop()