import customtkinter as ctk
from tkinter import filedialog
import pandas as pd
import os
import sys
from PIL import Image
import traceback
import win32com.client


# ===== DARK MODE =====
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")


# ===== GLOBAL =====
file_path = None
excel_path = None


# ===== RESOURCE PATH =====
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


# ===== APP =====
class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("CSV Converter Pro")
        self.geometry("900x600")
        self.minsize(800, 500)
        self.configure(fg_color="#1a1a1a")

        # ICON
        try:
            self.iconbitmap(resource_path("icon.ico"))
        except:
            pass

        # MAIN FRAME
        main_frame = ctk.CTkFrame(self)
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)

        # LOGO
        try:
            img = Image.open(resource_path("company_logo.png"))
            logo = ctk.CTkImage(light_image=img, dark_image=img, size=(120, 120))
            self.logo = logo
            ctk.CTkLabel(main_frame, image=self.logo, text="").pack(pady=10)
        except:
            pass

        # TITLE
        ctk.CTkLabel(
            main_frame,
            text="CSV Converter Pro",
            font=("Arial", 24, "bold")
        ).pack(pady=5)

        # BUTTONS
        btn_frame = ctk.CTkFrame(main_frame)
        btn_frame.pack(pady=15)

        ctk.CTkButton(btn_frame, text="Load CSV", command=self.load_csv).grid(row=0, column=0, padx=10)
        ctk.CTkButton(btn_frame, text="Convert to Excel", command=self.to_excel).grid(row=0, column=1, padx=10)
        ctk.CTkButton(btn_frame, text="Excel → PDF", command=self.to_pdf).grid(row=0, column=2, padx=10)

        # LOG BOX
        self.log_box = ctk.CTkTextbox(main_frame, height=200)
        self.log_box.pack(fill="both", expand=True, padx=10, pady=10)

    # ===== LOG =====
    def log(self, text):
        self.log_box.insert("end", text + "\n")
        self.log_box.see("end")

    # ===== LOAD CSV =====
    def load_csv(self):
        global file_path

        file_path = filedialog.askopenfilename(
            filetypes=[("CSV files", "*.csv")]
        )

        if file_path:
            self.log(f"Loaded: {file_path}")

    # ===== CSV → EXCEL =====
    def to_excel(self):
        global excel_path

        if not file_path:
            self.log("❌ Load CSV first")
            return

        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )

        if not save_path:
            return

        try:
            with pd.ExcelWriter(save_path, engine="openpyxl") as writer:
                df = pd.read_csv(file_path)
                df.to_excel(writer, index=False, sheet_name="Sheet1")

                ws = writer.sheets["Sheet1"]

                # автоширина колонок
                for column_cells in ws.columns:
                    length = max(len(str(cell.value)) if cell.value else 0 for cell in column_cells)
                    ws.column_dimensions[column_cells[0].column_letter].width = length + 2

            excel_path = save_path
            self.log(f"✅ Excel saved: {save_path}")

        except Exception as e:
            self.log(f"❌ Error: {str(e)}")
            self.log(traceback.format_exc())

    # ===== EXCEL → PDF =====
    def to_pdf(self):
        if not excel_path:
            self.log("❌ Convert to Excel first")
            return

        pdf_path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")]
        )

        if not pdf_path:
            return

        try:
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False

            wb = excel.Workbooks.Open(excel_path)
            ws = wb.Worksheets(1)

            # автоширина
            ws.Columns.AutoFit()

            # масштаб под страницу
            ws.PageSetup.Zoom = False
            ws.PageSetup.FitToPagesWide = 1
            ws.PageSetup.FitToPagesTall = False

            # альбомная ориентация
            ws.PageSetup.Orientation = 2

            wb.ExportAsFixedFormat(0, pdf_path)

            wb.Close(False)
            excel.Quit()

            self.log(f"✅ PDF saved: {pdf_path}")

        except Exception as e:
            self.log(f"❌ PDF Error: {str(e)}")
            self.log(traceback.format_exc())


# ===== RUN =====
if __name__ == "__main__":
    app = App()
    app.mainloop() 