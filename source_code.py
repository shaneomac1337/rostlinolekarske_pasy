import os
import sys
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from tkinter import filedialog
import openpyxl
from fuzzywuzzy import fuzz
import win32com.client

class PlantCodeFinder(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.master.title("Olinky aplikace na pasy :)")
        self.grid()
        self.create_widgets()

    def create_widgets(self):
        main_frame = ttk.Frame(self)
        main_frame.grid(row=0, column=0, padx=10, pady=10)

        # Section 1: Buttons
        section1_label = ttk.Label(main_frame, text="Akce", font=("Helvetica", 12, "bold"))
        section1_label.grid(row=0, column=0, pady=(0, 10), sticky="w")

        buttons_frame = ttk.Frame(main_frame, borderwidth=2, relief="groove", padding=10)
        buttons_frame.grid(row=1, column=0, sticky="w")

        # For 2 columns of buttons, divide the buttons_frame into 2 "sub-frames"
        left_buttons_frame = ttk.Frame(buttons_frame)
        left_buttons_frame.pack(side='left', padx=10)

        right_buttons_frame = ttk.Frame(buttons_frame)
        right_buttons_frame.pack(side='left', padx=10)

        self.process_button = ttk.Button(left_buttons_frame, text="Zpracovat soubory", command=self.process_files)
        self.process_button.grid(row=0, column=0, pady=10, sticky="w")

        self.show_unmatched_button = ttk.Button(left_buttons_frame, text="Zobrazit neshody", command=self.show_unmatched_names)
        self.show_unmatched_button.grid(row=1, column=0, pady=10, sticky="w")

        self.delete_codes_button = ttk.Button(right_buttons_frame, text="Smazat kódy a CZ", command=self.delete_codes_and_cz)
        self.delete_codes_button.grid(row=0, column=0, pady=10, sticky="w")

        self.missing_folder_button = ttk.Button(right_buttons_frame, text="Vytvořit složku missing", command=self.create_missing_folder)
        self.missing_folder_button.grid(row=1, column=0, pady=10, sticky="w")

        # Section 2: Checkboxes
        section2_label = ttk.Label(main_frame, text="Nastavení", font=("Helvetica", 12, "bold"))
        section2_label.grid(row=2, column=0, pady=(0, 10), sticky="w")

        checkboxes_frame = ttk.Frame(main_frame)
        checkboxes_frame.grid(row=3, column=0, sticky="w")

        self.delete_checkbox_var = tk.BooleanVar()
        self.delete_checkbox = ttk.Checkbutton(checkboxes_frame, text="Smazat soubory, kde je vše v poho", variable=self.delete_checkbox_var)
        self.delete_checkbox.grid(row=0, column=0, padx=10, pady=10, sticky="w")

        self.fulfill_cz_checkbox_var = tk.BooleanVar()
        self.fulfill_cz_checkbox = ttk.Checkbutton(checkboxes_frame, text="Vyplnit CZ ke kodum", variable=self.fulfill_cz_checkbox_var)
        self.fulfill_cz_checkbox.grid(row=1, column=0, padx=10, pady=10, sticky="w")

        # Section 3: Output console
        section3_label = ttk.Label(main_frame, text="Výstup", font=("Helvetica", 12, "bold"))
        section3_label.grid(row=4, column=0, pady=(0, 10), sticky="w")

        self.output_console = tk.Text(main_frame, wrap=tk.WORD, height=10, width=80, relief="sunken", borderwidth=1)
        self.output_console.grid(row=5, column=0, columnspan=2, padx=10, pady=10)

        # Load Custom Template button
        self.load_template_button = ttk.Button(main_frame, text="Olinky vlastní šablona", command=self.load_custom_template)
        self.load_template_button.grid(row=6, column=0, padx=10, pady=(0, 10), sticky="w")

        # Unload Custom Template button
        self.unload_template_button = ttk.Button(main_frame, text="Použít defaultní šablonu", command=self.unload_custom_template)
        self.unload_template_button.grid(row=6, column=0, padx=(170, 260), pady=(0, 10), sticky="e")

        # Quit button
        self.quit_button = ttk.Button(main_frame, text="Ukončit", command=self.quit)
        self.quit_button.grid(row=6, column=1, padx=10, pady=(0, 10), sticky="e")

        self.save_excels_as_pdfs_button = ttk.Button(main_frame, text="Uložit Excely jako PDFka", command=self.save_all_excels_as_pdfs)
        self.save_excels_as_pdfs_button.grid(row=6, column=0, padx=(170, 100), pady=(0, 10), sticky="e")

        # Initialize instance variables
        self.template_wb = openpyxl.load_workbook(self.get_template())
        self.codes = {}
        for row in range(2, self.template_wb.active.max_row + 1):
            name = self.template_wb.active.cell(row=row, column=1).value
            code = self.template_wb.active.cell(row=row, column=2).value
            self.codes[name] = code

    def has_excel_files(self):
        for filename in os.listdir('.'):
            if filename.endswith('.xlsx') and filename != 'template.xlsx':
                return True
        return False

    def get_template(self):
        if getattr(sys, 'frozen', False):
            template_path = os.path.join(sys._MEIPASS, 'template.xlsx')
        else:
            template_path = 'template.xlsx'
        return template_path

    def load_custom_template(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            self.template_wb = openpyxl.load_workbook(file_path)
            self.codes = {}
            for row in range(2, self.template_wb.active.max_row + 1):
                name = self.template_wb.active.cell(row=row, column=1).value
                code = self.template_wb.active.cell(row=row, column=2).value
                self.codes[name] = code
            self.output_console.insert(tk.END, f"Olinka podstrčila tuhle šablonu: {file_path}\n")
            self.output_console.see(tk.END)  # Auto-scroll to the end
            self.output_console.update()  # Ensure the output console is updated

    def unload_custom_template(self):
        self.codes = {}
        self.template_wb = openpyxl.load_workbook(self.get_template())
        for row in range(2, self.template_wb.active.max_row + 1):
            name = self.template_wb.active.cell(row=row, column=1).value
            code = self.template_wb.active.cell(row=row, column=2).value
            self.codes[name] = code
        self.output_console.insert(tk.END, "Defaultní šablona načtena.\n")
        self.output_console.see(tk.END)  # Auto-scroll to the end
        self.output_console.update()  # Ensure the output console is updated   

    def show_unmatched_names(self):
        unmatched_names = []
        for filename in os.listdir('.'):
            if not filename.endswith('.xlsx') or filename == 'template.xlsx':
                continue

            wb = openpyxl.load_workbook(filename)
            for sheet in wb:
                ws = wb[sheet.title]
                for row in range(13, ws.max_row + 1):
                    name = ws.cell(row=row, column=3).value
                    if name:
                        closest_match = max(self.codes.keys(), key=lambda x: fuzz.ratio(x, name))
                        if fuzz.ratio(closest_match, name) < 80:
                            unmatched_names.append((filename, sheet.title, name, row))

        if unmatched_names:
            message = "Nenalezené shody:\n\n"
            for filename, sheet_title, name, row in unmatched_names:
                message += f"{filename} - {sheet_title} - {name} (řádek {row})\n"
            self.output_console.insert(tk.END, message)
            self.output_console.see(tk.END)  # Auto-scroll to the end
            self.output_console.update()  # Ensure the output console is updated
        else:
            self.output_console.insert(tk.END, "Nenašel jsem neshody.\n")
            self.output_console.see(tk.END)  # Auto-scroll to the end
            self.output_console.update()  # Ensure the output console is updated


    def process_files(self):
    # Check if there are any Excel files to process
        if not self.has_excel_files():
            messagebox.showwarning("Kde jsou sakra Excely???", "Nejdřív mi musí Olinka navalit nějaký ten Excel do stejné složky (muže jich být nespočet), kde jsem byl spuštěn. Tak je koukej vysolit!!! Nebo vyšli Edíka s Kačenkou ať je donesou...")
            return

         # Create a "missing" folder if it doesn't exist
        if not os.path.exists("missing"):
            os.makedirs("missing")
            

        # Process all Excel files in the current directory
        for filename in os.listdir('.'):
            # Skip non-Excel files and the template file
            if not filename.endswith('.xlsx') or filename == 'template.xlsx':
                continue

            # Load the Excel file
            wb = openpyxl.load_workbook(filename)

            # Process all sheets in the Excel file
            for sheet in wb:
                # Get the active sheet
                ws = wb[sheet.title]

                # Process all rows in the sheet and add missing plant codes
                missing_names = []
                for row in range(13, ws.max_row + 1):
                    name = ws.cell(row=row, column=3).value
                    if name:
                        # Find the closest matching plant code in the dictionary
                        closest_match = max(self.codes.keys(), key=lambda x: fuzz.ratio(x, name))
                        if fuzz.ratio(closest_match, name) >= 80:
                            ws.cell(row=row, column=4).value = self.codes[closest_match]
                            if self.fulfill_cz_checkbox_var.get():
                                ws.cell(row=row, column=5).value = 'CZ'
                        else:
                            if self.fulfill_cz_checkbox_var.get():
                                ws.cell(row=row, column=5).value = 'CZ'
                            missing_names.append((name, row))                        

                # Save the modified Excel file
                wb.save(filename)

                # Write missing plant names to a text file in the "missing" folder
                if missing_names:
                    with open(f"missing/{filename}_{sheet.title}_missing_names.txt", 'w') as f:
                        f.write(f"Missing names v excel sheetu:  '{sheet.title}' v souboru '{filename}':\n")
                        for name, row in missing_names:
                            f.write(f"Jméno '{name}' chybí na řádce {row}\n")
                elif not self.delete_checkbox_var.get():
                    with open(f"missing/{filename}_{sheet.title}_is_okay.txt", 'w') as f:
                        f.write(f"Všechna jména se nachází v excel sheetu: '{sheet.title}' v souboru: '{filename}'\n")

        # Delete files without missing plant names if the checkbox is checked
        if self.delete_checkbox_var.get():
            files_to_delete = [filename for filename in os.listdir("missing") if filename.endswith("_is_okay.txt")]
            for filename in files_to_delete:
                os.remove(f"missing/{filename}")
                self.output_console.insert(tk.END, f"Deleted file: {filename}\n")
                self.output_console.see(tk.END)  # Auto-scroll to the end
                self.output_console.update()  # Ensure the output console is updated

        self.output_console.insert(tk.END, "Hotovo Olinko!\n")
        self.output_console.see(tk.END)  # Auto-scroll to the end
        self.output_console.update()  # Ensure the output console is updated

        messagebox.showinfo("Povedlo se!", "Teď už to jen stačí poslat blbečkům na mail :).")

    def delete_codes_and_cz(self):
        for filename in os.listdir('.'):
            if not filename.endswith('.xlsx') or filename == 'template.xlsx':
                continue 
            wb = openpyxl.load_workbook(filename)
            for sheet in wb:
                ws = wb[sheet.title]
                for row in range(13, ws.max_row + 1):
                    ws.cell(row=row, column=4).value = None
                    ws.cell(row=row, column=5).value = None

            wb.save(filename)

        self.output_console.insert(tk.END, "Vyčistil jsem kódy a CZ ze všech Excelů.\n")
        self.output_console.see(tk.END)  # Auto-scroll to the end
        self.output_console.update()  # Ensure the output console is updated

    def create_missing_folder(self):
        if not os.path.exists("missing"):
            os.makedirs("missing")
            self.output_console.insert(tk.END, "Složka missing vytvořena.\n")
            self.output_console.see(tk.END)  # Auto-scroll to the end
            self.output_console.update()  # Ensure the output console is updated
        else:
            self.output_console.insert(tk.END, "Složka missing už existuje.\n")
            self.output_console.see(tk.END)  # Auto-scroll to the end
            self.output_console.update()  # Ensure the output console is updated

    def save_excel_as_pdf(self, excel_file, sheet_name):
        """Save the given sheet in excel file as pdf in the 'pdf' folder."""
        pdf_folder = "pdf"
        if not os.path.exists(pdf_folder):
            os.makedirs(pdf_folder)

        pdf_file = f"{pdf_folder}/{sheet_name}.pdf"

        try:
            xlApp = win32com.client.Dispatch("Excel.Application")
            xlApp.Visible = False

            wb = xlApp.Workbooks.Open(os.path.abspath(excel_file), ReadOnly=1)
            ws = wb.Worksheets(sheet_name)
            ws.ExportAsFixedFormat(0, os.path.abspath(pdf_file))

        except Exception as e:
            print(f"Failed to convert {excel_file} - {sheet_name} to PDF: {e}")

        finally:
            wb.Close(SaveChanges=False)
            xlApp.Quit()

    def save_all_excels_as_pdfs(self):
        for filename in os.listdir('.'):
            if filename.endswith('.xlsx') and filename != 'template.xlsx':
                wb = openpyxl.load_workbook(filename)
                for sheet in wb:
                    self.save_excel_as_pdf(filename, sheet.title)
                    self.output_console.insert(tk.END, f"Uloženo jako:{sheet.title}\n")
                    self.output_console.see(tk.END)  # Auto-scroll to the end
                    self.output_console.update()  # Ensure the output console is updated

        self.output_console.insert(tk.END, "Všechny listy v excelu byly uloženy jako samostatné PDF.\n")
        self.output_console.see(tk.END)  # Auto-scroll to the end
        self.output_console.update()  # Ensure the output console is updated        
    
    def quit(self):
        self.master.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = PlantCodeFinder(root)
    root.resizable(False, False)  # Disable maximizing
    root.mainloop()