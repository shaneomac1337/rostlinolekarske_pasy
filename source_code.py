import os
import sys
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from tkinter import filedialog
import openpyxl
from fuzzywuzzy import fuzz
import win32com.client
import requests
import semver
import webbrowser
import win32com.client as win32

current_version = "v0.6.1"
url = 'https://api.github.com/repos/{owner}/{repo}/releases/latest'
response = requests.get(url.format(owner='shaneomac1337', repo='rostlinolekarske_pasy'))

if response.status_code == requests.codes.ok:
    latest_release = response.json()
    latest_version = latest_release['tag_name'][1:]
    if semver.compare(current_version[1:], latest_version) < 0:
        # Display a message box to the user
        app = tk.Tk()
        app.withdraw()
        result = messagebox.askyesno('Aktualizace dostupná', 'Nová verze toolu na pasy je k dispozici, přeje si Olinka navštívit stáhnout novou verzi z webu?')

        if result:
            # Otevřít GitHub k nalezení aktuální verze
            url = latest_release['html_url']
            webbrowser.open_new(url)

            # Update je k dispozici, detekce platformy
            assets = latest_release['assets']
            for asset in assets:
                if 'Windows' in asset['name'] and 'x86_64' in asset['name']:
                    download_url = asset['browser_download_url']
                    r = requests.get(download_url)
        else:
            # Nedělat nic, pokud zvoleno "Ne"
            pass
    else:
        # Update není k dispozici
        pass
else:
    # Neobdržel jsem info o updatu
    pass

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

        self.delete_codes_button = ttk.Button(right_buttons_frame, text="Smazat vše", command=self.delete_codes_and_cz)
        self.delete_codes_button.grid(row=0, column=0, pady=10, sticky="w")

        self.missing_folder_button = ttk.Button(right_buttons_frame, text="Vytvořit složku missing", command=self.create_missing_folder)
        self.missing_folder_button.grid(row=1, column=0, pady=10, sticky="w")

        # Automatické updaty
        self.check_updates_button = ttk.Button(main_frame, text="Zkontrolovat aktualizace", command=self.check_for_updates)
        self.check_updates_button.grid(row=0, column=0, padx=(100, 0), pady=(0, 0), sticky="e")

        # Label verze
        self.version_label = ttk.Label(main_frame, text=current_version, font=("Helvetica", 10, "bold"))
        self.version_label.grid(row=0, column=1, pady=(0, 10), sticky="e")

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
        self.insert_image_button = ttk.Button(main_frame, text="Vložit EU obrázky", command=self.insert_image_to_excel)
        self.insert_image_button.grid(row=7, column=0, padx=10, pady=(0, 10), sticky="w")

        # Load Custom Template button
        self.load_template_button = ttk.Button(main_frame, text="Olinky vlastní šablona", command=self.load_custom_template)
        self.load_template_button.grid(row=6, column=0, padx=10, pady=(0, 10), sticky="w")

        # Unload Custom Template button
        self.unload_template_button = ttk.Button(main_frame, text="Použít defaultní šablonu", command=self.unload_custom_template)
        self.unload_template_button.grid(row=6, column=0, padx=(170, 260), pady=(0, 10), sticky="e")
        # Manual input for missing code
        self.manual_code_button = ttk.Button(right_buttons_frame, text="Manually Add Code", command=self.manually_add_code)
        self.manual_code_button.grid(row=2, column=0, pady=10, sticky="w")

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

    def check_for_updates(self):
        url = 'https://api.github.com/repos/{owner}/{repo}/releases/latest'
        response = requests.get(url.format(owner='shaneomac1337', repo='rostlinolekarske_pasy'))

        if response.status_code == requests.codes.ok:
            latest_release = response.json()
            latest_version = latest_release['tag_name'][1:]
            if semver.compare(current_version[1:], latest_version) < 0:
                # Display a message box to the user
                result = messagebox.askyesno('Aktualizace dostupná', 'Nová verze toolu na pasy je k dispozici, přeje si Olinka stáhnout novou verzi z webu?')

                if result:
                    # Open the Github page for the latest release in the user's default web browser
                    url = latest_release['html_url']
                    webbrowser.open_new(url)

                    # An update is available, download the asset(s) that match your platform and architecture
                    assets = latest_release['assets']
                    for asset in assets:
                        if 'Windows' in asset['name'] and 'x86_64' in asset['name']:
                            download_url = asset['browser_download_url']
                            r = requests.get(download_url)
                            # Save the downloaded asset to a file
                            with open('aplikace_v.0.4.0.exe', 'wb') as f:
                                f.write(r.content)
                else:
                    # Do nothing if the user clicks "No"
                    pass
            else:
                # No update available
                messagebox.showinfo('Žádné aktualizace', 'Pro Olinku není k dispozici bohužel žádná aktualizace')
        else:
            # Failed to retrieve latest release info
            messagebox.showerror('Chyba', 'Nedokázal jsem zjistit informace o nejnovější verzi z GitHubu.')        

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
        if not self.has_excel_files():
            messagebox.showerror("Chyba", "Olinka už mi zase nedala Excely, dělá si ze mě prdel??")
        else:
            for filename in os.listdir('.'):
                if not filename.endswith('.xlsx') or filename == 'template.xlsx':
                    continue 
                wb = openpyxl.load_workbook(filename)
                for sheet in wb:
                    ws = wb[sheet.title]
                    for row in range(13, ws.max_row + 1):
                        ws.cell(row=row, column=3).value = None  # This line deletes the text from column C
                        ws.cell(row=row, column=4).value = None
                        ws.cell(row=row, column=5).value = None
                wb.save(filename)

            self.output_console.insert(tk.END, "Je to fuč..\n")
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
            ws.PageSetup.Orientation = 2  # 2 represents landscape orientation
            ws.PageSetup.Zoom = False  # Turn off Zoom property
            ws.PageSetup.FitToPagesWide = 1  # Fit to 1 page wide
            ws.PageSetup.FitToPagesTall = 1  # Fit to 1 page tall
            ws.ExportAsFixedFormat(0, os.path.abspath(pdf_file))

        except Exception as e:
            print(f"Failed to convert {excel_file} - {sheet_name} to PDF: {e}")

        finally:
            wb.Close(SaveChanges=False)
            xlApp.Quit()


    def save_all_excels_as_pdfs(self):
        has_excel_files = False
        for filename in os.listdir('.'):
            if filename.endswith('.xlsx') and filename != 'template.xlsx':
                has_excel_files = True
                wb = openpyxl.load_workbook(filename)
                for sheet in wb:
                    self.save_excel_as_pdf(filename, sheet.title)
                    self.output_console.insert(tk.END, f"Uloženo jako:{sheet.title}\n")
                    self.output_console.see(tk.END)  # Auto-scroll to the end
                    self.output_console.update()  # Ensure the output console is updated
    
        if not has_excel_files:
            messagebox.showerror("Chyba", "A teď zase Olinka nedodala Excely na konvertování do PDF. Bože muj.")
        else:
            self.output_console.insert(tk.END, "Všechny listy v excelu byly uloženy jako samostatné PDF.\n")
            self.output_console.see(tk.END)  # Auto-scroll to the end
            self.output_console.update()  # Ensure the output console is updated
    def insert_image_to_excel(self):
        def insert_image(excel_file_path, image_file_path, cell_name, row_height=None, column_width=None, pic_width=None, pic_height=None):
            # Open Excel
            excel = win32.gencache.EnsureDispatch('Excel.Application')
            excel.Visible = False  # If you want Excel Application to be visible during execution, set this to True

            # Open Workbook
            wb = excel.Workbooks.Open(excel_file_path)

            # Iterate through each sheet in the workbook
            for sheet in wb.Sheets:

                # Get cell dimensions
                target_cell = sheet.Range(cell_name)
                top = target_cell.Top
                left = target_cell.Left

                # Set row height and column width for the entire sheet if provided
                if row_height is not None:
                    sheet.Rows.RowHeight = row_height
                if column_width is not None:
                    sheet.Columns.ColumnWidth = column_width

                # Add picture
                pic = sheet.Pictures().Insert(image_file_path)

                # Set image position
                pic.Top = top
                pic.Left = left

                # Set image size
                if pic_width is not None:
                    pic.Width = pic_width
                if pic_height is not None:
                    pic.Height = pic_height

            # Save and Close
            wb.Save()
            wb.Close()

            # Quit Excel
            excel.Quit()

        # Get the current working directory
        directory_path = os.getcwd()
        cell_name = "C6"

        # Set your desired row height
        row_height = 15.75  # Adjust this value to change the row height for all rows in the sheet

        # Set your desired picture width and height
        pic_width = 147  # Adjust this value to set the picture width
        pic_height = 80  # Adjust this value to set the picture height

        # Get a list of all Excel files in the directory
        excel_files = [f for f in os.listdir(directory_path) if f.endswith('.xlsx') or f.endswith('.xls')]

        # Run the function on each Excel file
        for excel_file in excel_files:
            if excel_file == 'template.xlsx':
                continue  # Skip the template file
            excel_file_path = os.path.join(directory_path, excel_file)
            image_file_path = os.path.join(directory_path, "eu.png")  # The image file is in the same directory as the Excel files
            insert_image(excel_file_path, image_file_path, cell_name, row_height, None, pic_width, pic_height)
            self.output_console.insert(tk.END, f"Obrázek vložen do: {excel_file}\n")
            self.output_console.see(tk.END)  # Auto-scroll to the end
            self.output_console.update()  # Ensure the output console is updated       

    def manually_add_code(self):
        def submit_code():
            selected_name = name_listbox.get(name_listbox.curselection())
            code = code_entry.get()
            if selected_name and code:
                self.codes[selected_name] = code
                for filename, sheet_title, name, row in unmatched_names:
                    if name == selected_name:
                        wb = openpyxl.load_workbook(filename)
                        ws = wb[sheet_title]
                        ws.cell(row=row, column=3).value = selected_name  # update column 3 with the selected name
                        ws.cell(row=row, column=4).value = code  # update column 4 with the code
                        ws.cell(row=row, column=5).value = "CZ"  # update column 5 with "CZ"
                        wb.save(filename)
                        break
                top.destroy()
                self.output_console.insert(tk.END, f"Manually added code: {selected_name} - {code}\n")
                self.output_console.see(tk.END)
                self.output_console.update()

        def get_unmatched_names():
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
            return unmatched_names

        top = tk.Toplevel(self.master)
        top.title("Manually Add Code")

        name_label = ttk.Label(top, text="Name:")
        name_label.grid(row=0, column=0, padx=10, pady=10, sticky="w")
        name_listbox = tk.Listbox(top, selectmode=tk.SINGLE, exportselection=0)
        name_listbox.grid(row=0, column=1, padx=10, pady=10, sticky="w")

        unmatched_names = get_unmatched_names()
        for _, _, name, _ in unmatched_names:
            name_listbox.insert(tk.END, name)

        code_label = ttk.Label(top, text="Code:")
        code_label.grid(row=1, column=0, padx=10, pady=10, sticky="w")
        code_entry = ttk.Entry(top)
        code_entry.grid(row=1, column=1, padx=10, pady=10, sticky="w")

        submit_button = ttk.Button(top, text="Submit", command=submit_code)
        submit_button.grid(row=2, column=1, padx=10, pady=10, sticky="e")
            
    
    def quit(self):
        self.master.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = PlantCodeFinder(root)
    root.resizable(False, False)  # Disable maximizing
    root.mainloop()