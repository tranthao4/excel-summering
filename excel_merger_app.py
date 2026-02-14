import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import os
from config import COLUMNS, SUMMARY_COLUMNS, VALIDATION_RULES, UI_SETTINGS, FILES


class ExcelMergerApp:
    def __init__(self, root):
        self.root = root
        self.root.title(UI_SETTINGS['window_title'])

        # Låt fönstret anpassa sig efter innehållet
        self.root.resizable(True, True)

        self.summary_file = FILES['summary_file']
        self.current_file = None
        self.current_columns = []
        self.column_combos = {}  # Lagra alla comboboxes med ID som nyckel
        self.imported_files = set()  # Håll koll på importerade filer

        self.setup_ui()

        # Centrera fönstret på skärmen efter att UI är uppbyggt
        self.root.update_idletasks()
        self.center_window()
        
    def setup_ui(self):
        # Konfigurera PostNord TPL-färger
        primary_color = UI_SETTINGS.get('primary_color', '#00A0DC')
        bg_color = UI_SETTINGS.get('background_color', '#FFFFFF')

        # Sätt bakgrundsfärg
        self.root.configure(bg=bg_color)

        # Skapa stil för ttk widgets
        style = ttk.Style()
        style.configure('TFrame', background=bg_color)
        style.configure('TLabel', background=bg_color)
        style.configure('TLabelframe', background=bg_color)
        style.configure('TLabelframe.Label', background=bg_color)
        style.configure('PostNord.TButton', font=('Arial', 10, 'bold'))

        # Huvudram
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Logo-ram
        logo_frame = tk.Frame(main_frame, bg=bg_color)
        logo_frame.grid(row=0, column=0, columnspan=2, pady=(0, 10))

        # PostNord TPL Logo (text-baserad)
        logo_text = tk.Label(logo_frame, text="postnord",
                            font=('Arial', 28, 'bold'), fg=primary_color, bg=bg_color)
        logo_text.pack()

        logo_subtext = tk.Label(logo_frame, text="TPL",
                               font=('Arial', 16), fg=primary_color, bg=bg_color)
        logo_subtext.pack()

        # Välj fil knapp
        self.select_file_btn = ttk.Button(main_frame, text="Välj Excel-fil", command=self.select_file)
        self.select_file_btn.grid(row=2, column=0, columnspan=2, pady=10, sticky=tk.W+tk.E)

        # Filnamn label
        self.file_label = ttk.Label(main_frame, text="Ingen fil vald", foreground="gray")
        self.file_label.grid(row=3, column=0, columnspan=2, pady=5)

        # Mappning sektion - dynamiskt genererad från config
        mapping_frame = ttk.LabelFrame(main_frame, text="Kolumnmappning", padding="10")
        mapping_frame.grid(row=4, column=0, columnspan=2, pady=10, sticky=tk.W+tk.E)

        # Skapa UI-element för varje kolumn från config
        for row_idx, col_config in enumerate(COLUMNS):
            col_id = col_config['id']

            # Label
            ttk.Label(mapping_frame, text=col_config['label']).grid(row=row_idx, column=0, sticky=tk.W, pady=5)

            # Combobox
            combo = ttk.Combobox(mapping_frame, state="disabled", width=30)
            combo.grid(row=row_idx, column=1, pady=5, padx=5)
            self.column_combos[col_id] = combo

            # Clear button
            ttk.Button(mapping_frame, text=UI_SETTINGS['clear_button_text'],
                      width=UI_SETTINGS['clear_button_width'],
                      command=lambda c=combo: c.set('')).grid(row=row_idx, column=2, padx=2)

            # Help text - automatiskt genererad om inte angiven
            help_text = col_config.get('help_text')
            if help_text is None and not col_config['required']:
                help_text = '(Valfritt)'

            if help_text:
                ttk.Label(mapping_frame, text=help_text,
                         font=('Arial', 8), foreground="gray").grid(row=row_idx, column=3, sticky=tk.W, padx=5)

        # Behåll gamla referenser för bakåtkompatibilitet
        self.name_combo = self.column_combos['name']
        self.name_combo2 = self.column_combos['lastname']
        self.address_combo = self.column_combos['address']
        self.personnr_combo = self.column_combos['personnr']
        self.date_from_combo = self.column_combos['date_from']
        self.date_to_combo = self.column_combos['date_to']
        self.days_combo = self.column_combos['days']
        self.unit_combo = self.column_combos['unit']
        
        # Lägg till data knapp med PostNord-färg
        self.add_data_btn = ttk.Button(main_frame, text="Lägg till i summeringsfil",
                                       command=self.add_to_summary, state="disabled")
        self.add_data_btn.grid(row=5, column=0, columnspan=2, pady=20)

        # Knapp för att visa importerade filer
        self.show_imported_btn = ttk.Button(main_frame, text="Visa importerade filer (0)",
                                            command=self.show_imported_files)
        self.show_imported_btn.grid(row=6, column=0, columnspan=2, pady=5)

        # Lista för att spara info om importerade filer (för popup)
        self.imported_files_info = []

        # Röd knapp för att återställa summeringsfil
        self.reset_all_btn = tk.Button(main_frame, text="Återställ summeringsfil",
                                       command=self.confirm_reset_all,
                                       bg='#DC3545', fg='white',
                                       font=('Arial', 10, 'bold'),
                                       activebackground='#C82333', activeforeground='white',
                                       padx=15, pady=5, cursor='hand2',
                                       relief='flat', borderwidth=0)
        self.reset_all_btn.grid(row=7, column=0, columnspan=2, pady=20)

    def center_window(self):
        """Centrera fönstret på skärmen"""
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        self.root.geometry(f'{width}x{height}+{x}+{y}')

    def select_file(self):
        filename = filedialog.askopenfilename(
            title="Välj Excel-fil",
            filetypes=[("Excel filer", "*.xlsx *.xls"), ("Alla filer", "*.*")]
        )

        if filename:
            # Kontrollera om filen redan har importerats
            if filename in self.imported_files:
                messagebox.showwarning("Varning", f"Filen '{os.path.basename(filename)}' har redan importerats!")
                return

            try:
                df = pd.read_excel(filename)
                self.current_file = filename
                self.current_columns = list(df.columns)

                # Uppdatera UI
                self.file_label.config(text=os.path.basename(filename), foreground="black")

                # Rensa gamla val och aktivera comboboxes med nya kolumnnamn
                for combo in self.column_combos.values():
                    combo.set('')  # Rensa gammalt val
                    combo['values'] = [''] + self.current_columns  # Lägg till tom option
                    combo['state'] = "readonly"

                # Automatisk igenkänning av kolumner
                self.auto_detect_columns(df)

                self.add_data_btn['state'] = "normal"

            except Exception as e:
                messagebox.showerror("Fel", f"Kunde inte läsa filen:\n{str(e)}")

    def auto_detect_columns(self, df):
        """Automatisk igenkänning av kolumner baserat på namn och innehåll"""

        def match_column(col_name, keywords, exclude_keywords=None, exact=False):
            """Kolla om kolumnnamnet matchar något nyckelord"""
            col_lower = str(col_name).lower().strip()

            # Kolla om kolumnen innehåller något exkluderat ord
            if exclude_keywords:
                if any(excl in col_lower for excl in exclude_keywords):
                    return False

            if exact:
                # Exakt matchning (hela kolumnnamnet måste matcha)
                return col_lower in keywords
            else:
                # Partiell matchning
                return any(keyword in col_lower for keyword in keywords)

        def has_personnummer_format(series):
            """Kolla om kolumnen innehåller personnummer (10 siffror)"""
            if len(series) == 0:
                return False
            # Ta ett sample och kolla format
            sample = series.dropna().head(10).astype(str)
            # Räkna hur många som ser ut som personnummer (minst 10 siffror)
            min_digits = VALIDATION_RULES['personnummer']['min_digits']
            threshold = VALIDATION_RULES['personnummer']['match_threshold']
            matches = sum(1 for val in sample if len(str(val).replace('-', '').replace(' ', '').replace('.', '')) >= min_digits and
                         sum(c.isdigit() for c in str(val)) >= min_digits)
            return matches >= len(sample) * threshold

        def has_date_format(series):
            """Kolla om kolumnen innehåller datum"""
            if len(series) == 0:
                return False
            # Försök konvertera till datum
            try:
                pd.to_datetime(series.dropna().head(10), errors='coerce')
                return True
            except:
                return False

        # Först: Matcha kolumner med innehållsvalidering (t.ex. personnummer)
        for col_config in COLUMNS:
            if col_config.get('validate_content') == 'personnummer':
                col_id = col_config['id']
                combo = self.column_combos[col_id]

                if not combo.get():
                    for col in df.columns:
                        if match_column(col, col_config['keywords'], col_config.get('exclude_keywords', [])) and \
                           has_personnummer_format(df[col]):
                            combo.set(col)
                            break

        # Sedan: Matcha resten av kolumnerna
        for col_config in COLUMNS:
            # Skippa om redan matchad eller om den har innehållsvalidering (redan gjord ovan)
            if col_config.get('validate_content'):
                continue

            col_id = col_config['id']
            combo = self.column_combos[col_id]

            if not combo.get():
                for col in df.columns:
                    if match_column(col, col_config['keywords'], col_config.get('exclude_keywords', [])):
                        combo.set(col)
                        break

    def add_to_summary(self):
        # Validera att alla obligatoriska mappningar är valda (dynamiskt från config)
        missing_fields = []
        for col_config in COLUMNS:
            if col_config['required']:
                combo = self.column_combos[col_config['id']]
                if not combo.get() or not combo.get().strip():
                    missing_fields.append(col_config['label'].rstrip(':'))

        if missing_fields:
            messagebox.showwarning("Varning", f"Vänligen välj följande obligatoriska fält:\n{', '.join(missing_fields)}")
            return

        try:
            # Läs kundens fil
            df_customer = pd.read_excel(self.current_file)

            # Skapa mappning dynamiskt från config
            column_mapping = {}
            name_columns = []
            date_columns = []
            days_column = []

            for col_config in COLUMNS:
                col_id = col_config['id']
                combo = self.column_combos[col_id]
                selected_col = combo.get()

                # Skippa om inget är valt
                if not selected_col or not selected_col.strip():
                    continue

                output_name = col_config.get('output_name')

                # Speciella hanteringar för vissa kolumner
                if col_id == 'lastname':
                    # Efternamn kombineras med namn
                    name_columns.append(selected_col)
                elif col_id == 'name':
                    # Namn hanteras separat
                    name_columns.insert(0, selected_col)  # Förnamn först
                elif col_id == 'date_from':
                    # Datum från - spara med output_name
                    date_columns.append((selected_col, output_name))
                elif col_id == 'date_to':
                    # Datum till - spara med output_name
                    date_columns.append((selected_col, output_name))
                elif col_id == 'days':
                    # Anställningsdagar
                    days_column.append(selected_col)
                elif output_name:
                    # Vanlig kolumn med output_name
                    column_mapping[selected_col] = output_name

            # Validera att antingen datum ELLER dagar är angivna
            has_dates = len(date_columns) > 0
            has_days = len(days_column) > 0

            if not has_dates and not has_days:
                messagebox.showwarning("Varning", "Vänligen ange antingen Anställd från datum ELLER Anställningsdagar!")
                return

            # Välj och byt namn på kolumner
            date_col_names = [col[0] for col in date_columns]  # Extrahera kolumnnamn från tuples
            columns_to_select = list(column_mapping.keys()) + name_columns + date_col_names + days_column
            df_mapped = df_customer[columns_to_select].copy()

            # Konvertera anställningsdagar till int om det är en direkt kolumn (hanterar både str och int)
            if has_days and days_column:
                days_col_name = days_column[0]
                df_mapped[days_col_name] = pd.to_numeric(df_mapped[days_col_name], errors='coerce').fillna(0).astype(int)
                column_mapping[days_col_name] = 'Anställningsdagar'

            # Kombinera namnkolumner om det finns två
            if len(name_columns) == 2:
                df_mapped['Namn'] = df_mapped[name_columns[0]].astype(str) + ' ' + df_mapped[name_columns[1]].astype(str)
                df_mapped = df_mapped.drop(columns=name_columns)
            elif len(name_columns) == 1:
                df_mapped = df_mapped.rename(columns={name_columns[0]: 'Namn'})

            # Hantera datumkolumner - behåll dem och byt namn
            if has_dates:
                date_from_col, date_from_output = date_columns[0]
                df_mapped[date_from_col] = pd.to_datetime(df_mapped[date_from_col], errors='coerce')

                # Byt namn på från-datum
                df_mapped = df_mapped.rename(columns={date_from_col: date_from_output})

                if len(date_columns) == 2:
                    # Båda datum angivna - beräkna skillnad
                    date_to_col, date_to_output = date_columns[1]
                    df_mapped[date_to_col] = pd.to_datetime(df_mapped[date_to_col], errors='coerce')
                    df_mapped['Anställningsdagar'] = (df_mapped[date_to_col] - df_mapped[date_from_output]).dt.days
                    # Byt namn på till-datum
                    df_mapped = df_mapped.rename(columns={date_to_col: date_to_output})
                else:
                    # Endast från-datum - beräkna till idag
                    df_mapped['Anställningsdagar'] = (pd.Timestamp.now() - df_mapped[date_from_output]).dt.days

                # Formatera datum till endast datum (utan tid)
                if date_from_output in df_mapped.columns:
                    df_mapped[date_from_output] = df_mapped[date_from_output].dt.strftime('%Y-%m-%d')
                if len(date_columns) == 2 and date_to_output in df_mapped.columns:
                    df_mapped[date_to_output] = df_mapped[date_to_output].dt.strftime('%Y-%m-%d')

            # Byt namn på resterande kolumner
            df_mapped = df_mapped.rename(columns=column_mapping)

            # Konvertera anställningsdagar till int (hanterar både str och int)
            if 'Anställningsdagar' in df_mapped.columns:
                df_mapped['Anställningsdagar'] = pd.to_numeric(df_mapped['Anställningsdagar'], errors='coerce').fillna(0).astype(int)

            # Formatera personnummer: ta bort alla tecken utom siffror, behåll endast 10 siffror, fyll med ledande nollor
            df_mapped['Personnummer'] = df_mapped['Personnummer'].astype(str).str.replace(r'\D', '', regex=True).str[-10:].str.zfill(10)

            # Läs eller skapa summeringsfil
            if os.path.exists(self.summary_file):
                df_summary = pd.read_excel(self.summary_file, dtype={'Personnummer': str})
                # Formatera personnummer i befintlig fil på samma sätt
                if 'Personnummer' in df_summary.columns and len(df_summary) > 0:
                    df_summary['Personnummer'] = df_summary['Personnummer'].astype(str).str.replace(r'\D', '', regex=True).str[-10:].str.zfill(10)
            else:
                df_summary = pd.DataFrame(columns=SUMMARY_COLUMNS)

            # Säkerställ att df_mapped har alla kolumner som behövs
            for col in SUMMARY_COLUMNS:
                if col not in df_mapped.columns:
                    df_mapped[col] = ''

            # Lägg till alla rader (samma personnummer kan förekomma flera gånger)
            added = len(df_mapped)
            df_summary = pd.concat([df_summary, df_mapped], ignore_index=True)
            
            # Spara summeringsfil med personnummer som text
            with pd.ExcelWriter(self.summary_file, engine='openpyxl') as writer:
                df_summary.to_excel(writer, index=False, sheet_name='Sheet1')
                # Formatera personnummer-kolumnen som text
                worksheet = writer.sheets['Sheet1']
                for row in range(2, len(df_summary) + 2):  # Börja från rad 2 (efter header)
                    cell = worksheet.cell(row=row, column=3)  # Kolumn C (Personnummer)
                    cell.number_format = '@'  # Text format

            messagebox.showinfo("Klart!", f"Data tillagd i summeringsfilen!\n\nNya rader: {added}")

            # Lägg till filen i listan över importerade filer
            self.imported_files.add(self.current_file)
            self.imported_files_info.append(f"✓ {os.path.basename(self.current_file)} ({added} rader)")

            # Uppdatera knapptext med antal
            self.show_imported_btn.config(text=f"Visa importerade filer ({len(self.imported_files_info)})")

            # Rensa vald fil och mappning så man kan börja om
            self.reset_selection()

        except Exception as e:
            messagebox.showerror("Fel", f"Kunde inte lägga till data:\n{str(e)}")

    def reset_selection(self):
        """Rensar vald fil och alla kolumnmappningar"""
        self.current_file = None
        self.current_columns = []
        self.file_label.config(text="Ingen fil vald", foreground="gray")

        # Rensa alla comboboxes
        for combo in self.column_combos.values():
            combo.set('')
            combo['values'] = []
            combo['state'] = 'disabled'

        # Inaktivera lägg till-knappen
        self.add_data_btn['state'] = 'disabled'

    def show_imported_files(self):
        """Visar en popup med alla importerade filer"""
        if not self.imported_files_info:
            messagebox.showinfo("Importerade filer", "Inga filer har importerats ännu.")
            return

        # Skapa popup-fönster
        popup = tk.Toplevel(self.root)
        popup.title("Importerade filer")
        popup.geometry("500x300")
        popup.transient(self.root)
        popup.grab_set()

        # Färger från config
        primary_color = UI_SETTINGS.get('primary_color', '#00A0DC')
        bg_color = UI_SETTINGS.get('background_color', '#FFFFFF')
        popup.configure(bg=bg_color)

        # Rubrik
        tk.Label(popup, text="Importerade filer", font=('Arial', 14, 'bold'),
                fg=primary_color, bg=bg_color).pack(pady=10)

        # Ram för listan
        list_frame = tk.Frame(popup, bg=bg_color)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # Listbox med scrollbar
        listbox = tk.Listbox(list_frame, font=('Arial', 10), height=10)
        listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        listbox.config(yscrollcommand=scrollbar.set)

        # Fyll listan
        for file_info in self.imported_files_info:
            listbox.insert(tk.END, file_info)

        # Stäng-knapp
        ttk.Button(popup, text="Stäng", command=popup.destroy).pack(pady=10)

    def confirm_reset_all(self):
        """Visar bekräftelsedialog innan allt raderas"""
        # Kolla om det finns något att radera
        if not os.path.exists(self.summary_file) and not self.imported_files_info:
            messagebox.showinfo("Info", "Det finns ingen summeringsfil eller importerade filer att ta bort.")
            return

        # Bekräftelsedialog
        result = messagebox.askyesno(
            "Bekräfta - Börja om",
            "⚠️ Är du säker på att du vill börja om?\n\n"
            "Detta kommer att:\n"
            "• Ta bort summeringsfilen\n"
            "• Rensa listan över importerade filer\n\n"
            "Denna åtgärd kan inte ångras!",
            icon='warning'
        )

        if result:
            self.reset_all()

    def reset_all(self):
        """Tar bort summeringsfilen och rensar allt"""
        # Ta bort summeringsfilen om den finns
        if os.path.exists(self.summary_file):
            os.remove(self.summary_file)

        # Rensa importerade filer
        self.imported_files.clear()
        self.imported_files_info.clear()

        # Uppdatera knapptext
        self.show_imported_btn.config(text="Visa importerade filer (0)")

        # Rensa nuvarande val
        self.reset_selection()

        messagebox.showinfo("Klart!", "Summeringsfilen har tagits bort och allt har återställts.\n\nDu kan nu börja om!")


if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelMergerApp(root)
    root.mainloop()

