import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill
import os

class KPIAutomatorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("KPI Report Automator")
        self.root.geometry("600x400")

        # Variables
        self.template_path = tk.StringVar()
        self.raw_data_path = tk.StringVar()
        self.output_dir = tk.StringVar()

        # UI Elements
        self.create_widgets()

    def create_widgets(self):
        # Title
        title_label = tk.Label(self.root, text="KPI Report Automator", font=("Helvetica", 16, "bold"))
        title_label.pack(pady=10)

        # Template File Selection
        frame_template = tk.Frame(self.root)
        frame_template.pack(fill='x', padx=20, pady=5)
        tk.Label(frame_template, text="Template File:").pack(side='left')
        tk.Entry(frame_template, textvariable=self.template_path, width=50).pack(side='left', padx=10)
        tk.Button(frame_template, text="Browse", command=self.browse_template).pack(side='left')

        # Raw Data File Selection
        frame_raw = tk.Frame(self.root)
        frame_raw.pack(fill='x', padx=20, pady=5)
        tk.Label(frame_raw, text="Raw Data File:").pack(side='left')
        tk.Entry(frame_raw, textvariable=self.raw_data_path, width=50).pack(side='left', padx=10)
        tk.Button(frame_raw, text="Browse", command=self.browse_raw).pack(side='left')

        # Output File Selection
        frame_out = tk.Frame(self.root)
        frame_out.pack(fill='x', padx=20, pady=5)
        tk.Label(frame_out, text="Save As:").pack(side='left')
        tk.Entry(frame_out, textvariable=self.output_dir, width=50).pack(side='left', padx=10)
        tk.Button(frame_out, text="Browse", command=self.browse_output).pack(side='left')

        # Generate Button
        self.generate_btn = tk.Button(self.root, text="Generate Excel Report", font=("Helvetica", 12, "bold"), bg="#4CAF50", fg="white", command=self.generate_report)
        self.generate_btn.pack(pady=30)

        # Status Label
        self.status_label = tk.Label(self.root, text="Ready", fg="blue")
        self.status_label.pack(side="bottom", pady=10)

    def browse_template(self):
        filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if filepath:
            self.template_path.set(filepath)

    def browse_raw(self):
        filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if filepath:
            self.raw_data_path.set(filepath)

    def browse_output(self):
        filepath = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx *.xls")],
            initialfile="Generated_KPI_Report.xlsx"
        )
        if filepath:
            self.output_dir.set(filepath)

    def map_kpi(self, kpi_name):
        kpi_lower = kpi_name.lower()
        if 'availability' in kpi_lower:
            return 'ORA_3G_Cell Availability, excluding BLU (%)'
        elif 'cssr_cs' in kpi_lower:
            return 'ORA_3G_CSSR CS with Cell PCH/URA PCH (%)'
        elif 'cssr_ps' in kpi_lower:
            return 'ORA_3G_CSSR_PS_with_Cell_PCH (%)'
        elif 'call_drop_cs' in kpi_lower:
            return 'ORA_3G_Drop Call Rate CS(%)'
        elif 'call_drop_ps' in kpi_lower:
            return 'ORA_3G_Call Drop Data All Services(%)'
        elif 'voice traffic' in kpi_lower:
            return 'ORA_3G_CS Voice Traffic (Erl)'
        elif 'traffic total data' in kpi_lower:
            return 'ORA_3G_Traffic Total Data_MAC (GB)'
        return None

    def generate_report(self):
        template = self.template_path.get()
        raw_data = self.raw_data_path.get()
        out_dir = self.output_dir.get()

        if not template or not raw_data or not out_dir:
            messagebox.showerror("Error", "Please select all files and the save destination.")
            return

        self.status_label.config(text="Processing data...", fg="orange")
        self.root.update_idletasks()

        try:
            # Load raw data
            df_raw = pd.read_excel(raw_data, sheet_name=0)

            # Find the 10 most recent unique Begin Time values
            unique_times = sorted(df_raw['Begin Time'].dropna().unique())
            last_10_times = unique_times[-10:] if len(unique_times) >= 10 else unique_times

            # Filter raw data for these 10 hours
            df_filtered = df_raw[df_raw['Begin Time'].isin(last_10_times)].copy()

            # Keep only necessary columns based on KPI mapping
            needed_columns = ['NodeB Name', 'Begin Time']
            all_mapped_kpis = [
                'ORA_3G_Cell Availability, excluding BLU (%)',
                'ORA_3G_CSSR CS with Cell PCH/URA PCH (%)',
                'ORA_3G_CSSR_PS_with_Cell_PCH (%)',
                'ORA_3G_Drop Call Rate CS(%)',
                'ORA_3G_Call Drop Data All Services(%)',
                'ORA_3G_CS Voice Traffic (Erl)',
                'ORA_3G_Traffic Total Data_MAC (GB)'
            ]
            for col in all_mapped_kpis:
                if col in df_filtered.columns:
                    needed_columns.append(col)

            df_filtered = df_filtered[needed_columns]

            # Pivot data so we can easily query it: NodeB Name, KPI -> [Time1, Time2, ...]
            # data_dict[node][begin_time][raw_kpi_name] = value
            data_dict = {}
            for _, row in df_filtered.iterrows():
                node = str(row['NodeB Name'])
                t = str(row['Begin Time'])
                if node not in data_dict:
                    data_dict[node] = {}
                if t not in data_dict[node]:
                    data_dict[node][t] = {}

                for kpi in all_mapped_kpis:
                    if kpi in row and pd.notna(row[kpi]):
                        data_dict[node][t][kpi] = row[kpi]

            # Convert times to string for exact matching
            last_10_times_str = [str(t) for t in last_10_times]

            # Load template using openpyxl
            wb_template = openpyxl.load_workbook(template)

            # The requirement specifies the result should be in 1 single sheet.
            # We will use "Sheet2" or the first one having "Row Labels"
            # delete all other sheets.
            target_sheet_name = None
            # Preference is to use a sheet named 'Sheet2' as it is formatted cleaner in the template
            if 'Sheet2' in wb_template.sheetnames:
                target_sheet_name = 'Sheet2'
                ws = wb_template[target_sheet_name]
                for r in range(1, 20):
                    if str(ws.cell(row=r, column=1).value) == "Row Labels":
                        start_row = r
                        break
            else:
                for sheet_name in wb_template.sheetnames:
                    ws = wb_template[sheet_name]
                    for r in range(1, 20):
                        if str(ws.cell(row=r, column=1).value) == "Row Labels":
                            target_sheet_name = sheet_name
                            start_row = r
                            break
                    if target_sheet_name:
                        break

            if not target_sheet_name:
                raise ValueError("Could not find 'Row Labels' in the template.")

            # Delete all sheets except target_sheet
            for sheet_name in wb_template.sheetnames:
                if sheet_name != target_sheet_name:
                    del wb_template[sheet_name]

            ws = wb_template[target_sheet_name]
            ws.title = "KPI_Report" # Rename the single sheet

            # Remove conditional formatting as it overrides our colors
            ws.conditional_formatting._cf_rules = {}

            # Write the 10 time columns in the header row
            for idx, t in enumerate(last_10_times_str):
                # In template, times start at column 2 (B)
                ws.cell(row=start_row, column=idx+2).value = t

            # Clear any existing times beyond the 10th
            for c in range(10 + 2, 20):
                ws.cell(row=start_row, column=c).value = ""

            # Colors based on rules
            # The template conditional format uses FFFFC7CE for faded red (CSSR) and FF00B050 for green (Availability/Call Drop)
            # Actually, standard red in openpyxl might be FFFF0000.
            # Let's define them explicitly
            # Note: in openpyxl, start_color.index expects a hex code with aRRrrggbb, usually 8 chars.
            # If we give '00B050', openpyxl interprets it as 0000B050.
            # Let's use standard openpyxl named colors if we want or proper ARGB: FFFF0000 for pure red.
            # We'll use FFFF0000 for pure red or FFFFC7CE for faded red, FF00B050 for green.
            GREEN_FILL = PatternFill(start_color='FF00B050', end_color='FF00B050', fill_type='solid')
            FADED_RED_FILL = PatternFill(start_color='FFFFC7CE', end_color='FFFFC7CE', fill_type='solid')

            current_node = None

            # The phrase "contenir uniquement les valeurs des KPI du raw data ayant leur correspondant dant le template"
            # means we must ONLY use the KPI *variables* that are in the template, but we need to generate rows
            # for the actual nodes present in the raw data.

            # Step 1: Delete all existing rows below the header
            if ws.max_row > start_row:
                ws.delete_rows(start_row + 1, ws.max_row - start_row)

            # Step 2: Dynamically recreate the template structure for each node found in RAW DATA
            next_row = start_row + 1
            for node in data_dict.keys():
                # Write Node Name
                ws.cell(row=next_row, column=1).value = node
                next_row += 1

                # Write the corresponding KPIs and their values
                for kpi_label, raw_kpi in [
                    ('Average of ORA_3G_Cell Availability (%)', 'ORA_3G_Cell Availability, excluding BLU (%)'),
                    ('Average of OCM_3G_CSSR_CS_CellPCH_URAPCH_New(%)', 'ORA_3G_CSSR CS with Cell PCH/URA PCH (%)'),
                    ('Average of OCM_3G_CSSR_PS_CellPCH_URAPCH_New(%)', 'ORA_3G_CSSR_PS_with_Cell_PCH (%)'),
                    ('Average of OCM_3G_Call_Drop_CS_New(%)', 'ORA_3G_Drop Call Rate CS(%)'),
                    ('Average of ORA_3G_Call_Drop_PS_All_Data_Services_New(%)', 'ORA_3G_Call Drop Data All Services(%)'),
                    ('Sum of ORA_3G_CS Voice Traffic (Erl)', 'ORA_3G_CS Voice Traffic (Erl)'),
                    ('Sum of ORA_3G_Traffic Total Data_MAC (GB)', 'ORA_3G_Traffic Total Data_MAC (GB)')
                ]:
                    ws.cell(row=next_row, column=1).value = kpi_label

                    for idx, t in enumerate(last_10_times_str):
                        cell = ws.cell(row=next_row, column=idx+2)
                        val = data_dict[node].get(t, {}).get(raw_kpi)

                        if val is not None:
                            display_val = val
                            kpi_lower = kpi_label.lower()

                            if 'availability' in kpi_lower or 'cssr' in kpi_lower:
                                if display_val <= 1.0 and display_val > 0:
                                    display_val *= 100

                            # Note: Call drops are left untouched if they represent typical low percentages (e.g., 0.5)
                            # because multiplying by 100 could distort it. Color conditions work perfectly.

                            cell.value = display_val

                            if 'availability' in kpi_lower:
                                if display_val >= 98.5:
                                    cell.fill = GREEN_FILL
                            elif 'cssr' in kpi_lower:
                                if display_val < 98.5:
                                    # Use pure red since the user might want explicitly "rouge"
                                    # or faded red. The requirement states "rouge delave comme dans le template".
                                    # FFFFC7CE is faded red. FFFF0000 is pure red.
                                    # We used FFFFC7CE earlier. Let's make sure it's applied.
                                    cell.fill = FADED_RED_FILL
                            elif 'call_drop' in kpi_lower:
                                if display_val <= 0.7:
                                    cell.fill = GREEN_FILL

                    next_row += 1

            # Save the file
            out_file = out_dir
            wb_template.save(out_file)

            self.status_label.config(text=f"Success! Saved to: {out_file}", fg="green")
            messagebox.showinfo("Success", f"Report generated successfully!\nSaved to:\n{out_file}")

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")
            self.status_label.config(text="Failed", fg="red")

if __name__ == "__main__":
    root = tk.Tk()
    app = KPIAutomatorApp(root)
    root.mainloop()
