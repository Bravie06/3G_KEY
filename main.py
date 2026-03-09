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

        # Output Directory Selection
        frame_out = tk.Frame(self.root)
        frame_out.pack(fill='x', padx=20, pady=5)
        tk.Label(frame_out, text="Output Folder:").pack(side='left')
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
        dirpath = filedialog.askdirectory()
        if dirpath:
            self.output_dir.set(dirpath)

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
            messagebox.showerror("Error", "Please select all files and output directory.")
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
            # Using groupby since Pivot might be complex with multiple value columns
            # Actually, let's keep it simple: create a nested dictionary for fast lookup
            # data_dict[node][begin_time][raw_kpi_name] = value
            data_dict = {}
            for _, row in df_filtered.iterrows():
                node = str(row['NodeB Name'])
                t = row['Begin Time']
                if node not in data_dict:
                    data_dict[node] = {}
                if t not in data_dict[node]:
                    data_dict[node][t] = {}

                for kpi in all_mapped_kpis:
                    if kpi in row and pd.notna(row[kpi]):
                        data_dict[node][t][kpi] = row[kpi]

            # Load template using openpyxl
            wb_template = openpyxl.load_workbook(template)

            # The template is on 'Sheet1' and 'Sheet2'
            # based on analysis, we should use 'Sheet2' since it's cleaner or just process the active sheet
            # Actually, looking at the template, Sheet2 contains the data we need to overwrite.
            # But let's process all sheets that have the same structure (e.g. Row Labels at A2)
            for sheet_name in wb_template.sheetnames:
                ws = wb_template[sheet_name]

                # We need to find the "Row Labels" cell
                start_row = None
                for r in range(1, 20):
                    cell_val = ws.cell(row=r, column=1).value
                    if isinstance(cell_val, str) and "Row Labels" in cell_val:
                        start_row = r
                        break

                if not start_row:
                    continue # Skip sheets that don't have this structure

                # Write the 10 time columns in the header row
                for idx, t in enumerate(last_10_times):
                    # In template, times start at column 2 (B)
                    ws.cell(row=start_row, column=idx+2).value = str(t)

                # Now clear any existing times beyond the 10th if any
                for c in range(10 + 2, 20):
                    if ws.cell(row=start_row, column=c).value:
                        ws.cell(row=start_row, column=c).value = ""

                # Iterate rows below the start_row to fill in data
                current_node = None

                # Colors based on rules
                # green: 00B050 (hex) -> let's use '00FF00' or whatever is standard green, maybe openpyxl likes AARRGGBB
                GREEN_FILL = PatternFill(start_color='FF00B050', end_color='FF00B050', fill_type='solid')
                RED_FILL = PatternFill(start_color='FFFF0000', end_color='FFFF0000', fill_type='solid') # or faded red: 'FFFFC7CE'
                FADED_RED_FILL = PatternFill(start_color='FFFFC7CE', end_color='FFFFC7CE', fill_type='solid')

                for r in range(start_row + 1, ws.max_row + 1):
                    kpi_cell_val = ws.cell(row=r, column=1).value
                    if not kpi_cell_val:
                        continue

                    kpi_cell_str = str(kpi_cell_val)

                    # Check if this row is a NodeB Name or a KPI
                    # Nodes in template start with LIT_ or CTR_ generally, but safely:
                    # if it doesn't contain KPI keywords, it's a node
                    if self.map_kpi(kpi_cell_str) is None:
                        current_node = kpi_cell_str
                        # clear out data cells for node row if any
                        for c in range(2, 12):
                            ws.cell(row=r, column=c).value = ""
                            ws.cell(row=r, column=c).fill = PatternFill(fill_type=None)
                        continue

                    if current_node and current_node in data_dict:
                        raw_kpi = self.map_kpi(kpi_cell_str)
                        if raw_kpi:
                            for idx, t in enumerate(last_10_times):
                                cell = ws.cell(row=r, column=idx+2)

                                # Clear existing styling first
                                cell.fill = PatternFill(fill_type=None)

                                val = data_dict[current_node].get(t, {}).get(raw_kpi)
                                if val is not None:
                                    cell.value = val

                                    # Apply color logic
                                    # Color logic
                                    check_val = val
                                    if check_val <= 1.0 and ('cssr' in kpi_cell_str.lower() or 'availability' in kpi_cell_str.lower()):
                                        check_val *= 100

                                    kpi_lower = kpi_cell_str.lower()
                                    if 'availability' in kpi_lower:
                                        if check_val >= 98.5:
                                            cell.fill = GREEN_FILL
                                    elif 'cssr' in kpi_lower:
                                        if check_val < 98.5:
                                            cell.fill = FADED_RED_FILL
                                    elif 'call_drop' in kpi_lower:
                                        drop_val = val
                                        if drop_val < 1.0 and drop_val > 0:
                                            drop_val *= 100
                                        if drop_val <= 0.7:
                                            cell.fill = GREEN_FILL
                                    elif 'traffic' in kpi_lower:
                                        # Traffic always white, meaning no color fill
                                        pass
                                else:
                                    cell.value = ""

                                # If the data isn't in raw_data, we might just be keeping the template values,
                                # but the requirement says "only containing the values of the KPI from the raw data
                                # having their corresponding in the template for the last 10 hours".
                                # If val is None and we didn't explicitly set it, we should clear it if it's the 10 hours timeframe.
                                if val is None:
                                    cell.value = ""

                # For nodes present in raw_data but NOT in template, we should probably add them
                # The requirements: "le fichier generer doit contenir uniquement les valeurs des KPI du raw data ayant leur correspondant dant le template pour les 10 dernieres heures uniquement"
                # This could mean:
                # Option A: Only nodes present in the template should be updated with raw data.
                # Option B: Nodes present in the raw data should be appended to the template if they are missing.
                # "le fichier generer doit contenir uniquement les valeurs des KPI du raw data ayant leur correspondant dant le template"
                # Actually, I will add missing nodes from the raw data that are not in the template.

                # First, collect nodes already in template
                template_nodes = set()
                for r in range(start_row + 1, ws.max_row + 1):
                    val = str(ws.cell(row=r, column=1).value or "")
                    if val and self.map_kpi(val) is None and "Grand Total" not in val:
                        template_nodes.add(val)

                # Then append missing nodes from data_dict
                next_row = ws.max_row + 1
                for node in data_dict.keys():
                    if node not in template_nodes:
                        # Append this node and its mapped KPIs
                        ws.cell(row=next_row, column=1).value = node
                        next_row += 1

                        # Add KPIs
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

                            for idx, t in enumerate(last_10_times):
                                cell = ws.cell(row=next_row, column=idx+2)
                                val = data_dict[node].get(t, {}).get(raw_kpi)
                                if val is not None:
                                    cell.value = val

                                    # Color logic
                                    # Need to multiply by 100 for some KPIs if they are fractions (e.g. 0.9989 -> 99.89%)
                                    # Let's adjust based on rules and typical val
                                    check_val = val
                                    if check_val <= 1.0 and ('cssr' in kpi_label.lower() or 'availability' in kpi_label.lower()):
                                        check_val *= 100

                                    kpi_lower = kpi_label.lower()
                                    if 'availability' in kpi_lower:
                                        if check_val >= 98.5:
                                            cell.fill = GREEN_FILL
                                    elif 'cssr' in kpi_lower:
                                        if check_val < 98.5:
                                            cell.fill = FADED_RED_FILL
                                    elif 'call_drop' in kpi_lower:
                                        # For call drop, sometimes it's 0.001 which is 0.1%
                                        # The rule says <= 0.7. If the val is 0.001, 0.001 <= 0.7 is true.
                                        # However, if it's stored as fraction, 0.001 * 100 = 0.1 <= 0.7
                                        # Just use val * 100 for comparison if val < 1 to be safe, but wait,
                                        # if it is 0.001, and we check <= 0.7, it works either way.
                                        # But let's be consistent with template scale.
                                        drop_val = val
                                        if drop_val < 1.0 and drop_val > 0:
                                            drop_val *= 100
                                        if drop_val <= 0.7:
                                            cell.fill = GREEN_FILL

                            next_row += 1

            # Remove conditional formatting rules that mess up our manual coloring
            for sheet_name in wb_template.sheetnames:
                ws = wb_template[sheet_name]
                ws.conditional_formatting._cf_rules = {}

            # Save the file
            out_file = os.path.join(out_dir, "Generated_KPI_Report.xlsx")
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
