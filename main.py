import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from datetime import datetime
import openpyxl
from openpyxl.styles import PatternFill
import os
from pathlib import Path

class KPIApp:
    def __init__(self, root):
        self.root = root
        self.root.title("3G KPI Report Generator")
        self.root.geometry("600x400")

        self.setup_ui()

    def setup_ui(self):
        # Frame for Event Name
        event_frame = ttk.Frame(self.root)
        event_frame.pack(pady=20, padx=20, fill='x')

        ttk.Label(event_frame, text="Event Name:").pack(side='left', padx=5)
        self.event_entry = ttk.Entry(event_frame, width=30)
        self.event_entry.pack(side='left', padx=5)

        # Frame for File Selection
        file_frame = ttk.Frame(self.root)
        file_frame.pack(pady=20, padx=20, fill='x')

        self.file_label = ttk.Label(file_frame, text="No file selected", foreground="gray")
        self.file_label.pack(side='bottom', pady=5)

        self.btn_select = ttk.Button(file_frame, text="Select Raw Data File", command=self.select_file)
        self.btn_select.pack(side='top')

        self.raw_file_path = None

        # Action Button
        self.btn_generate = ttk.Button(self.root, text="Generate Report", command=self.run_generation)
        self.btn_generate.pack(pady=30)

    def select_file(self):
        file_path = filedialog.askopenfilename(
            title="Select Raw Data File",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if file_path:
            self.raw_file_path = file_path
            self.file_label.config(text=file_path.split('/')[-1], foreground="black")

    def run_generation(self):
        event_name = self.event_entry.get().strip()

        if not event_name:
            messagebox.showwarning("Warning", "Please enter an event name.")
            return

        if not self.raw_file_path:
            messagebox.showwarning("Warning", "Please select a raw data file.")
            return

        try:
            # Process Data
            unique_hours, report_data = self.process_data(self.raw_file_path)

            # Generate Excel
            output_file = self.generate_excel(unique_hours, report_data, event_name)

            messagebox.showinfo("Success", f"Report generated successfully:\n{output_file}")

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred:\n{str(e)}")

    def process_data(self, raw_file_path):
        # Mappings from template KPI names to Raw Data KPI names
        kpi_mapping = {
            'Average of ORA_3G_Cell Availability (%)': 'ORA_3G_Cell Availability, excluding BLU (%)',
            'Average of OCM_3G_CSSR_CS_CellPCH_URAPCH_New(%)': 'ORA_3G_CSSR CS with Cell PCH/URA PCH (%)',
            'Average of OCM_3G_CSSR_PS_CellPCH_URAPCH_New(%)': 'ORA_3G_CSSR_PS_with_Cell_PCH (%)',
            'Average of OCM_3G_Call_Drop_CS_New(%)': 'ORA_3G_Drop Call Rate CS(%)',
            'Average of ORA_3G_Call_Drop_PS_All_Data_Services_New(%)': 'ORA_3G_Call Drop Data All Services(%)',
            'Sum of ORA_3G_CS Voice Traffic (Erl)': 'ORA_3G_CS Voice Traffic (Erl)',
            'Sum of ORA_3G_Traffic Total Data_MAC (GB)': 'ORA_3G_Traffic Total Data_MAC (GB)'
        }

        # Read the raw data
        try:
            df = pd.read_excel(raw_file_path)
        except Exception as e:
            raise Exception(f"Error reading raw file: {str(e)}")

        # Ensure we have the necessary columns
        if 'Begin Time' not in df.columns:
            raise Exception("The raw file must contain the 'Begin Time' column.")

        # Determine the entity column (NodeB Name for sites, Group or similar for cities)
        entity_col = None
        for col in ['NodeB Name', 'Group', 'City', 'Site Name', 'Nom du Site', 'Ville']:
            if col in df.columns:
                entity_col = col
                break

        if not entity_col:
            # Fallback to the first string column after time
            for col in df.columns:
                if col not in ['Begin Time', 'End Time', 'Granularity', 'SubnetWork ID', 'SubnetWork Name', 'ManagedElement ID', 'RNC Managed NE', 'NodeB ID']:
                    if df[col].dtype == object and df[col].notna().any() and isinstance(df[col].dropna().iloc[0], str):
                        entity_col = col
                        break

        if not entity_col:
            raise Exception("Unable to find the entity column (Site name or City).")

        # Convert Begin Time to datetime to easily sort and find the last 10 hours
        df['Begin Time'] = pd.to_datetime(df['Begin Time'])

        # Get the unique 10 most recent hours
        unique_hours = sorted(df['Begin Time'].dropna().unique())[-10:]

        # Filter dataframe for only these last 10 hours
        df_filtered = df[df['Begin Time'].isin(unique_hours)].copy()

        # Make sure values are numeric where possible, if not fill with 0
        for raw_kpi in kpi_mapping.values():
            if raw_kpi in df_filtered.columns:
                df_filtered[raw_kpi] = pd.to_numeric(df_filtered[raw_kpi], errors='coerce').fillna(0)

        # Build the final structured data for our report

        report_data = []
        entities = sorted(df_filtered[entity_col].dropna().unique())

        for entity in entities:
            df_entity = df_filtered[df_filtered[entity_col] == entity]

            # First row for the Entity is just its name in the first column
            report_data.append({
                'KPI': entity,
                **{hour: None for hour in unique_hours}
            })

            # Then one row for each KPI
            for template_kpi, raw_kpi in kpi_mapping.items():
                kpi_row = {'KPI': template_kpi}
                for hour in unique_hours:
                    # Find the value for this specific hour
                    val_df = df_entity[df_entity['Begin Time'] == hour]
                    if not val_df.empty and raw_kpi in val_df.columns:
                        kpi_row[hour] = val_df.iloc[0][raw_kpi]
                    else:
                        kpi_row[hour] = None
                report_data.append(kpi_row)

        return unique_hours, report_data

    def generate_excel(self, unique_hours, report_data, event_name):
        # Determine the filename
        # Format KEA_3G_nameOfEvent_dateOfDay_hours
        now = datetime.now()
        date_of_day = now.strftime("%Y%m%d")
        hours_str = now.strftime("%H%M%S")
        output_filename = f"KEA_3G_{event_name}_{date_of_day}_{hours_str}.xlsx"

        # Save in the user's Downloads folder
        downloads_path = str(Path.home() / "Downloads" / output_filename)

        # Create workbook and worksheet
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Report"

        # Write header
        # Row 3 is "Row Labels" and the hours
        ws.cell(row=3, column=1, value="Row Labels")
        for i, hour in enumerate(unique_hours):
            ws.cell(row=3, column=2+i, value=hour)

        # Write data
        current_row = 4

        green_fill = PatternFill(start_color="FFC6EFCE", end_color="FFC6EFCE", fill_type="solid")
        red_fill = PatternFill(start_color="FFFFC7CE", end_color="FFFFC7CE", fill_type="solid")

        for row_dict in report_data:
            kpi_name = row_dict['KPI']
            ws.cell(row=current_row, column=1, value=kpi_name)

            # If this is not a KPI row (it's a NodeB Name row)
            if kpi_name not in [
                'Average of ORA_3G_Cell Availability (%)',
                'Average of OCM_3G_CSSR_CS_CellPCH_URAPCH_New(%)',
                'Average of OCM_3G_CSSR_PS_CellPCH_URAPCH_New(%)',
                'Average of OCM_3G_Call_Drop_CS_New(%)',
                'Average of ORA_3G_Call_Drop_PS_All_Data_Services_New(%)',
                'Sum of ORA_3G_CS Voice Traffic (Erl)',
                'Sum of ORA_3G_Traffic Total Data_MAC (GB)'
            ]:
                # Apply bold to nodeB name (optional but good)
                ws.cell(row=current_row, column=1).font = openpyxl.styles.Font(bold=True)
                current_row += 1
                continue

            for i, hour in enumerate(unique_hours):
                val = row_dict[hour]
                cell = ws.cell(row=current_row, column=2+i, value=val)

                if val is None:
                    continue

                # Apply conditional formatting
                if kpi_name == 'Average of ORA_3G_Cell Availability (%)':
                    if float(val) >= 98.5:
                        cell.fill = green_fill
                    else:
                        cell.fill = red_fill

                elif kpi_name in ['Average of OCM_3G_CSSR_CS_CellPCH_URAPCH_New(%)', 'Average of OCM_3G_CSSR_PS_CellPCH_URAPCH_New(%)']:
                    if float(val) >= 98.5:
                        cell.fill = green_fill
                    else:
                        cell.fill = red_fill

                elif kpi_name in ['Average of OCM_3G_Call_Drop_CS_New(%)', 'Average of ORA_3G_Call_Drop_PS_All_Data_Services_New(%)']:
                    if float(val) <= 0.7:
                        cell.fill = green_fill
                    else:
                        cell.fill = red_fill

                # Traffic CS / Traffic Data don't get colored (always white/no fill)

            current_row += 1

        # Auto-size the first column
        ws.column_dimensions['A'].width = 60

        # Save workbook
        wb.save(downloads_path)
        return downloads_path

if __name__ == "__main__":
    root = tk.Tk()
    app = KPIApp(root)
    root.mainloop()
