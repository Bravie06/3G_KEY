import pandas as pd
from data_processor import DataProcessor

class KPIMatcher:
    # Mapping defined as: Output KPI Name -> Raw Data Column Name
    # Based on the user constraints and template files
    KPI_MAPPING = {
        'Average of ORA_3G_Cell Availability (%)': 'ORA_3G_Cell Availability, excluding BLU (%)',
        'Average of OCM_3G_CSSR_CS_CellPCH_URAPCH_New(%)': 'ORA_3G_CSSR CS with Cell PCH/URA PCH (%)',
        'Average of OCM_3G_CSSR_PS_CellPCH_URAPCH_New(%)': 'ORA_3G_CSSR_PS_with_Cell_PCH (%)',
        'Average of OCM_3G_Call_Drop_CS_New(%)': 'ORA_3G_Drop Call Rate CS(%)',
        'Average of ORA_3G_Call_Drop_PS_All_Data_Services_New(%)': 'ORA_3G_Call Drop Data All Services(%)',
        'Sum of ORA_3G_CS Voice Traffic (Erl)': 'ORA_3G_CS Voice Traffic (Erl)',
        'Sum of ORA_3G_Traffic Total Data_MAC (GB)': 'ORA_3G_Traffic Total Data_MAC (GB)'
    }

    def __init__(self, filtered_df):
        self.filtered_df = filtered_df
        self.nodes = None
        self.times = None
        self.shaped_data = {}

    def process(self):
        """
        Transforms the filtered flat dataframe into a structured format
        (dict of nodes -> dict of KPIs -> dict of times -> values).
        """
        # Ensure we're sorted correctly
        df = self.filtered_df.sort_values(by=['NodeB Name', 'Begin Time'])

        # Get unique nodes and times
        self.nodes = df['NodeB Name'].dropna().unique().tolist()
        self.times = sorted(df['Begin Time'].dropna().unique().tolist())

        # Build shaped structure
        for node in self.nodes:
            node_df = df[df['NodeB Name'] == node]
            self.shaped_data[node] = {}
            for kpi_out, kpi_raw in self.KPI_MAPPING.items():
                self.shaped_data[node][kpi_out] = {}
                if kpi_raw in node_df.columns:
                    for t in self.times:
                        # Find the value for this specific time
                        row = node_df[node_df['Begin Time'] == t]
                        if not row.empty:
                            val = row[kpi_raw].values[0]
                            # If nan, fallback to None
                            # Scale percentages back to 100 for proper display if needed, some raw data uses 1.0 for 100%
                            if pd.notna(val):
                                # Many raw percentages might be in 0-1 scale, but according to user prompt:
                                # "availability >= 98.5" which implies 100-scale is expected for formatting.
                                # Let's multiply by 100 if it's less than or equal to 1.0 and it's a percentage KPI.
                                if "(%)" in kpi_out and abs(val) <= 1.0:
                                    val = val * 100
                                self.shaped_data[node][kpi_out][t] = val
                            else:
                                self.shaped_data[node][kpi_out][t] = None
                        else:
                            self.shaped_data[node][kpi_out][t] = None
                else:
                    # Column not found in raw data
                    for t in self.times:
                        self.shaped_data[node][kpi_out][t] = None

        return self.shaped_data, self.times, list(self.KPI_MAPPING.keys())

if __name__ == "__main__":
    processor = DataProcessor('Performance Management-History Query-3G_KPI_Reporting_Template-DFBG6870-20260309082244.xlsx')
    df = processor.filter_last_10_hours()

    matcher = KPIMatcher(df)
    shaped, times, kpis = matcher.process()
    print("Times mapped:", times)
    print("KPIs mapped:", kpis)
    sample_node = list(shaped.keys())[0]
    print(f"Sample data for {sample_node}:")
    for kpi, values in shaped[sample_node].items():
        print(f"  {kpi}: {list(values.values())}")
