import pandas as pd
import warnings
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

class DataProcessor:
    def __init__(self, raw_data_path):
        self.raw_data_path = raw_data_path
        self.raw_df = None
        self.filtered_df = None

    def load_data(self):
        """Loads the raw data from the Excel file."""
        # Typically the main data is in 'Sheet0'
        try:
            self.raw_df = pd.read_excel(self.raw_data_path, sheet_name='Sheet0')
        except ValueError:
            # Fallback if 'Sheet0' does not exist
            self.raw_df = pd.read_excel(self.raw_data_path)

        # Convert 'Begin Time' to datetime
        if 'Begin Time' in self.raw_df.columns:
            self.raw_df['Begin Time'] = pd.to_datetime(self.raw_df['Begin Time'])
        else:
            raise ValueError("Le fichier raw data ne contient pas la colonne 'Begin Time'.")

    def filter_last_10_hours(self):
        """Filters the dataframe to keep only the last 10 hours based on the latest time in 'Begin Time'."""
        if self.raw_df is None:
            self.load_data()

        max_time = self.raw_df['Begin Time'].max()
        # The last 10 hours means the max_time and 9 hours before it.
        # However, let's take the 10 most recent distinct hours available.
        latest_hours = sorted(self.raw_df['Begin Time'].dropna().unique(), reverse=True)[:10]

        self.filtered_df = self.raw_df[self.raw_df['Begin Time'].isin(latest_hours)].copy()

        # Sort values
        self.filtered_df = self.filtered_df.sort_values(by=['NodeB Name', 'Begin Time'])
        return self.filtered_df

if __name__ == "__main__":
    processor = DataProcessor('Performance Management-History Query-3G_KPI_Reporting_Template-DFBG6870-20260309082244.xlsx')
    df = processor.filter_last_10_hours()
    print("Filtered distinct times:")
    print(df['Begin Time'].unique())
    print("Nodes count:", df['NodeB Name'].nunique())
    print("DataFrame shape:", df.shape)
