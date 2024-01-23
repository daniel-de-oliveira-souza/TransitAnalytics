import os
import glob
import json
import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt
import matplotlib.ticker as mtick
from datetime import datetime, timedelta
import pandas as pd

class DataProcessor:
    def __init__(self, directory):
        self.directory = directory

    @staticmethod
    def add_week_number_column_from_filename(df, week_number):
        """Add WeekNumber Column to the DataFrame based on provided week number."""
        df['WeekNumber'] = week_number[-7:-5]
        return df

    @staticmethod
    def add_month_column_from_filename(df, filename):
        """Add Month Column to the DataFrame based on Filename last word preceding .xlsx."""
        base_filename = os.path.splitext(os.path.basename(filename))[0]
        words = base_filename.split("_")
        if len(words) > 1:
            month = words[1]
            df['Month'] = month
        return df

    @staticmethod
    def add_date_column_from_filename(df, week_number):
        """Add Date Column (always Mondays) to the DataFrame based on the MTA week number."""
        reference_monday = datetime.strptime("2022-12-26", "%Y-%m-%d")
        monday_date = reference_monday + timedelta(days=(week_number) * 7)
        df['Date'] = monday_date.date()
        return df

    def get_xlsx_files(self):
        """Get all .xlsx files in the directory."""
        xlsx_files = [os.path.join(self.directory, f) for f in os.listdir(self.directory) if f.endswith('.xlsx')]
        return xlsx_files

    def read_xlsx_file(self, file_path):
        """Read an Excel file and return a DataFrame."""
        try:
            df = pd.read_excel(file_path)
            return df
        except Exception as e:
            print(f"Error reading {file_path}: {e}")
            return None

    def process_xlsx_file(self, file_path):
        """Process an Excel file and return a DataFrame."""
        df = self.read_xlsx_file(file_path)
        if df is not None:
            if 'LIRR' in file_path:
                df = DataProcessor.add_month_column_from_filename(df, file_path)
                df = df.rename(columns={'% StkOut': '%StockOut'})

                # Mapping of full month names to numerical values
                month_mapping = {'January': 1, 'February': 2, 'March': 3, 'April': 4, 'May': 5, 'June': 6,
                    'July': 7, 'August': 8, 'September': 9, 'October': 10, 'November': 11, 'December': 12}
                # Map the month names to numbers
                df['Month_Num'] = df['Month'].map(month_mapping)

            else:
                week_number = int(file_path[-7:-5])
                df = DataProcessor.add_week_number_column_from_filename(df, file_path)
                df = DataProcessor.add_date_column_from_filename(df, week_number)
            # Remove the last three rows of each file
            df = df.iloc[:-3]
            df = df.rename(columns={'% StockOut': '%StockOut'})
            if len(df) > 6:
                df.columns = df.columns.str.replace('Sys StkOut', 'Sys_StkOut')
        return df

    def process_xlsx_files(self):
        """Process all Excel files in the directory."""
        dataframes = [self.process_xlsx_file(file) for file in self.get_xlsx_files()]
        concatenated_df = pd.concat(dataframes, ignore_index=True)
        return concatenated_df

class DataVisualizer:
    def __init__(self, df):
        self.df = df

    def create_bar_plot(self, df, srdescription_or_whname, monthnumber_or_weeknumber, wh_or_sr):
        """Create and save a bar plot.

        Parameters:
        df (DataFrame): The DataFrame containing the data to plot.
        srdescription_or_whname (str): The column name for the x-axis.
        month_or_weeknumber (str): The column name for the hue.
        wh_or_sr (str): Indicator for warehouse ('WH') or storeroom ('SR').

        Returns:
        str: The file path of the saved plot.
        """
        # Set Seaborn style
        sns.set(style="darkgrid", palette="viridis")

        # Increase the figure size
        plt.figure(figsize=(18, 10))

        # Filter and sort the DataFrame
        filtered_df = df[df[srdescription_or_whname] != "Total"]
        sorted_df = filtered_df.sort_values(by="%StockOut", ascending=False)

        # Create a bar plot using Seaborn
        plot = sns.barplot(
            x=srdescription_or_whname,
            y="%StockOut",
            hue=monthnumber_or_weeknumber,
            data=sorted_df
        )

        # Customize the plot
        plot.set_title('Percentage_StockOut_by_Warehouse' if wh_or_sr == 'WH' else 'Percentage_StockOut_by_Storeroom', fontsize=16)
        plot.set_xlabel('Warehouse' if wh_or_sr == 'WH' else 'Storeroom', fontsize=14)
        plot.set_xticklabels(plot.get_xticklabels(), rotation=45, horizontalalignment='right', fontsize=10)

        # Save the Seaborn plot to an image
        plot1_filename = 'Percentage_StockOut_by_Warehouse.png' if wh_or_sr == 'WH' else 'Percentage_StockOut_by_Storeroom.png'
        plot1_path = os.path.join('/workspaces/transitanalytics/static/images', plot1_filename)
        plt.savefig(plot1_path)

        return plot1_path

    def create_line_plot(self, totals_df_reset, monthnumber_or_weeknumber):
        """Create and save a line plot of average stock out change over time.

        Parameters:
        totals_df (DataFrame): The DataFrame containing stock out data.
        month_or_weeknumber (str): The column name representing the month or week number.

        Returns:
        str: The file path of the saved plot, or None if the plot is not created.
        """

        try:
            plt.figure(figsize=(15, 7))
            sns.lineplot(x=totals_df_reset.index.get_level_values(0), y='Avg_%StockOut_Change', data=totals_df_reset, marker='o', color='teal', label='Average % StockOut Change')

            plt.title('Average % StockOut Change Over Time', fontsize=16, fontweight='bold')
            plt.xlabel(monthnumber_or_weeknumber, fontsize=14)
            plt.ylabel('Average % StockOut Change', fontsize=14)
            plt.gca().yaxis.set_major_formatter(mtick.PercentFormatter(1.0))
            plt.xticks(rotation=45, fontsize=12)
            plt.yticks(fontsize=12)
            plt.legend(fontsize='12')
            plt.tight_layout()

            plot2_filename = 'avg_stockout_change_plot.png'
            plot2_path = os.path.join('/workspaces/transitanalytics/static/images', plot2_filename)
            plt.savefig(plot2_path)
            return plot2_path
        except KeyError as e:
            print(f"Error creating plot: {e}")
            return None

class JsonToMarkdownConverter:
    def __init__(self, json_dir, markdown_dir):
        self.json_dir = json_dir
        self.markdown_dir = markdown_dir

    def convert(self):
        for json_file in os.listdir(self.json_dir):
            if json_file.endswith('.json'):
                json_path = os.path.join(self.json_dir, json_file)
                markdown_file = json_file.replace('.json', '.md')
                markdown_path = os.path.join(self.markdown_dir, markdown_file)
                self.json_to_markdown(json_path, markdown_path)
        print("Conversion complete!")

    @staticmethod
    def json_to_markdown(json_file, output_file):
        with open(json_file, 'r') as file:
            json_data = json.load(file)

        markdown = ''
        # Add front matter
        title = os.path.basename(json_file).replace('.json', '').replace('_', ' ').title()
        date = datetime.now().strftime("%Y-%m-%d")
        markdown += f'---\ntitle: "{title}"\ndate: {date}\n---\n\n'

        # Add summary with Read More link
        markdown += f'## {title}\n*{date}*\n\n'
        markdown += '[Read More](#full-content)\n\n'

        # Add anchor for full content
        markdown += '<div id="full-content"></div>\n\n'

        # Generate table from JSON data
        if json_data:
            # Create table headers
            headers = json_data['columns']
            markdown += '| ' + ' | '.join(headers) + ' |\n'
            markdown += '| ' + ' | '.join(['---'] * len(headers)) + ' |\n'

            # Create table rows
            for row in json_data['data']:
                markdown += '| ' + ' | '.join(str(cell) for cell in row) + ' |\n'

        with open(output_file, 'w') as file:
            file.write(markdown)