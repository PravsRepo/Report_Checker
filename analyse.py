import pandas as pd


class Analyzer:

    def __init__(self, file_path):
        self.file_path = file_path

    def read_file(self):
        df = pd.read_excel(self.file_path, sheet_name="Attendance-consolidated", header=1)
        return df

    def analyze_data(self, df):
        source_df = df.get(["Sno", "Name", "Total 100 marks"])
        max_mark = source_df["Total 100 marks"].max()
        min_mark = source_df["Total 100 marks"].min()
        return source_df, min_mark, max_mark

    def mark_falls(self, source_df):
        source_df_copy = source_df.copy()
        bin_values = [0, 10, 20, 30, 40, 50, 60, 70, 80, 90, 100]
        label_values = ["0 to 10", "11 to 20", "21 to 30", "31 to 40", "41 to 50", "51 to 60", "61 to 70", "71 to 80", "81 to 90", "91 to 100"]
        source_df_copy["Mark distribution"] = pd.cut(x=source_df_copy["Total 100 marks"], bins=bin_values, labels=label_values, include_lowest=True)
        bin_counts = source_df_copy["Mark distribution"].value_counts().sort_index()
        bin_details = bin_counts.to_frame(name="Count of Mark distribution")
        bin_percentage = source_df_copy["Mark distribution"].value_counts(normalize=True).sort_index()*100
        bin_details["Distribution Percentage"] = bin_percentage
        return bin_details
    

    def summary(self, output_file_path, min_mark, max_mark, bin_details):
        data = {"Min Marks": [min_mark], "Max Marks": [max_mark]}
        df = pd.DataFrame(data)
        with pd.ExcelWriter(output_file_path, engine="openpyxl", mode="a", if_sheet_exists="overlay") as writer:
            df.to_excel(excel_writer=writer, sheet_name="Summary", startrow=12, index=False)
            bin_details.to_excel(excel_writer=writer, sheet_name="Summary", startrow=16, index=True)




file_path = "D:/Report_checker/input/LPB01-B02-Prince BC Evaluation V2.0.xlsx"
output_file_path = "D:/Report_checker/output/LPB01-B02-Prince BC Evaluation V2.0_output.xlsx"
analyze_obj = Analyzer(file_path)
df = analyze_obj.read_file()
source_df, min_mark, max_mark = analyze_obj.analyze_data(df)
bin_details = analyze_obj.mark_falls(source_df)
analyze_obj.summary(output_file_path, min_mark, max_mark, bin_details)
