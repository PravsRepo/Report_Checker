import pdfplumber
import pandas as pd
import os
import nltk

# nltk.download('punkt')
# nltk.download('averaged_perceptron_tagger_eng')
# nltk.download('maxent_ne_chunker_tab')
# nltk.download('words')

from nltk import ne_chunk, pos_tag, word_tokenize
from nltk.tree import Tree
from pptx import Presentation
from docx import Document


class ReportChecker:

    def __init__(self, file_name, input_path):
        self.file_name = file_name
        self.input_path = input_path

    # read excel sheet to get students name and their team number
    def read_file(self):
        df = pd.read_excel(
            self.file_name, sheet_name="Sheet1")
        # print(df.head())
        return df

    def get_report_extn(self):
        file_extn = os.listdir(self.input_path)
        for file in file_extn:
            # print(file)
            raw_name, extension = os.path.splitext(file)
            team_num = raw_name.replace("Team ", "")
            print(f"{team_num},{extension}")
            if extension == ".pptx":
                text_runs = self.read_ppt(f"{input_path}/{file}",team_num)
            elif extension == ".pdf":
                self.read_pdf(f"{input_path}/{file}", team_num)
            elif extension == "docx":
                self.read_word(f"{input_path}/{file}", team_num)
            else:
                print(f"file name: {file}", team_num)
        # return text_runs

    def read_ppt(self, source_file, team_num):
        prs = Presentation(source_file)
        # print(source_file)
        text_runs = []
        # print(prs.slides[0])
        first_slide = prs.slides[0]
        for shape in first_slide.shapes:
            if shape.has_table:
                table = shape.table
                for cell in table.iter_cells():
                    text_runs.append(cell.text)
            else:
                None
        return text_runs

    def read_pdf(self, source_file):
        with pdfplumber.open(source_file) as pdf:
            first_page = pdf.pages[0]
            print(first_page.extract_text())

    def read_word(self):
        pass

    # check student names in the report ouput and return the output
    def check_names(self, df, text_runs):
        stop_words = {"Student Name", "Year of Study", "Department", "IT", "S. No.", "Roll #", "AIDS", "MECHANICAL", "Faculty Name"}
        names = []
        for run in text_runs:
            clean_run = run.strip()
            if clean_run in stop_words or clean_run.isnumeric() or clean_run=='':
                None
            else:
                names.append(clean_run.strip())
        # print(names)
        name_df = df[df["Teams"]==1]
        pattern = "|".join(names)
        # cross check the names with the dataframe
        filtered_df = name_df["Name"].str.contains(pattern, case=False, na=False)
        filtered_df.name = "Result"
        # print(filtered_df)
        return filtered_df

    def write_excel(self, result_df):
        file_path = "D:/Report_checker/output/LPB01-B02-Prince BC Evaluation V0.6_output.xlsx"

        with pd.ExcelWriter(file_path, engine="openpyxl", mode="a", if_sheet_exists="overlay") as writer:
            result_df.to_excel(excel_writer=writer, sheet_name="Sheet1", startcol=2, index=False)






file_name = "D:/Report_checker/input/LPB01-B02-Prince BC Evaluation V0.6.xlsx"
input_path = "D:/Report_checker/input/Testing_folder"
checker_obj = ReportChecker(file_name, input_path)
df = checker_obj.read_file()
text_runs = checker_obj.get_report_extn()
# result_df = checker_obj.check_names(df, text_runs)
# checker_obj.write_excel(result_df)
