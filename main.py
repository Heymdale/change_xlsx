import glob
import json
import os.path

import pandas as pd


def get_xlsx_files():
    files = glob.glob("./in/*.xlsx")
    return files


def work_with_file(file):
    df = parse_xlsx(file)
    filename = os.path.basename(file)
    save_xlsx(filename, df)


def save_xlsx(filename: str, df_data: pd.DataFrame, sheet_name="sheet1"):
    with open(f"./out/{filename}", "wb") as file:
        df_data.to_excel(file, sheet_name=sheet_name, index=False)


def parse_xlsx(file):
    print(f"Parse the {file} file.")
    xl = pd.read_excel(file, sheet_name=None)
    sheets = list(xl.keys())
    df = xl[sheets[0]]
    df_new = pd.DataFrame(columns=[None, "Topic", "Name"])
    for i, j_string in enumerate(df["topic"]):
        j_object = json.loads(j_string)
        needed_topics = (j_object[0]["topic"][-1])
        for topic in needed_topics:
            new_row = pd.DataFrame([{"Topic": topic, "Name": df.iloc[i]["text"]}])
            df_new = pd.concat([df_new, new_row])
    return df_new


def main():
    xl_files = get_xlsx_files()
    for file in xl_files:
        work_with_file(file)


if __name__ == '__main__':
    main()
