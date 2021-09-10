import pandas as pd
from docx import Document
import os
import pathlib
import time


def read_docx_table(document, table_num=1, nheader=1):
    table = document.tables[table_num - 1]
    data = [[cell.text for cell in row.cells] for row in table.rows]
    df = pd.DataFrame(data)

    if nheader == 1:
        df = df.rename(columns=df.iloc[0]).drop(df.index[0]).reset_index(drop=True)
    elif nheader == 2:
        outside_col, inside_col = df.iloc[0], df.iloc[1]
        hier_index = pd.MultiIndex.from_tuples(list(zip(outside_col, inside_col)))
        df = pd.DataFrame(data, columns=hier_index).drop(df.index[[0, 1]]).reset_index(drop=True)
    elif nheader > 2:
        print("More than two headers not supported")
        df = pd.DataFrame()
    return df


def main():
    docx_list = []
    docx_list_cleaned = []

    data_to_be_processed = "data_to_be_processed"
    cleaned_data = "cleaned_data"

    if not os.path.exists(data_to_be_processed):
        os.makedirs(data_to_be_processed)

    if not os.path.exists(cleaned_data):
        os.makedirs(cleaned_data)

    os.chdir(data_to_be_processed)

    docx = pathlib.Path().glob("*.docx")
    docx_cleaned = pathlib.Path("cleaned_data").glob("*.xlsx")

    for file in docx:
        docx_list.append(file)

    os.chdir("../")

    for i in docx_list:
        print(i)
        file = f"{i}"
        cleaned_file = f"{file} Cleaned.xlsx"

        if os.path.exists(cleaned_file):
            os.remove(cleaned_file)

        os.chdir(data_to_be_processed)
        document = Document(
            f"{file}"
        )
        df = read_docx_table(document, table_num=2, nheader=2)
        # print(df)

        os.chdir("../")
        os.chdir(cleaned_data)
        df.to_excel(f"{cleaned_file}", header=True, index=True)
        os.chdir("../")

    for file in docx_cleaned:
        docx_list_cleaned.append(file)
        print(file.name)


if __name__ == '__main__':
    main()
