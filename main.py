import pandas as pd
from docx import Document


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
    document = Document(
        r"Z:\DMZ_Proj_Files\Chemical Morning First Semester 2017\chemical 17 (1) M (01) Applied Chemistry (Th).docx"
    )
    df = read_docx_table(document, table_num=2, nheader=2)
    print(df)
    df.to_excel("test.xlsx")


if __name__ == '__main__':
    main()
