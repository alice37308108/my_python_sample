"""
PDFからテーブルを抽出し、データフレームに変換します。
pip install PyMuPDF
pip install pandas
pip install openpyxl
"""

import fitz
import pandas as pd


def extract_and_transform_tables(pdf_path, start_page):
    """
    PDFからテーブルを抽出し、データフレームに変換します。

    :param pdf_path: PDFファイルのパス
    :param start_page: 開始ページ番号
    :return: テーブルを変換したデータフレームのリスト
    """
    doc = fitz.open(pdf_path)
    extracted_dfs = []

    for page_num in range(start_page, doc.page_count + 1):
        page = doc[page_num - 1]  # ページ番号は0から始まるため、1を引く
        tables = page.find_tables()  # テーブルを検索

        if tables.tables:  # テーブルが見つかった場合
            table_data = tables[0].extract()
            columns = table_data[0]
            data_rows = table_data[1:]

            df = pd.DataFrame(data_rows, columns=columns)
            extracted_dfs.append(df)

    return extracted_dfs


def split_and_clean_columns(df):
    """
    データフレームの6列目を分割してクリーンアップします。

    :param df: データフレーム
    :return: 分割されたデータフレーム
    """
    new_columns = df.iloc[:, 5].str.split('\n', expand=True)
    new_columns.columns = ['振り仮名', '名前']

    new_columns['振り仮名'] = new_columns['振り仮名'].str.replace(' ', '')
    new_columns['名前'] = new_columns['名前'].str.replace(' ', '')

    return new_columns


if __name__ == "__main__":
    pdf_path = 'test.pdf'
    start_page = 4  # 開始ページ番号
    excel_path = 'combined_tables.xlsx'

    extracted_dataframes = extract_and_transform_tables(pdf_path, start_page)

    combined_df = pd.concat(extracted_dataframes, ignore_index=True)

    new_columns = split_and_clean_columns(combined_df)
    combined_df = pd.concat([combined_df.drop(columns=combined_df.columns[5]), new_columns], axis=1)

    combined_df.to_excel(excel_path, index=False)
    print(f"変換されたクリーンなデータフレームは {excel_path} に保存されました")
