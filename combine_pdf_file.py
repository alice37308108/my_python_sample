"""
同一ディレクトリにあるPDFファイルを1つのファイルに結合する
pip install PyPDF2
"""

import PyPDF2
from pathlib import Path
from datetime import datetime


def merge_pdfs():
    """PDFファイルを結合する関数"""

    merger = PyPDF2.PdfMerger()
    # 現在の作業ディレクトリを入力フォルダとする
    input_folder = Path().cwd()
    # 入力フォルダのパスをPathオブジェクトとして取得
    base_path = Path(input_folder)

    # 入力フォルダ内のPDFファイルをリストとして取得し、ファイル名でソート
    pdf_list = list(base_path.glob('*.pdf'))
    sorted_pdf_list = sorted(pdf_list)

    # PDFファイルを順に結合
    for sorted_pdf in sorted_pdf_list:
        merger.append(str(sorted_pdf))

    # 出力ファイル名を指定し、結合したPDFファイルを保存
    now = datetime.now()
    dt = now.strftime('%Y%m%d%H%M%S')
    output_path = base_path / (dt + '.pdf')  # 出力ファイル名を指定してパスを生成

    merger.write(str(output_path))
    merger.close()


if __name__ == "__main__":
    merge_pdfs()