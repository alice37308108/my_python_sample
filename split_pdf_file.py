"""
同一ディレクトリにあるPDFファイルをページごとに分割する
pip install PyPDF2
"""

from pathlib import Path
import PyPDF2


def split_pdf(input_path):
    # 入力ファイルのPathオブジェクトを作成
    input_path = Path(input_path)

    # 出力ディレクトリを作成（元のファイル名と同じ名前のディレクトリを作成）
    output_dir = input_path.with_name(input_path.stem)
    output_dir.mkdir(exist_ok=True)

    # PDFファイルを開く
    with input_path.open('rb') as file:
        # PyPDF2のPdfReaderオブジェクトを作成
        pdf_reader = PyPDF2.PdfReader(file)

        # 各ページを個別のPDFファイルとして保存
        for page_number in range(len(pdf_reader.pages)):
            pdf_writer = PyPDF2.PdfWriter()
            pdf_writer.add_page(pdf_reader.pages[page_number])

            output_path = output_dir / f'{page_number + 1}_{input_path.stem}.pdf'

            with output_path.open('wb') as output_file:
                pdf_writer.write(output_file)


if __name__ == '__main__':
    # カレントディレクトリにある全てのPDFファイルを取得
    pdf_files = [file for file in Path.cwd().iterdir() if file.suffix == '.pdf']

    # 各PDFファイルを分割する
    for pdf_file in pdf_files:
        split_pdf(pdf_file)
