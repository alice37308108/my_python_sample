# docファイルをdocxファイルに変換する
# pip install pywin32

import os
import glob
import win32com.client as win32


def convert_doc_to_docx(doc_file_path):
    """
    .docファイルを.docxファイルに変換する
    :param doc_file_path:ファイルパス
    :return:
    """
    word = win32.Dispatch('Word.Application')
    doc = word.Documents.Open(doc_file_path)

    # .docxファイルのパスを生成
    docx_file_path = os.path.splitext(doc_file_path)[0] + '.docx'

    # .docファイルを.docx形式で保存します (FileFormat=16: .docx format)
    doc.SaveAs2(docx_file_path, FileFormat=16)  # 16: .docx format

    doc.Close()
    word.Quit()


def batch_convert_docs_in_folder(folder_path):
    """
    フォルダ内の全ての.docファイルを.docxファイルに変換する
    :param folder_path: フォルダパス
    :return:
    """
    doc_files = glob.glob(os.path.join(folder_path, '*.doc'))
    for doc_file_path in doc_files:
        convert_doc_to_docx(doc_file_path)
    print('完了しました！')


if __name__ == '__main__':
    # フォルダパスを指定して、フォルダ内の全ての.docファイルを.docxに変換する
    folder_path = r'folder_path'
    batch_convert_docs_in_folder(folder_path)
