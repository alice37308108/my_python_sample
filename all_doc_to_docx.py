# 選択したフォルダのサブフォルダ内も含めて.docファイルを.docxファイルに変換する
# .docxに変換したら.docファイルはゴミ箱に削除する
# pip install pywin32
# pip install send2trash

import os
from pathlib import Path

import PySimpleGUI as sg
import send2trash
import win32com.client as win32


def convert_doc_to_docx(doc_file_path):
    """
    .docファイルを.docxファイルに変換する

    :param doc_file_path:
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
    .docファイルはゴミ箱に削除する

    :param folder_path:
    :return:
    """
    for doc_file_path in folder_path.glob('**/*.doc'):
        convert_doc_to_docx(str(doc_file_path))
        send2trash.send2trash(str(doc_file_path))

    sg.popup('完了しました！')


def main():
    """
    フォルダ内の全ての.docファイルを.docxファイルに変換する

    :return:
    """
    layout = [
        [sg.Text('変換するフォルダを選択してください')],
        [sg.InputText(key='folder'), sg.FolderBrowse()],
        [sg.Button('変換'), sg.Button('キャンセル')]
    ]

    window = sg.Window('Doc to Docx Converter', layout)

    while True:
        event, values = window.read()

        if event == sg.WINDOW_CLOSED or event == 'キャンセル':
            break
        elif event == '変換':
            folder_path = Path(values['folder'])
            batch_convert_docs_in_folder(folder_path)
            break

    window.close()


if __name__ == '__main__':
    main()
