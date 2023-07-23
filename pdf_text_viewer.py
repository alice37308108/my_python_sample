import os
import fitz
import PySimpleGUI as sg


class GuiFrontend:
    def __init__(self):
        self.title = 'PDF テキスト表示'  # ウィンドウのタイトルを設定

    def left_col(self):
        """左側の列を返す　ファイル選択画面"""
        layout = [
            [sg.Text('フォルダを選択してください')],
            [sg.Input(key='-FOLDER-', enable_events=True), sg.FolderBrowse()],
            [sg.Listbox(values=[], size=(40, 20), key='-FILE_LIST-', enable_events=True)],
        ]

        return sg.Column(layout, vertical_alignment='t', size=(400, 800))

    def right_col(self):
        """右側の列を返す　テキスト表示画面"""
        layout = [
            [sg.Multiline(size=(80, 80), key='-PDF_CONTENT-', disabled=True)],
        ]

        return sg.Column(layout, vertical_alignment='t', size=(700, 8000))

    def layout(self):
        """ウィンドウのレイアウトを定義"""
        return [
            [self.left_col(), sg.VSeparator(), self.right_col()]  # 左列、セパレータ、右列の配置
        ]

    def window(self):
        """ウィンドウを作成して返す"""
        return sg.Window(title=self.title,
                         layout=self.layout(),
                         return_keyboard_events=True,
                         size=(1000, 750),
                         finalize=True)


class PdfReader:
    def __init__(self):
        self.folder_path = None
        self.pdf_files = []

    def get_pdf_file_list(self):
        self.pdf_files = [file for file in os.listdir(self.folder_path) if file.lower().endswith('.pdf')]

    def read_pdf_text(self, file_path):
        pdf_text = ""
        try:
            pdf_document = fitz.open(file_path)
            num_pages = pdf_document.page_count
            for page_num in range(num_pages):
                page = pdf_document.load_page(page_num)
                pdf_text += page.get_text()
        except Exception as e:
            pdf_text = f"Error: {str(e)}"
        return pdf_text

    def run(self):
        gui_frontend = GuiFrontend()
        window = gui_frontend.window()

        while True:
            event, values = window.read()

            if event == sg.WIN_CLOSED:
                break

            if event == '-FOLDER-':
                self.folder_path = values['-FOLDER-']
                self.get_pdf_file_list()
                window['-FILE_LIST-'].update(self.pdf_files)

            elif event == '-FILE_LIST-':
                selected_file = values['-FILE_LIST-'][0]
                if selected_file:
                    pdf_path = os.path.join(self.folder_path, selected_file)
                    pdf_text = self.read_pdf_text(pdf_path)
                    window['-PDF_CONTENT-'].update(pdf_text)

        window.close()


if __name__ == '__main__':
    pdf_reader = PdfReader()
    pdf_reader.run()
