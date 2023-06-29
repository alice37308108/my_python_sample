import os

import fitz
import PySimpleGUI as sg


class GuiFrontend:
    """PDF RenameアプリのGUIフロントエンドを定義するクラス"""

    def __init__(self):
        self.title = 'PDF Rename'  # ウィンドウのタイトルを設定

    @staticmethod
    def left_col():
        """左側のカラムを返す"""

        # 受け入れるファイルタイプを設定
        accepted_file_types = (("PDF Files", "*.pdf"),)

        layout = [
            [sg.Text('PDF'), sg.InputText(key='DOC_NAME', enable_events=True),  # PDFファイルパスの入力フィールド
             sg.FileBrowse(file_types=accepted_file_types, button_text='選択'),  # ファイル選択ダイアログを表示するボタン
             sg.Button('前へ'),  # 前のページに移動するボタン
             sg.Button('次へ')],  # 次のページに移動するボタン
            [sg.Image(data=None, key='IMAGE')],  # 画像を表示するためのイメージウィジェット
        ]

        return sg.Column(layout=layout, vertical_alignment='t', size=(700, 800))

    @staticmethod
    def right_col():
        """右側のカラムを返す"""

        layout = [
            [sg.Text('日　付'), sg.Input(key='date_input')],  # 日付の入力フィールド
            [sg.Text('取引先'), sg.Input(key='partner_input')],  # 取引先の入力フィールド
            [sg.Text('金　額'), sg.Input(key='amount_input')],  # 金額の入力フィールド
            [sg.Button('実行', bind_return_key=True)],  # 実行ボタン
        ]

        return sg.Column(layout=layout, vertical_alignment='t', size=(400, 800))

    def layout(self):
        """ウィンドウのレイアウトを定義"""
        return [
            [self.left_col(), sg.VSeparator(), self.right_col()]  # 左カラム、セパレータ、右カラムの配置
        ]

    def window(self):
        """ウィンドウを作成して返す"""
        return sg.Window(title=self.title,
                         layout=self.layout(),
                         return_keyboard_events=True,
                         use_default_focus=False,
                         size=(1000, 750),
                         finalize=True)


class GuiBackend:
    def __init__(self):
        self.doc = None

    def set_doc(self, doc_name):
        self.doc = fitz.open(doc_name)
        file_name = os.path.basename(doc_name)  # ファイルパスからファイル名のみを抽出
        return file_name


    def get_page_count(self):
        return len(self.doc)

    def get_doc_list_tab(self):
        page_count = self.get_page_count()
        return [None] * page_count

    def get_doc_list(self, page_num):
        doc_list_tab = self.get_doc_list_tab()
        return doc_list_tab[page_num]

    def get_page(self, page_num=0, zoom=0):
        """PDFの指定されたページを返す"""
        doc_list = self.get_doc_list(page_num)
        doc_list_tab = self.get_doc_list_tab()
        if not doc_list:
            doc_list_tab[page_num] = self.doc[page_num].get_displaylist()
            doc_list = doc_list_tab[page_num]

        pix = doc_list.get_pixmap(alpha=False)

        # もしファイルのサイズの幅が500以上だったら、横幅が500以下になるように縮小する
        if pix.width > 680:
            zoom = 680 / pix.width
            pix = doc_list.get_pixmap(alpha=False, matrix=fitz.Matrix(zoom, zoom))

        return pix.tobytes()


class PdfReader:
    """PDFリーダーGUI"""

    def __init__(self):
        self.backend = GuiBackend()
        self.frontend = GuiFrontend()
        self.window = self.frontend.window()
        self.page = 0
        self.total_page = 0
        self.doc_name = None

    @staticmethod
    def get_next_page(page, total_count):
        """次のページ番号を返す"""
        page += 1
        # トータルページ数に到達していた場合は最初のページ
        if page >= total_count:
            return 0
        else:
            return page

    @staticmethod
    def get_prev_page(page, total_count):
        """前のページ番号を返す"""
        page -= 1
        # マイナスの値になった場合は最後のページ
        if page < 0:
            return total_count - 1
        else:
            return page

    def event_loop(self):
        """イベントループする"""
        next_page_event = ("次へ", "MouseWheel:Down")
        prev_page_event = ("前へ", "MouseWheel:Up")
        enter_event = chr(13)

        while True:
            event, values = self.window.read(timeout=100)
            zoom = 0
            # ページ更新の制御
            is_page_update = False

            if event == sg.WIN_CLOSED:
                break

            if event == 'DOC_NAME':
                self.doc_name = values['DOC_NAME']
                file_name = self.backend.set_doc(self.doc_name)  # ファイル名を取得
                self.window['DOC_NAME'].update(value=file_name)  # ファイル名を表示

                self.total_page = self.backend.get_page_count()
                self.page = 0
                is_page_update = True

            # doc_nameが指定されておらず、何らかのイベントが発生
            if event and not self.doc_name:
                continue

            # 次ページ
            if event in next_page_event:
                self.page = self.get_next_page(self.page, self.total_page)
                is_page_update = True

            # 前ページ
            if event in prev_page_event:
                self.page = self.get_prev_page(self.page, self.total_page)
                is_page_update = True

            # 表示ページの更新
            if is_page_update:
                data = self.backend.get_page(self.page, zoom)
                self.window['IMAGE'].Update(data=data)

            if event == '実行':


                date = values['date_input']
                partner = values['partner_input']
                amount = values['amount_input']

                # もし日付が8桁の数字でなかったら、メッセージを表示して再度入力する
                if not date.isdigit() or len(date) != 8:
                    sg.popup('日付を8桁の数字で入力してください')
                    continue

                if date and partner and amount:
                    if self.doc_name:
                        # base_name = os.path.basename(self.doc_name)
                        # base_name_without_ext = os.path.splitext(base_name)[0]
                        self.backend.doc.close()  # ファイルを閉じる
                        new_filename = f"{date}_{partner}_{amount}.pdf"
                        new_filepath = os.path.join(os.path.dirname(self.doc_name), new_filename)

                        os.rename(self.doc_name, new_filepath)

                        sg.popup(f'ファイル名を変更しました！ {new_filename}', title='完了')

                    else:
                        sg.popup("Please select a PDF file")
                else:
                    sg.popup("Please enter all fields")


def job():
    gui = PdfReader()
    gui.event_loop()


if __name__ == '__main__':
    job()
