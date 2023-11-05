import os

import PySimpleGUI as sg
import fitz
import win32com.client as win32


class GuiFrontend:
    def __init__(self):
        self.title = 'PDF Rename'  # ウィンドウのタイトルを設定

    @staticmethod
    def left_col():
        """左側の列を返す"""

        # 受け入れるファイルタイプを設定
        accepted_file_types = (('PDF Files', '*.pdf'),)

        layout = [
            [sg.Text('PDF'), sg.InputText(key='DOC_NAME', enable_events=True, disabled=True),  # PDFファイルパスの入力フィールド
             sg.FileBrowse(file_types=accepted_file_types, button_text='選択'),  # ファイル選択ダイアログを表示するボタン
             sg.Button('前へ'),  # 前のページに移動するボタン
             sg.Button('次へ')],  # 次のページに移動するボタン
            [sg.Image(data=None, key='IMAGE')],  # 画像を表示するためのイメージウィジェット
        ]

        return sg.Column(layout=layout, vertical_alignment='t', size=(700, 800))

    @staticmethod
    def right_col():
        """右側の列を返す"""
        layout = [
            [sg.Text('日　付'), sg.Input(key='date_input')],  # 日付の入力フィールド
            [sg.Text('取引先'), sg.Input(key='partner_input')],  # 取引先の入力フィールド
            [sg.Text('金　額'), sg.Input(key='amount_input')],  # 金額の入力フィールド
            [sg.Text('区　分'), sg.Input(key='section_input')],  # 区分の入力フィールド
            [sg.Text('不採用'), sg.Checkbox('', key='not_adopted_input')],  # 不採用を自動入力するかどうかのチェックボックス
            [sg.Button('リネーム実行', key='rename_button'),  # リネーム実行ボタン
             sg.Button('メール送信', key='send_email_button')],  # メール送信ボタン
        ]

        return sg.Column(layout=layout, vertical_alignment='t', size=(400, 800))

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


class GuiBackend:
    def __init__(self):
        self.doc = None
        self.doc_list_tab = []

    def set_doc(self, doc_name):
        """PDFドキュメントを設定する"""
        self.doc = fitz.open(doc_name)
        file_name = os.path.basename(doc_name)
        return file_name

    def get_page_count(self):
        """ページ数を返す"""
        return len(self.doc)

    def get_page(self, page_num=0):
        """
        指定されたページ番号に対応するPDFのページを返す
        :param page_num: ページ番号 (デフォルト: 0)
        :return: ページの画像データ  (バイト列)
        """

        # もし表示リストが存在しない場合、またはリストの長さがページ番号+1よりも短い場合
        # またはリストの該当する位置がNoneである場合、表示リストを取得してリストに格納する
        if len(self.doc_list_tab) < page_num + 1 or not self.doc_list_tab[page_num]:
            self.doc_list_tab.extend([None] * (page_num + 1 - len(self.doc_list_tab)))
            self.doc_list_tab[page_num] = self.doc[page_num].get_displaylist()

        # 指定されたページ番号に対応する表示リストを取得する
        doc_list = self.doc_list_tab[page_num]

        # 表示リストからピクセルマップを取得する
        pix = doc_list.get_pixmap(alpha=False)

        # もしファイルのサイズの幅が680以上だったら、横幅が680以下になるように縮小する
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
        next_page_event = ('次へ', 'MouseWheel:Down')
        prev_page_event = ('前へ', 'MouseWheel:Up')

        while True:
            event, values = self.window.read(timeout=100)
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

            # doc_nameが指定されていないときにイベントが発生したら、何もしない
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
                data = self.backend.get_page(self.page)
                self.window['IMAGE'].Update(data=data)

            # 実行ボタンが押されたら、入力された値を取得する
            if event == 'rename_button':
                self.process_rename(values)

            # メール送信ボタンが押されたらリネームしてメールを送信する
            if event == 'send_email_button':
                new_filename, values_dict = self.process_rename(values)
                # new_filenameがNoneでなければメールを送信する
                if new_filename:
                    renamed_file_path = os.path.join(os.path.dirname(self.doc_name), new_filename)
                    self.send_email(renamed_file_path, values_dict)

    def process_rename(self, values):
        """リネーム処理を行う"""
        date = values['date_input']
        partner = values['partner_input']
        amount = values['amount_input']
        section = values['section_input']
        not_adopted = values['not_adopted_input']
        values_dict = {'date': date, 'partner': partner, 'amount': amount, 'section': section,}

        new_filename = self.rename_pdf(date, partner, amount, section, not_adopted)
        return new_filename, values_dict

    def rename_pdf(self, date, partner, amount, section, not_adopted):
        """PDFをリネームする"""
        if not date.isdigit() or len(date) != 8:
            sg.popup('日付を8桁の数字で入力してください')
            return

        if not amount.isdigit():
            sg.popup('金額を数字で入力してください')
            return

        if not section:
            sg.popup('区分を入力してください')
            return

        adopted_text = '_不' if not_adopted else ''
        if date and partner and amount and section:
            if self.doc_name:
                self.backend.doc.close()  # ファイルを閉じる
                new_filename = f'{date}_{partner}_{amount}_{section}{adopted_text}.pdf'
                new_filepath = os.path.join(os.path.dirname(self.doc_name), new_filename)

                os.rename(self.doc_name, new_filepath)

                sg.popup(f'ファイル名を変更しました！ {new_filename}', title='完了')
                self.window['IMAGE'].update(data=None)
                self.window['DOC_NAME'].update(value='')
                return new_filename
        else:
            sg.popup('すべて入力してください')

    def send_email(self, file_path, values_list):
        outlook = win32.Dispatch('Outlook.Application')
        mail_item = outlook.CreateItem(0)  # メールアイテムを作成

        mail_text = f'{values_list["date"]} {values_list["partner"]} {values_list["amount"]} {values_list["section"]}'
        mail_item.Subject = 'メールの件名'  # 件名を設定
        mail_item.Body = mail_text  # 本文を設定

        # ファイルを添付
        attachment = os.path.abspath(file_path)
        mail_item.Attachments.Add(attachment)

        # メールを表示（送信前確認）
        mail_item.Display()


def main():
    gui = PdfReader()
    gui.event_loop()


if __name__ == '__main__':
    main()
