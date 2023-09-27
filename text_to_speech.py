"""
テキスト読み上げアプリ
pip install PySimpleGUI
pip install gtts
pip install playsound==1.2.2
"""

import random
import threading
import time
from pathlib import Path

import PySimpleGUI as sg
from gtts import gTTS
from playsound import playsound

# ダジャレのリスト
dajare_list = [
    '布団が吹っ飛んだ',
    '猫が寝込んだ',
    'アルミ缶の上にあるミカン',
    '土管の中でドッカーン',
    'チョコをちょこっと',
]

# レイアウトの定義
layout = [
    [sg.Text("回数: "), sg.InputText(key='count', default_text='1')],
    [sg.Text("読み上げるテキスト: "), sg.Text('', size=(30, 1), key='output')],
    [sg.Button('実行'), sg.Button('終了')],
]

# ウィンドウの生成
window = sg.Window('テキスト読み上げアプリ', layout)


def read_aloud(count):
    """
    テキストを読み上げる
    :param count: 読み上げる回数
    :return:
    """
    for _ in range(int(count)):
        dajare = random.choice(dajare_list)  # ランダムにダジャレを選択
        window['output'].update(dajare)  # テキストを表示
        # テキストを音声に変換
        tts = gTTS(text=dajare, lang='ja')
        tts.save('dajare.mp3')

        # 音声を再生
        playsound('dajare.mp3')

        # ファイルを削除
        file_to_delete = Path('dajare.mp3')
        if file_to_delete.exists():
            file_to_delete.unlink()

        # 次のダジャレまで待機（テキストボックスからの秒数を使用）
        time.sleep(5)


# ボタンのイベントハンドラ
def button_handler(event):
    """
    ボタンのイベントハンドラ
    :param event:
    :return:
    """
    if event == '実行':
        count = window['count'].get()
        threading.Thread(target=read_aloud, args=(count,)).start()  # 読み上げを別スレッドで実行
    elif event == '終了':
        window.close()  # ウィンドウを閉じる


def main():
    while True:
        event, _ = window.read()  # イベント待機

        if event == sg.WIN_CLOSED:
            break
        else:
            button_handler(event)

    window.close()


if __name__ == '__main__':
    main()
