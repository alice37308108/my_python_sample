"""
テキスト読み上げアプリ
pip install PySimpleGUI
pip install gtts
pip install playsound==1.2.2
"""

import random
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


def create_gui():
    sg.theme('LightGray2')

    # レイアウトの定義
    layout = [
        [sg.Text("回数 ", size=(12, 1)),
         sg.Slider(range=(1, 10), orientation='h', size=(30, 20), key='count_slider', default_value=1)],
        [sg.Text("待機時間（秒）: ", size=(12, 1)),
         sg.Slider(range=(1, 10), orientation='h', size=(30, 20), key='slider', default_value=5)],
        [sg.Text("実行状況 ", size=(12, 1)), sg.ProgressBar(100, orientation='h', size=(20, 20), key='progressbar'),
         sg.Text('', size=(12, 1), key='progress_text')],
        [sg.Button('実行', key='execute_button'), sg.Button('停止'), sg.Button('終了')],
        [sg.Text('_' * 70)],  # 線を引く
        [
            sg.Text('', size=(15, 1)),
            sg.Text('', size=(30, 1), key='output'),
            sg.Column([
                [sg.Text('', size=(30, 1), key='output')],
                [sg.Image(filename='cthulhu_deep_ones.png', key='image')],
            ], justification='right')
        ],
    ]
    # ウィンドウの生成
    window = sg.Window('テキスト読み上げアプリ', layout)
    return window


def read_aloud_thread(count, sleep_duration):
    progress_bar = window['progressbar']
    progress_text = window['progress_text']
    for i in range(int(count)):
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

        # 次のダジャレまで待機
        time.sleep(sleep_duration)

        # プログレスバーを更新
        progress_value = int(((i + 1) / int(count)) * 100)
        progress_bar.update(progress_value)
        progress_text.update(f'{progress_value}% 完了')


if __name__ == "__main__":
    window = create_gui()

    while True:
        event, values = window.read()

        if event == 'execute_button':
            count = int(values['count_slider'])
            sleep_duration = int(values['slider'])
            window.start_thread(lambda: read_aloud_thread(count, sleep_duration), end_key='-THREAD_END-')
        elif event == '停止' or event == sg.WIN_CLOSED:
            break
        elif event == '終了':
            window.close()
