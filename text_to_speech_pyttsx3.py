"""
テキスト読み上げアプリ
pip install PySimpleGUI
pip install pyttsx3
"""

import random
import time

import PySimpleGUI as sg
import pyttsx3

# ダジャレのリスト
dajare_list = [
    '布団が吹っ飛んだ',
    '猫が寝込んだ',
    'アルミ缶の上にあるミカン',
    '土管の中でドッカーン',
    'チョコをちょこっと',
]


def create_gui():
    """
    GUIを作成する

    :return: window ウィンドウ
    """
    sg.theme('LightGray2')

    # レイアウトの定義
    layout = [
        [sg.Text('回数 ', size=(12, 1)),
         sg.Slider(range=(1, 10), orientation='h', size=(30, 20), key='count_slider', default_value=1)],
        [sg.Text('待機時間（秒）: ', size=(12, 1)),
         sg.Slider(range=(1, 10), orientation='h', size=(30, 20), key='slider', default_value=5)],
        [sg.Text('実行状況 ', size=(12, 1)), sg.ProgressBar(100, orientation='h', size=(20, 20), key='progressbar'),
         sg.Text('', size=(12, 1), key='progress_text')],
        [sg.Button('実行'), sg.Button('終了')],
        [sg.Text('_' * 70)],
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
    return sg.Window('テキスト読み上げアプリ', layout)


def read_aloud_thread(window, count, sleep_duration):
    """
    ダジャレを読み上げる

    :param window: ウィンドウ
    :param count: 読み上げる回数
    :param sleep_duration: 待機時間
    :return:
    """
    progress_bar = window['progressbar']
    progress_text = window['progress_text']

    engine = pyttsx3.init()
    engine.setProperty('rate', 130)

    for i in range(int(count)):
        dajare = random.choice(dajare_list)  # ランダムにダジャレを選択
        window['output'].update(dajare)  # テキストを表示
        engine.say(dajare)  # テキストを音声に変換
        engine.runAndWait()  # 音声を再生

        # 次のダジャレまで待機
        time.sleep(sleep_duration)

        # プログレスバーを更新
        progress_value = int(((i + 1) / int(count)) * 100)
        progress_bar.update(progress_value)
        progress_text.update(f'{progress_value}% 完了')


def main():
    window = create_gui()

    while True:
        event, values = window.read()

        if event == '実行':
            count = int(values['count_slider'])
            sleep_duration = int(values['slider'])
            window.start_thread(lambda: read_aloud_thread(window, count, sleep_duration), end_key='-THREAD_END-')
        elif event in ('終了', sg.WIN_CLOSED):
            break

    window.close()


if __name__ == '__main__':
    main()
