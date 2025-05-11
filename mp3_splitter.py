# pip install mutagen

import os
import tkinter as tk
from tkinter import filedialog, ttk, messagebox

from mutagen.mp3 import MP3


class MP3SplitterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("MP3分割ツール")
        self.root.geometry("600x500")
        self.root.resizable(False, False)

        # 入力ファイル関連
        self.input_frame = ttk.LabelFrame(root, text="入力ファイル")
        self.input_frame.pack(fill="x", padx=10, pady=10)

        self.input_path = tk.StringVar()
        ttk.Entry(self.input_frame, textvariable=self.input_path, width=50).pack(side="left", padx=5, pady=5)
        ttk.Button(self.input_frame, text="参照...", command=self.browse_input_file).pack(side="left", padx=5, pady=5)

        # 出力フォルダ関連
        self.output_frame = ttk.LabelFrame(root, text="出力フォルダ")
        self.output_frame.pack(fill="x", padx=10, pady=10)

        self.output_path = tk.StringVar()
        ttk.Entry(self.output_frame, textvariable=self.output_path, width=50).pack(side="left", padx=5, pady=5)
        ttk.Button(self.output_frame, text="参照...", command=self.browse_output_folder).pack(side="left", padx=5,
                                                                                              pady=5)

        # 分割時間設定
        self.time_frame = ttk.LabelFrame(root, text="分割設定")
        self.time_frame.pack(fill="x", padx=10, pady=10)

        # 分割方式選択（固定時間または指定時間）
        self.split_mode = tk.StringVar(value="fixed")
        ttk.Radiobutton(self.time_frame, text="固定時間で分割", variable=self.split_mode,
                        value="fixed", command=self.toggle_split_mode).pack(anchor="w", padx=5, pady=2)

        # 固定時間分割の設定
        self.fixed_frame = ttk.Frame(self.time_frame)
        self.fixed_frame.pack(fill="x", padx=10, pady=5)

        # 時間入力（時:分:秒）
        ttk.Label(self.fixed_frame, text="分割時間:").pack(side="left", padx=5)

        self.hours = tk.IntVar(value=0)
        ttk.Spinbox(self.fixed_frame, from_=0, to=24, textvariable=self.hours, width=2).pack(side="left")
        ttk.Label(self.fixed_frame, text="時間").pack(side="left")

        self.minutes = tk.IntVar(value=5)
        ttk.Spinbox(self.fixed_frame, from_=0, to=60, textvariable=self.minutes, width=2).pack(side="left", padx=(5, 0))
        ttk.Label(self.fixed_frame, text="分").pack(side="left")

        self.seconds = tk.IntVar(value=0)
        ttk.Spinbox(self.fixed_frame, from_=0, to=59, textvariable=self.seconds, width=2).pack(side="left", padx=(5, 0))
        ttk.Label(self.fixed_frame, text="秒").pack(side="left")

        # カスタム時間分割の設定
        ttk.Radiobutton(self.time_frame, text="カスタム時間で分割", variable=self.split_mode,
                        value="custom", command=self.toggle_split_mode).pack(anchor="w", padx=5, pady=2)

        self.custom_frame = ttk.Frame(self.time_frame)
        self.custom_frame.pack(fill="x", padx=10, pady=5)

        ttk.Label(self.custom_frame, text="分割時間ポイント (h:mm:ss または mm:ss, カンマ区切り):").pack(anchor="w",
                                                                                                         padx=5, pady=2)
        self.custom_times = tk.StringVar(value="00:01:30, 00:03:45, 01:07:20")
        custom_entry = ttk.Entry(self.custom_frame, textvariable=self.custom_times, width=50)
        custom_entry.pack(fill="x", padx=5, pady=2)
        ttk.Label(self.custom_frame, text="例: 00:01:30, 00:03:45, 01:07:20 または 1:30, 3:45, 67:20").pack(anchor="w",
                                                                                                            padx=5,
                                                                                                            pady=0)

        # 初期状態でカスタムフレームは無効化
        for child in self.custom_frame.winfo_children():
            child.configure(state="disabled")

        # 実行ボタン
        self.button_frame = ttk.Frame(root)
        self.button_frame.pack(fill="x", padx=10, pady=20)
        ttk.Button(self.button_frame, text="分割開始", command=self.split_mp3, width=20).pack(anchor="center")

        # 結果表示領域
        self.result_frame = ttk.LabelFrame(root, text="結果")
        self.result_frame.pack(fill="both", expand=True, padx=10, pady=10)

        self.result_text = tk.Text(self.result_frame, wrap="word", height=8)
        self.result_text.pack(fill="both", expand=True, padx=5, pady=5)

        # ファイル情報
        self.file_info = None

    def toggle_split_mode(self):
        if self.split_mode.get() == "fixed":
            # 固定時間モードを有効化、カスタム時間モードを無効化
            for child in self.fixed_frame.winfo_children():
                child.configure(state="normal")
            for child in self.custom_frame.winfo_children():
                child.configure(state="disabled")
        else:
            # カスタム時間モードを有効化、固定時間モードを無効化
            for child in self.fixed_frame.winfo_children():
                child.configure(state="disabled")
            for child in self.custom_frame.winfo_children():
                child.configure(state="normal")

    def browse_input_file(self):
        """入力ファイルを選択するダイアログを表示"""
        file_path = filedialog.askopenfilename(
            filetypes=[("MP3ファイル", "*.mp3"), ("すべてのファイル", "*.*")]
        )
        if file_path:
            self.input_path.set(file_path)

            # 出力フォルダを入力ファイルと同じフォルダに設定
            output_dir = os.path.dirname(file_path)
            self.output_path.set(output_dir)

            # MP3ファイル情報を取得
            try:
                mp3_file = MP3(file_path)
                duration = mp3_file.info.length  # 秒単位の長さ

                hours = int(duration // 3600)
                minutes = int((duration % 3600) // 60)
                seconds = int(duration % 60)

                self.file_info = {
                    "duration": duration,
                    "hours": hours,
                    "minutes": minutes,
                    "seconds": seconds,
                    "bitrate": mp3_file.info.bitrate,
                    "sample_rate": mp3_file.info.sample_rate,
                    "channels": 2 if mp3_file.info.mode != 3 else 1  # モード3ならモノラル
                }

                time_str = f"{hours}時間 " if hours > 0 else ""
                time_str += f"{minutes}分 {seconds}秒"

                info_text = f"ファイル情報:\n長さ: {time_str}\nビットレート: {mp3_file.info.bitrate // 1000} kbps\nサンプルレート: {mp3_file.info.sample_rate} Hz"
                self.result_text.delete(1.0, tk.END)
                self.result_text.insert(tk.END, info_text)

            except Exception as e:
                self.result_text.delete(1.0, tk.END)
                self.result_text.insert(tk.END, f"ファイル情報の取得に失敗しました: {e}")

    def browse_output_folder(self):
        """出力フォルダを選択するダイアログを表示"""
        folder_path = filedialog.askdirectory()
        if folder_path:
            self.output_path.set(folder_path)

    def parse_time_to_seconds(self, time_str):
        """時間文字列を秒数に変換（h:mm:ss または mm:ss 形式対応）"""
        time_str = time_str.strip()

        # h:mm:ss 形式の場合
        if time_str.count(':') == 2:
            h, m, s = map(int, time_str.split(':'))
            return h * 3600 + m * 60 + s

        # mm:ss 形式の場合
        elif time_str.count(':') == 1:
            m, s = map(int, time_str.split(':'))
            return m * 60 + s

        # 秒数のみの場合
        else:
            return int(time_str)

    def format_time(self, seconds):
        """秒数を h:mm:ss 形式に変換"""
        h = int(seconds // 3600)
        m = int((seconds % 3600) // 60)
        s = int(seconds % 60)

        if h > 0:
            return f"{h:02d}:{m:02d}:{s:02d}"
        else:
            return f"{m:02d}:{s:02d}"

    def sanitize_filename(self, filename):
        """ファイル名から不正な文字を削除して安全なファイル名にする"""
        # Windowsで使用できない文字を除去
        invalid_chars = '<>:"/\\|?*'

        # ファイル名から無効な文字を削除
        for char in invalid_chars:
            filename = filename.replace(char, '_')

        # 長すぎるファイル名を短くする（240文字以内）
        if len(filename) > 240:
            name, ext = os.path.splitext(filename)
            filename = name[:240 - len(ext)] + ext

        return filename

    def split_mp3(self):
        """MP3分割を実行"""
        input_file = self.input_path.get()
        output_folder = self.output_path.get()

        # 入力チェック
        if not input_file:
            messagebox.showerror("エラー", "入力ファイルを選択してください。")
            return

        if not output_folder:
            messagebox.showerror("エラー", "出力フォルダを選択してください。")
            return

        if not os.path.exists(output_folder):
            try:
                os.makedirs(output_folder)
            except Exception as e:
                messagebox.showerror("エラー", f"出力フォルダの作成に失敗しました: {e}")
                return

        # 分割ポイントの計算
        split_points = []

        if self.split_mode.get() == "fixed":
            # 固定時間モードの場合
            hours = self.hours.get()
            minutes = self.minutes.get()
            seconds = self.seconds.get()

            if hours == 0 and minutes == 0 and seconds == 0:
                messagebox.showerror("エラー", "分割時間を設定してください。")
                return

            interval_seconds = hours * 3600 + minutes * 60 + seconds
            total_duration = self.file_info["duration"]

            current_time = interval_seconds
            while current_time < total_duration:
                split_points.append(current_time)
                current_time += interval_seconds

        else:
            # カスタム時間モードの場合
            custom_times = self.custom_times.get().strip()
            if not custom_times:
                messagebox.showerror("エラー", "分割時間を設定してください。")
                return

            try:
                for time_str in custom_times.split(","):
                    split_points.append(self.parse_time_to_seconds(time_str))

                # 昇順に並べ替え
                split_points.sort()

            except Exception as e:
                messagebox.showerror("エラー", f"時間形式が正しくありません: {e}")
                return

        # 結果テキストをクリア
        self.result_text.delete(1.0, tk.END)
        self.result_text.insert(tk.END, "分割処理を開始します...\n")
        self.root.update()

        try:
            # MP3ファイルを読み込む
            self.result_text.insert(tk.END, "ファイルを読み込んでいます...\n")
            self.root.update()

            # ファイル名の取得
            base_filename = os.path.basename(input_file)
            filename_without_ext = os.path.splitext(base_filename)[0]

            # MP3ファイルの読み込み
            with open(input_file, 'rb') as mp3_file:
                mp3_data = mp3_file.read()

            # MP3ファイルの解析
            mp3_info = MP3(input_file)
            total_duration = mp3_info.info.length
            file_size = len(mp3_data)

            # 素朴な方法: ファイルサイズを時間で比例配分して分割
            # これは厳密ではありませんが、外部ライブラリなしで近似的に機能します

            # 分割ポイントを秒からバイト位置に変換
            byte_positions = [int(split_point / total_duration * file_size) for split_point in split_points]

            # ファイルの分割実行
            segments = []
            start_pos = 0

            for i, pos in enumerate(byte_positions):
                segments.append((start_pos, pos))
                start_pos = pos

            # 最後のセグメント
            segments.append((start_pos, file_size))

            # 分割ファイルの作成
            for i, (start, end) in enumerate(segments):
                segment_data = mp3_data[start:end]

                # 出力ファイル名
                if i == 0:
                    # 最初のセグメント
                    start_time = "00_00_00"
                    end_time = self.format_time(split_points[i]).replace(":", "_")
                    output_filename = f"{filename_without_ext}_{start_time}-{end_time}.mp3"
                elif i == len(segments) - 1:
                    # 最後のセグメント
                    start_time = self.format_time(split_points[i - 1]).replace(":", "_")
                    end_time = self.format_time(total_duration).replace(":", "_")
                    output_filename = f"{filename_without_ext}_{start_time}-{end_time}.mp3"
                else:
                    # 中間セグメント
                    start_time = self.format_time(split_points[i - 1]).replace(":", "_")
                    end_time = self.format_time(split_points[i]).replace(":", "_")
                    output_filename = f"{filename_without_ext}_{start_time}-{end_time}.mp3"

                # ファイル名をサニタイズ
                output_filename = self.sanitize_filename(output_filename)
                output_path = os.path.join(output_folder, output_filename)

                # ファイルに書き込み
                try:
                    with open(output_path, 'wb') as out_file:
                        out_file.write(segment_data)

                    self.result_text.insert(tk.END, f"分割ファイル作成: {output_filename}\n")
                    self.root.update()
                except Exception as e:
                    self.result_text.insert(tk.END, f"ファイル保存エラー ({output_filename}): {e}\n")
                    self.root.update()

            result_msg = f"\n完了しました！\n合計 {len(segments)} 個のファイルを {output_folder} に保存しました。"
            self.result_text.insert(tk.END, result_msg)
            messagebox.showinfo("完了", f"MP3の分割が完了しました！\n{len(segments)} 個のファイルを作成しました。")

        except Exception as e:
            error_msg = f"エラーが発生しました: {e}"
            self.result_text.insert(tk.END, error_msg)
            messagebox.showerror("エラー", error_msg)

        finally:
            self.result_text.insert(tk.END,
                                    "\n\n注意: このアプリはFFmpegなしでMP3を分割しているため、正確な分割ではない可能性があります。")
            self.result_text.insert(tk.END,
                                    "\n分割ポイント付近で音声が途切れる場合があります。より正確な分割にはFFmpegの使用をお勧めします。")


if __name__ == "__main__":
    root = tk.Tk()
    app = MP3SplitterApp(root)
    root.mainloop()
