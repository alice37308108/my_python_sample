import os
import shutil


def copy_files_with_pattern(pattern_str):
    """
    指定されたパターンを含むファイルをtestフォルダから探してnewフォルダにコピーする

    Args:
        pattern_str (str): 検索するファイル名のパターン
    """
    os.makedirs("new", exist_ok=True)

    for root, dirs, files in os.walk("test"):
        for file in files:
            if file.endswith(".csv") and pattern_str in file:
                src = os.path.join(root, file)
                dst = os.path.join("new", file)
                shutil.copy2(src, dst)


if __name__ == "__main__":
    pattern = input("検索するパターンを入力してください: ")
    copy_files_with_pattern(pattern)