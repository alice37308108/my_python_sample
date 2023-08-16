# Description: フォルダ全体をZIPファイルにバックアップする。

import zipfile
from pathlib import Path


def generate_unique_backup_filename(folder, backup_folder):
    """
    バックアップファイル名を生成する関数。

    :param folder: バックアップをとるフォルダ
    :param backup_folder: バックアップ先のフォルダ
    :return:
    """
    folder_name = folder.name
    number = 1
    while True:
        zip_filename = f'{folder_name}_{number}.zip'
        backup_zip_path = backup_folder / zip_filename
        if not backup_zip_path.exists():
            return backup_zip_path
        number += 1


def backup_to_zip(folder, backup_folder):
    """
    バックアップをZIPファイルにする関数。

    :param folder: バックアップをとるフォルダ
    :param backup_folder: バックアップ先のフォルダ
    :return:
    """
    folder = Path(folder).resolve()
    backup_folder = Path(backup_folder).resolve()

    backup_zip_path = generate_unique_backup_filename(folder, backup_folder)

    # ZIPファイルを作成。
    print(f'{backup_zip_path.name} を作成中...')
    with zipfile.ZipFile(backup_zip_path, 'w') as backup_zip:
        # フォルダ内のすべてのフォルダとファイルをバックアップ。
        for item in folder.glob('**/*'):
            print(f'追加中: {item}')
            arcname = item.relative_to(folder)
            backup_zip.write(item, arcname=arcname)

    print('完了。')


if __name__ == "__main__":
    source_folder = r'folder_path'
    backup_destination = r'folder_path'
    backup_to_zip(source_folder, backup_destination)
