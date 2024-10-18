"""
폴더를 순회하며 HWP 문서가 발견될시 스크립트를 진행합니다
"""


import os
import shutil
import win32com.client as win32
from natsort import natsorted

from module.log import write_log


def process_folder(input_path):
    """입력받은 폴더를 순회하여 HWP문서 스크립트를 진행합니다"""
    for root, _, files in os.walk(input_path):
        for file in natsorted(files):
            if not file.lower().endswith('.hwp'):
                continue
            new_hwp_file_path = create_folder(root, file)
            print(f"{file}변환 중...")
            hwp2xml(new_hwp_file_path, input_path)


def create_folder(root, file):
    """hwp문서 발견시 파일명으로 폴더를 생성하고 파일을 이동합니다"""
    hwp_folder_path = os.path.join(root, os.path.splitext(file)[0])
    hwp_file_path = os.path.join(root, file)

    if not os.path.exists(hwp_folder_path):
        os.makedirs(hwp_folder_path, False)

    new_hwp_file_path = os.path.join(hwp_folder_path, file)
    shutil.move(hwp_file_path, new_hwp_file_path)

    return new_hwp_file_path


def hwp2xml(hwp_file, input_path):
    """hwp에서 xml로 변환합니다"""
    try:
        hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
        hwp.SetMessageBoxMode(0x00000020)
        hwp.RegisterModule("FilePathCheckDLL", "SecurityModule")
        hwp.Open(hwp_file, arg="versionwarning:False;suspendpassword:True")
        hwp_file_only_name = os.path.splitext(hwp_file)[0]

        hwp.SaveAs(f"{hwp_file_only_name}.xml", "XML")
        hwp.Save()
        hwp.Quit()
    except Exception:  # pylint: disable=W0718
        print("변환 오류")
        write_log(os.path.join(input_path, 'log.txt'), hwp_file)
