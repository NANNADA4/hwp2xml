"""
main함수.
"""


import os


from module.process_folder import process_folder


def main():
    """main 함수"""
    print("-"*24)
    print("\n>>>>>>HWP XML 변환<<<<<<\n")
    print("-"*24)
    folder_path = input("폴더 경로를 입력하세요 (종료는 0을 입력) : ")

    if folder_path == '0':
        return 0

    if not os.path.isdir(folder_path):
        print("입력 폴더의 경로를 다시 한번 확인하세요")
        return main()

    folder_path = os.path.join("\\\\?\\", folder_path)
    process_folder(folder_path)

    print("\n~~~모든 문서 변환이 완료되었습니다~~~\n")

    return main()


if __name__ == "__main__":
    main()
