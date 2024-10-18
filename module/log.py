"""
Exception 발생시 로그로 남깁니다
"""


def write_log(log_path, file_path):
    """Exception 발생시 로그로 남깁니다"""
    with open(log_path, 'a', encoding='UTF-8') as file:
        file.write(f"{file_path}\n")
