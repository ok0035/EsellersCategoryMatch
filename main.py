# 샘플 Python 스크립트입니다.

# ⌃R을(를) 눌러 실행하거나 내 코드로 바꿉니다.
# 클래스, 파일, 도구 창, 액션 및 설정을 어디서나 검색하려면 ⇧ 두 번을(를) 누릅니다.
import openpyxl


def print_hi(name):
    # 스크립트를 디버그하려면 하단 코드 줄의 중단점을 사용합니다.
    print(f'Hi, {name}')  # 중단점을 전환하려면 ⌘F8을(를) 누릅니다.


def copyAndPasteCell():
    # A파일 열기
    wb1 = openpyxl.load_workbook('A.xlsx')
    ws1 = wb1.active  # 첫 번째 시트 선택

    # B파일 열기
    wb2 = openpyxl.load_workbook('B.xlsx')
    ws2 = wb2.active  # 첫 번째 시트 선택

    # A파일의 a셀 값 읽기
    value = ws1['a'].value

    # B파일의 b셀에 값 쓰기
    ws2['b'].value = value

    # B파일 저장하기
    wb2.save('B.xlsx')


# 스크립트를 실행하려면 여백의 녹색 버튼을 누릅니다.
if __name__ == '__main__':
    print_hi('PyCharm')

    # https://www.jetbrains.com/help/pycharm/에서 PyCharm 도움말 참조
