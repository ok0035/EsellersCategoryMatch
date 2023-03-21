<<<<<<< HEAD
# This is a sample Python script.
# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.

import xlrd


def main():
    # Use a breakpoint in the code line below to debug your script.
    # print(f'Hi, {name}')  # Press Ctrl+F8 to toggle the breakpoint.breakpoint

    category_excel_file = load_excel("원본상품등록양식Ver.1.0.4.1.xls")
    market_category_excel = load_excel("마켓카테고리매칭정보.xls")
    download_set = load_excel("D:\Mine\스마트스토어\이셀러스\ZSM_20230316_esellers.xls")

    # 이셀러스 카테고리 시트
    esellers_category_sheet = category_excel_file.sheet_by_name("이셀러스표준카테고리")

    # 마켓 카테고리 시트 (원하는 마켓으로 시트 선택)
    market_category_sheet = market_category_excel.sheet_by_name("쿠팡")

    # 다운로드 세트 확장정보
    # 이셀러스 카테고리 번호와 매칭되는 마켓 카테고리 번호를 여기에 저장
    download_set_category = download_set.sheet_by_name("확장정보")

    # Print the result
    for row in range(esellers_category_sheet.nrows):
        for col in range(esellers_category_sheet.ncols):
            print(esellers_category_sheet.cell(row, col).value, end='\t')
        print()

    # A Excel 파일에서 B Excel파일로 값을 복사 붙여넣기 하는 법
    # value = sheetA["A1"].value
    # sheetB["C3"].value = value


def load_excel(file):
    return xlrd.open_workbook(filename=file)


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    main()

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
=======
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
>>>>>>> 8f4c16b2d5de0db2444944906ba27ac0ecb596f1
