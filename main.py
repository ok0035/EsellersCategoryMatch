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
