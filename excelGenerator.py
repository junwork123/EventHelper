from os import listdir
from openpyxl import load_workbook, Workbook
from Event import *
from openpyxl.styles import PatternFill, Color # 셀을 컬러로 칠할 때 사용할 예정

OUT_KEYWORDS = []
IN_KEYWORDS = []

# 이벤트 옵션을 전역변수에 세팅
def setOptions(Event):
    OUT_KEYWORDS.extend(Event.getOpt_Cherry())
    IN_KEYWORDS.extend(Event.getOpt_URL())
    IN_KEYWORDS.extend(Event.getOpt_Keywords())

# 이벤트 조건과 맞을때 셀에 추가
def check_Keyword(value):

    for word in IN_KEYWORDS:
        if (not (word in value)):
            break

        if (word is IN_KEYWORDS[-1]):
            return True

    return False


# 체리피커를 걸러내는 함수
def check_Cherry(value):

    # 체리피커 키워드가 있을때 거른다
    for word in OUT_KEYWORDS:
        if (word in value):
            return False

    return True

def generateExcel(Event):

    # 키워드 세팅
    setOptions(Event)

    files = listdir(".")
    src_xlsx = load_workbook("src.xlsx")
    src_sheet = src_xlsx.active

    src_sheet.delete_cols(1)  # Post Reaction
    src_sheet.delete_cols(6, 3)  # Likes Count , Tagged User Count on This Comment, Tagged User Count in Entire Post
    col_date = src_sheet['E']  # DATE COL

    result_xlsx = Workbook()
    result_sheet = result_xlsx.active

    count = 0
    # ID     CommentLink	Name	Date	Text
    for row in src_sheet.iter_rows():

        cell_data = []
        # 첫 줄에 구분자 삽입
        if count == 0:
            cell_data.append("No")

            # 첫 줄(카테고리) 삽입
            for cell in row:
                cell_data.append(cell.value)
            result_sheet.append(cell_data)
            count = count + 1
            continue

        # 첫째칸이 비어있거나, Text가 비어있다면 패스
        if (row[0].value is None or row[4].value is None):
            continue

        # 이벤트 조건

        # 이벤트 조건과 맞을때 셀에 추가

        if (check_Keyword(row[4].value) and check_Cherry(row[4].value)):
            # 두번째 줄(A2) 부터 데이터 복사
            cell_data.append(count)
            count = count + 1

            for cell in row:
                if (cell.value is None):
                    continue
                cell_data.append(cell.value)

            result_sheet.append(cell_data)

    result_xlsx.save("result.xlsx")