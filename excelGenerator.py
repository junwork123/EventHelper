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

def copyRange(startCol, startRow, endCol, endRow, sheet):
    rangeSelected = []
    # Loops through selected Rows
    for i in range(startRow, endRow + 1, 1):
        # Appends the row to a RowSelected list
        rowSelected = []
        for j in range(startCol, endCol + 1, 1):
            rowSelected.append(sheet.cell(row=i, column=j).value)
        # Adds the RowSelected List and nests inside the rangeSelected
        rangeSelected.append(rowSelected)

    return rangeSelected

def pasteRange(startCol, startRow, endCol, endRow, sheetReceiving, copiedData):
    countRow = 0
    for i in range(startRow, endRow + 1, 1):
        countCol = 0
        for j in range(startCol, endCol + 1, 1):
            sheetReceiving.cell(row=i, column=j).value = copiedData[countRow][countCol]
            countCol += 1
        countRow += 1

def generateExcel(myEvent):

    # 키워드 세팅
    setOptions(myEvent)

    files = listdir(".")
    src_xlsx = load_workbook("src.xlsx")
    src_sheet = src_xlsx.active

    src_sheet.delete_cols(8, 3)  # Tagged User Count on This Comment, Tagged User Count in Entire Post
    src_sheet.delete_cols(1, 3)  # Post Reaction, ID, Comment Link


    result_xlsx = Workbook()
    # 조건에 맞는 사람들을 뽑아내는 리스트 시트
    list_sheet = result_xlsx. active
    list_sheet['A1'] = "No"
    list_sheet['B1'] = "이름"
    list_sheet['C1'] = "날짜"
    list_sheet['D1'] = "댓글 내용"
    list_sheet['E1'] = "좋아요 수"

    # 추첨 시트 생성
    raffle_sheet = result_xlsx.create_sheet("추첨페이지")

    raffle_sheet['A1'] = "No"
    raffle_sheet['B1'] = "이름"
    raffle_sheet['C1'] = "난수값"
    raffle_sheet['D1'] = "순위"
    raffle_sheet['E1'] = "당첨자"
    raffle_sheet['F1'] = "상품"

    count = 1
    for row in src_sheet.iter_rows():

        cell_data = []

        # 첫째칸이 비어있거나, Text가 비어있다면 패스
        if (row[0].value is None or row[2].value is None):
            continue

        # 이벤트 조건과 맞을때 셀에 추가
        if (check_Keyword(row[2].value) and check_Cherry(row[2].value)):
            # 두번째 줄(A2) 부터 데이터 복사
            cell_data.append(count)
            count = count + 1

            for cell in row:
                if (cell.value is None):
                    continue
                cell_data.append(cell.value)

            list_sheet.append(cell_data)

    # 'No.', '이름' 복사-붙여넣기
    copiedData = copyRange(1,2,2,count,list_sheet)
    pasteRange(1,2,2,count,raffle_sheet,copiedData)

    # '난수값' 열에 난수값 함수 삽입
    for i in range(2,count):
        raffle_sheet.cell(row = i, column = 3).value = "=RAND()"
        if( i < 2 + 70 ):
            raffle_sheet.cell(row = i, column = 4).value = i - 1
            raffle_sheet.cell(row = i, column = 5).value = "=VLOOKUP(D"+str(i)+",$A$2:$C$"+str(count)+",2,0)"
    result_xlsx.save("result.xlsx")