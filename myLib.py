from os import listdir
from openpyxl import load_workbook, Workbook


# 이벤트 조건과 맞을때 셀에 추가
def check_Inkwd(value):
    IN_KEYWORDS = ["www", "facebook", "청해부대"]

    for word in IN_KEYWORDS:
        if (not (word in value)):
            break

        if (word is IN_KEYWORDS[-1]):
            return True

    return False


# 체리피커를 걸러내는 함수
def check_outkwd(value):
    OUT_KEYWORDS = ["초대", "함께해", "함께해요"]

    # 체리피커 키워드가 있을때 거른다
    for word in OUT_KEYWORDS:
        if (word in value):
            return False

    return True
