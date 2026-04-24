import xlwings as xw

def main():
    wb = xw.Book.caller()
    sheet = wb.sheets[0]  # sheet명을 입력해도 됨

    # 셀입력
    sheet["A1"].value = "xlwings 셀입력1"
    sheet.range(2,1).value = "xlwings 셀입력2"    
    sheet.range(3,1).value = ['리스트1','리스트2','리스트3']
    sheet.range(4,1).value = [['리스트1','리스트2','리스트3'],['리스트1','리스트2','리스트3']]
    sheet.range(6,1).options(transpose=True).value = ['리스트1','리스트2','리스트3']
    sheet.range(9,1).value = {'딕트':1}



if __name__ == "__main__":
    xw.Book("testwings.xlsm").set_mock_caller()
    main()