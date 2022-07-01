
# main.py


import Excel2Html

if __name__ == '__main__':
    a =Excel2Html.Excel2Html("b.xls")
    print(a)
    Excel2Html.writeHtml(a)