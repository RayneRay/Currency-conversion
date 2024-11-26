import requests
from bs4 import BeautifulSoup
import openpyxl


excel = openpyxl.load_workbook("Calculation table.xlsx")

excel_act = excel.active

url = "https://cbr.ru/currency_base/daily/"

page = requests.get(url)

soup = BeautifulSoup(page.text, 'lxml')

table1 = soup.find('tbody')

for j in table1.find_all("tr")[1:]:
    """
    Parsing currency rates on the website and displaying them rounded to one decimal place
    """
    row_data = j.find_all("td")
    row = [i.text for i in row_data]
    if 'EUR' in row:
        EUR_value = row[4]
        EUR_value = EUR_value.replace(",", ".")
        EUR_value = round(float(EUR_value), 1)
    if 'USD' in row:
        USD_value = row[4]
        USD_value = USD_value.replace(",", ".")
        USD_value = round(float(USD_value), 1)
    if 'CNY' in row:
        CNY_value = row[4]
        CNY_value = CNY_value.replace(",", ".")
        CNY_value = round(float(CNY_value), 1)


def table_sum_ru(num, num1, num2, value):
    """
    Conversion from local currencies to local currency (rubles) in accordance
    with the exchange rate (from one table to another)
    """
    for cell in excel_act[num2][num1:num1+1]:
        if num1 == 13:
            break
        else:
            try:
                cell.value = excel_act[num][num1].value * value
                excel_act["A13"].value = None
                table_sum_ru(num=num, num1=num1 + 1, num2=num2, value=value)
            except TypeError:
                excel_act["A13"].value = (
                "Вы ввели не числовые значения либо же оставили пустые ячейки. Пожалуйста, введите число, "
                "в ином случае, введите '0'")
                break



table_sum_ru(num=3, num1=1, num2=9, value=EUR_value)
table_sum_ru(num=4, num1=1, num2=10, value=USD_value)
table_sum_ru(num=5, num1=1, num2=11, value=CNY_value)



def sum_value(num1, summary, num2=1, value=14):
    """
    Converting the amount of local income into local currency (rubles) from one table to another
    """
    for cell in excel_act[num1][value:value+1]:
        if num2 == 13:
            cell.value = summary
            break
        else:
            try:
                summary += excel_act[num1][num2].value
                sum_value(num1=num1, num2=num2 + 1, value=value, summary=summary)
            except TypeError:
                excel_act["A13"].value = (
                "Вы ввели не числовые значения либо же оставили пустые ячейки. Пожалуйста, введите число, "
                "в ином случае, введите '0'")



sum_value(num1=3, summary=0)
sum_value(num1=4, summary=0)
sum_value(num1=5, summary=0)

sum_value(num1=9, summary=0)
sum_value(num1=10, summary=0)
sum_value(num1=11, summary=0)


def sum_value_all(num1, summary, num4, const, num2=1, value=16):
    """
    Calculation of the entire amount of income in local currency (rubles) and entry into the table
    """
    for cell in excel_act[num4][value:value+1]:
        if num1 == const+3:
            cell.value = summary
            break
        elif num2 == 12:
            cell.value = summary
            sum_value_all(num1=num1+1, num2=1, value=value, summary=summary, num4=num4, const=const)
        else:
            try:
                summary += excel_act[num1][num2].value
                sum_value_all(num1=num1, num2=num2 + 1, value=value, summary=summary, num4=num4, const=const)
            except TypeError:
                excel_act["A13"].value = (
                "Вы ввели не числовые значения либо же оставили пустые ячейки. Пожалуйста, введите число, "
                "в ином случае, введите '0'")


sum_value_all(num1=9, summary=0, num4=10, const=9)


excel.save("Calculation table.xlsx")

