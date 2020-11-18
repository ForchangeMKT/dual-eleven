import random
import time
import xlrd


class people:
    code = ''

    def __init__(self, c):
        self.code = c


def load_excel():
    work_book = xlrd.open_workbook('2020双11用户抽奖序列码.xlsx')
    sheet = work_book.sheet_by_name('SheetJS')
    data = []
    for a in range(sheet.nrows):
        cells = sheet.row_values(a)
        data.append(people(cells[0]))
    return data


if __name__ == '__main__':
    users = load_excel()
    limit = [(25, "幸运奖"), (3, "三等奖"), (2, "二等奖"), (1, "一等奖")]

    print("========================================")
    print("****************************************")
    print("***         2020 风变双11            ***")
    print("****************************************")
    print("========================================")
    print("Are you ready?")
    for i in range(4):
        time.sleep(1)
        print(3 - i)

    for rounds in limit:
        print("========================================")
        print(rounds[1], ":")
        for j in range(8888):
            print('\r', users[j].code, end='  ')
            time.sleep(0.004)
        print("\r*********")
        for index in range(0, rounds[0]):
            lucky_dog = random.randint(1, len(users))
            print('\r', users[lucky_dog].code)
            users.remove(users[lucky_dog])
        print("*********")
        print("winner winner chicken dinner\nCongratulation!!!!!!!!!!!!!!")
        print("========================================")
        if rounds[0] == 1:
            print("See u")
        else:
            input("next?")
