import pandas as pd
import csv
from openpyxl import load_workbook
from datetime import datetime
import time
import sys

start_time = datetime.now()


def get_letter(sheet_load):
    flg = True
    letters = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U',
               'V', 'W',
               'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO',
               'AP', 'AQ',
               'AR', 'AS', 'AT', 'AU', 'AV', 'AW', 'AX', 'AY', 'AZ']
    count = 0
    while flg:
        value = (sheet_load[f'{letters[count]}1'].value)
        if str(value) == '21vekOnliner':
            flg = False
            return letters[count]
        count += 1


def get_id_list(sheet_load):
    flg = True
    artic = []
    count = 2
    letter = get_letter(sheet_load)

    while flg:
        value = (sheet_load[f'{letter}{count}'].value)
        if str(value) != 'None':
            artic.append(str(value).replace(' ', '').replace('.', ''))
        if str(sheet_load[f'A{count + 1}'].value) == 'None':
            flg = False
        count += 1
    return artic


try:
    print('Загрузка... (примерно 2 минуты)')
    unload = load_workbook('source/main.xlsx')
    sheet_load = unload['Лист 1']
    artic = get_id_list(sheet_load)
except FileNotFoundError:
    print('Файл не найден. Проверьте папку source. Файл должен называться main.xlsx')
    time.sleep(5)
    sys.exit()
except KeyError:
    print('Ошибка. Переименуйте лист в файле main.xlsx в "Лист 1"')
    time.sleep(5)
    sys.exit()

current_datetime = str(datetime.now()).replace(' ', '-').replace(':', '-').replace('.', '-')
df = pd.DataFrame()
filename = f'{str(current_datetime)}.csv'
supp_id = []

try:
    csvfile_read = open(f'output/{filename}', 'w', newline='')
    csfvile_read1 = open('output/21vekPars.csv', 'w', newline='')
    writer = csv.writer(csvfile_read, delimiter=';')
    writer1 = csv.writer(csfvile_read1, delimiter=';')
    with open('source/file1.csv', 'r', newline='') as csvfile:
        letters = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T',
                   'U',
                   'V', 'W', 'X', 'Y', 'Z']
        flg = True
        spamreader = csv.reader(csvfile)
        reader = csv.reader(csvfile)
        for row in spamreader:
            strin = ((''.join(row)).split(';'))
            if flg:
                strin = [i.replace('"', '').lower() for i in strin]
                writer.writerow(
                    (
                        strin[0], strin[1], strin[2], strin[3], strin[4], strin[5], strin[6], strin[7], strin[8],
                        strin[9],
                        strin[10], strin[11], strin[12], strin[13], strin[14], strin[15], strin[16], strin[17],
                        strin[18], strin[19],
                        strin[20], strin[21], strin[22], strin[23]
                    )
                )
                writer1.writerow(
                    (
                        strin[0], strin[1], strin[2], strin[3], strin[4], strin[5], strin[6], strin[7], strin[8],
                        strin[9],
                        strin[10], strin[11], strin[12], strin[13], strin[14], strin[15], strin[16], strin[17],
                        strin[18], strin[19],
                        strin[20], strin[21], strin[22], strin[23]
                    )
                )
                flg = False
            else:
                supp_id.append(strin[0])

    answer = list(set(supp_id) & set(artic))
    artic.clear()
    supp_id.clear()

    with open('source/file1.csv', 'r', newline='') as csvfile:

        spamreader = csv.reader(csvfile)
        print('Создание файла... Пожалуйста, не выключайте программу. (Примерное время ожидания 5-10 минут)')
        print('Окно закроется само. Готовый файл будет в папке output')
        time.sleep(10)
        for row in spamreader:
            pass
            break
        for row in spamreader:
            strin = ''
            for i in range(len(row) - 1):
                strin += ''.join(row[i]) + ','
            strin += ''.join(row[len(row) - 1])
            strin = strin.split(';')
            if str(strin[0]) in answer:
                del answer[answer.index(str(strin[0]))]
                for i in range(0, 24):
                    if strin[i] and strin[i][0] == '"':
                        strin[i] = strin[i][1: -1]
                writer.writerow(
                    [strin[i] for i in range(0, 24)]
                )
                writer1.writerow(
                    [strin[i] for i in range(0, 24)]
                )

except:
    print('Файл не найден. Проверьте папку. Файл должен называться file1.csv')
    time.sleep(5)
    sys.exit()

print('Файл создан. Проверьте папку "output"')
time.sleep(5)
print(datetime.now() - start_time)
