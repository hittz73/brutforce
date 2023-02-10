from string import digits, punctuation, ascii_letters
import itertools
import win32com.client as client
from datetime import datetime
import time


#pip install pywin32

#password = list('1234567890!@#$%^&*()ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz')
#n = int(input("Введите количество символов в пароле: "))
#number_kolvo = int(input("Введите количество паролей: "))
#random.shuffle(password)

# for i in range(number_kolvo):
#      #перемешивает последовательность (изменяется сама последовательность). Поэтому функция не работает для неизменяемых объектов.
#     password = ''.join([random.choice(password) for x in range(n)])
#     #random.choice - случайный элемент непустой последовательности.
#     print(password)
#     f = open('passwd.txt','a') #запись в файл. режим 'a' обозначает дозапись в файл
#     f.write(password + '\n')
#     f.close()


# symbols = digits + punctuation + ascii_letters
# print(symbols)

def brute_excel_doc():
    print("Hello")
    try:
        password_lenght = input("Введите длину пароля, от скольки до скольки:")
        password_lenght = [int(item) for item in password_lenght.split('-')] #разбили строку по дефису, int приводит к числу
    except:
        print("Проверьте введеные данные")

    print("Если пароль содержит только цифры, введите: 1 \nЕсли пароль содержит только буквы, введите: 2\n"
          "Если пароль содержит цифры и буквы, введите: 3\nЕсли пароль содержит цифры,буквы и спец.символы введите: 4")
    try:
        choice = int(input(": "))
        if choice == 1:
            possible_symbols = digits
        elif choice == 2:
            possible_symbols = ascii_letters
        elif choice == 3:
            possible_symbols = digits + ascii_letters
        elif choice == 4:
            possible_symbols = digits + ascii_letters + punctuation
        else:
            possible_symbols = "Выберите вариант с 1 по 4"
        print(possible_symbols)
    except:
        possible_symbols = "Выберите вариант с 1 по 4"


    #brute excel doc
    start_timestamp = time.time() #время начала
    print(f"Старт в  - {datetime.utcfromtimestamp(time.time()).strftime('%H:%M:%S:')}")
    count = 0 #счетчик попыток
    for pass_lenght in range(password_lenght[0], password_lenght[1] + 1):
        for password in itertools.product(possible_symbols, repeat = pass_lenght):   #делаем послед. вложенных циклов
            password = "".join(password)
            #print(password)

            opened_doc = client.Dispatch("Excel.Application")
            count += 1

            try:
                opened_doc.WorkBooks.Open(          #открываем файл экселя
                    r"C:\Users\Денис\PycharmProjects\bruteforce\test.xlsx",
                    False,
                    True,
                    None,
                    password
                )
                time.sleep(0.1)
                print(f"Закончил в - {datetime.utcfromtimestamp(time.time()).strftime('%H:%M:%S:')}")
                print(f"Скрипт работал - {time.time() - start_timestamp}")
                return f"Попытки #{count} Пароль: {password}"
            except:
                print(f"Попытки #{count} Неверный пароли {password}")
                pass

def main():
    brute_excel_doc()

if __name__ == '__main__':
    main()