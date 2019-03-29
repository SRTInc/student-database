from datetime import datetime
import xlrd
from xlutils import copy


def validate(DOB):

    try:
        datetime.strptime(DOB, "%d-%m-%Y")
        #raise ValueError
        #print('True')
        return True

    except ValueError:
        #print('False')
        return False


def agee(DOB):

        check = validate(DOB)

        if check == True:
            DOB = DOB.split('-')
            DOB=list(map(int,DOB))
           # for i in range(len(DOB)):
            #        temp = DOB[i]
             #z       DOB[i] = int(temp)

            Birth_year = DOB[2]
            current_date = datetime.now()
            current_year = current_date.year
            age = current_year - Birth_year
            return age

        else:
            print("Entered format is wrong, Re-Enter in this (dd-mm-yyy) format")
            agee()


def autoage(main):

    main = main
    workbook = xlrd.open_workbook("studbase.xls")
    sheet = workbook.sheet_by_index(0)
    d = sheet.cell_value(main, 5)
    return d


def infor(main):

    main = main
    info = ['Name', 'Exam Roll', 'Gender', 'DOB', 'Address', 'city', 'Postal code', 'State', 'Nationality',
            'Phone No.', 'E-mail', 'Father name', 'Father occupation', 'Mother Name', 'Mother occupation',
            'Annual Income', 'Age']
    loc = ("studbase.xls")
    wb = xlrd.open_workbook(loc)
    sheet = wb.sheet_by_index(0)
    print('\t\t\t\t\tProfile\n\n')

    for i in range(len(info)):
        print(info[i], ': ', sheet.cell_value(main, i + 2))


def newmarks(w_sheet,sp, wb, stud):

    a = input('Enter the Test Name  : ')
    w_sheet.write(0, 3 + sp, a)
    w_sheet.write(0, 9 + sp, 'NULL')  # to store the Null in the test heading column after the 9 positions

    subjects = []  # to store the entered subjects

    for o in range(7):

        if o == 6:

            w_sheet.write(1, o + sp, 'NULL')  # to store the Null in the subjects column after the subjects

        else:
            a = input('Enter the Subject Names: ')
            subjects.append(a)
            w_sheet.write(1, o + sp, a)

    wb.save('studbase.xls')
    return subjects


