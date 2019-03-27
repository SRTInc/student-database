from datetime import  datetime
from modules import agee,autoage,infor
import xlutils, xlrd, xlwt
from xlutils.copy import copy

print('\t\t\t\t\t\tWELCOME TO STUDENT DATABASE MANAGEMENT')
a = input('\n\nPress \'ENTER\' to continue')
print('------------------------------------------------------------------------------------------')
wb = xlrd.open_workbook('studbase.xls')
cwb = copy(wb)
rsheet = wb.sheet_by_index(0)

sheet = cwb.get_sheet('Login')


def iput():

    if a == '':

        print('\t\t\t\t\t\t\t\t\t\tLOGIN')
        user_name = input('User name:  ')
        pwd = input('Password :  ')
        found = False  # for verifying the user and pass line no:154 that ask to re enter the user and pass

        for i in range(rsheet.nrows):

            te = rsheet.cell_value(i, 0)
            se = (str(te)[:4])
            p = str(rsheet.cell_value(i, 1))
            main = i  # it is important to save the data in correct student data

            if se == user_name and pwd == p:

                print('\n------------------------------------------------------------------------------------------\n')
                print('\t\t\t\tWelcome ',rsheet.cell_value(main,2))
                print('(1) Profile\n(2) Edit Profile\n(3) Marks\n(4) Log out')
                opt = input('\nEnter your option: ')

                if opt == '1':

                    rb = xlrd.open_workbook('studbase.xls')
                    r_sheet = rb.sheet_by_index(0)
                    chk = r_sheet.cell_value(main, 19)


                    if chk == False:

                        print(
                            '\n------------------------------------------'
                            '------------------------------------------------\n')
                        print('\n\t\t\t\tStudent Information\n')
                        array = []
                        attributes = ['Name', 'Exam rollno', 'Gender']

                        for i in range(len(attributes)):
                            c = input('Enter the ' + attributes[i] + ' :')
                            array.append(c)

                        DOB = input('Enter your Date of Birth(dd-mm-yyyy): ')
                        array.append(DOB)
                        age = agee(DOB)
                        attributes = ['Address', 'City','Postal Code', 'State', 'Nationality',
                                      'Cell_phone', 'Email', 'Father name', 'Father Occupation', 'Mother Name',
                                      'Mother Occupation', 'Annual income']

                        for i in range(len(attributes)):
                            c = input('Enter the ' + attributes[i] + ' :')
                            array.append(c)

                        wob = copy(rb)
                        w_sheet = wob.get_sheet(0)

                        for i in range(len(array)):
                            val = array[i]  # (r,c,value)
                            w_sheet.write(main, i + 2, val)

                        w_sheet.write(main, 18, age)
                        w_sheet.write(main, 19, 'True')
                        w_sheet.write(main, 20,  str(datetime.now()))
                        wob.save('studbase.xls')

                    else:

                        dob = autoage(main)
                        rb = xlrd.open_workbook('studbase.xls')
                        cb = copy(rb)
                        w_sheet = cb.get_sheet(0)
                        DOB = dob.split('-')
                        DOB = list(map(int, DOB))
                        Birth_year = DOB[2]
                        current_date = datetime.now()
                        current_year = current_date.year
                        current_year = 2020
                        age = current_year - Birth_year
                        w_sheet.write(main, 18, age)
                        cb.save('studbase.xls')
                        infor(main)

                elif opt == '2':

                    print('Edit')

                    print('\t\t\t\tProfile Editing\n\n')
                    info = ['1, Name', '\t2, Exam Roll', '\t3, Gender', '\n4, DOB', '\t5, Address', '\t6, city',
                            '\n7, Postal code',
                            '\t8, State', '\t9, Nationality', '\n10, Phone No.', '\t11, E-mail', '\t12, Father name',
                            '\n13, Father occupation',
                            '\t14, Mother Name', '\t15, Mother occupation', '\n16, Annual Income', '\t17, Age']

                    print(''.join(map(str, info)))  # for printing the list without square brackets and comma
                    info = ['Name', 'Exam Roll', 'Gender', 'DOB', 'Address', 'city', 'Postal code',
                            'State', 'Nationality', '10, Phone No.', 'E-mail', 'Father name', 'Father occupation',
                            'Mother Name', 'Mother occupation', 'Annual Income', ' Age']
                    e = input('how many options that u want to edit:')
                    ed = input('\nEnter the options that you want to edit(with space): ')
                    edar = ed.split()
                    edr = list(map(int, edar))
                    print(edr)
                    rb = xlrd.open_workbook('studbase.xls')
                    wb = copy(rb)
                    w_sheet = wb.get_sheet(0)
                    ar = []
                    main = 9

                    for i in edr:
                        edi = input('Enter ' + info[i - 1] + ': ')
                        ar.append(edi)
                        n = 0

                    for i in edr:
                        w_sheet.write(main, i + 1, ar[n])
                        n += 1

                    wb.save('studbase.xls')

                elif opt == '3':

                    print('Marks')

                else:

                    iput()

                found = True
                break

            if user_name == 'admin' and str(rsheet.cell_value(1, 0 + 1)):

                print('\t\t\t\tWelcome Admin')
                print('(1) View Student Profile\n(2) Enter Students Mark\n(3) View Students Marks\n(4) Log out\n')
                ch = input('Enter the option: ')

                if ch == str(1) :

                    loc = ("studbase.xls")
                    wb = xlrd.open_workbook(loc)
                    sheet = wb.sheet_by_index(0)
                    vi = input('Enter the student Username to view profile: ')

                    for i in range(sheet.nrows):

                        te = sheet.cell_value(i, 0)
                        ro = (str(te)[:4])
                        main = i  # it is important to save the data in correct student data

                        if vi == ro :

                            infor(main)

                elif ch == 2:

                    print('(1) New Entery\n(2) Update\n(3) Main Menu\n')
                    mk = input('Enter the option: ')

                    if mk == 1:

                        print('New entery')

                    elif mk == 2:

                        print('Update')

                    else:

                        print('Main menu')

                elif ch == 3:
                    print('View mark')

                else:
                    iput()

                found = True
                break

        if found == False:
            print('User name or password or both are wrong!!!\n')
            print('If you want to renter the login details or  \'Exit\'\n')
            cont = input('Options\n\n1, Login\n2, Exit\n\nChoose your option: ')
            if cont == '1':
                iput()
            else:
                print('\n\t\t\t\tThank You, Good Bye')


iput()

