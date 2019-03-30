from datetime import datetime
from modules import agee, autoage, infor, newmarks
import xlrd
from xlutils.copy import copy

print('\t\t\t\t\t\tWELCOME TO STUDENT DATABASE MANAGEMENT')
print('')
entry = input('\n\nPress \'ENTER\' to continue')
print('------------------------------------------------------------------------------------------')
wb = xlrd.open_workbook('studbase.xls')
cwb = copy(wb)
rsheet = wb.sheet_by_index(0)

sheet = cwb.get_sheet('Login')


def iput():

    if entry == '':

        print('\t\t\t\t\t\t\t\t\t\tLOGIN')
        user_name = input('User name:  ')  # to get the username
        pwd = input('Password :  ')  # to get the password
        found = False  # for verifying the user and pass line no:154 that ask to re enter the user and pass

        for i in range(rsheet.nrows):

            user_no = rsheet.cell_value(i, 0)  # to get the username from the exl sheet
            exl_user = (str(user_no)[:4])  # The user number from the exl is in float form (4001.0) so we take the first  4 values
            exl_pwd = str(rsheet.cell_value(i, 1))  # the password got from the excel

            main = i  # it is important to locate the  student or student data
            def pro(a, b):
                    user_name = a
                    pwd = b

                    if exl_user == user_name and pwd == exl_pwd:
                        found = True

                        print('\n------------------------------------------------------------------------------------------\n')
                        print('\t\t\t\tWelcome ', rsheet.cell_value(main, 2))
                        print('(1) Profile\n(2) Edit Profile\n(3) Marks\n(4) Change Password\n(5) Log out\n')
                        opt = input('\nEnter your option: ')

                        if opt == '1':  # Profile

                            rb = xlrd.open_workbook('studbase.xls')
                            r_sheet = rb.sheet_by_index(0)
                            chk = r_sheet.cell_value(main, 19) # it refers to column 19 in excel for all students in that column
                            #  the default value is False for everyone, because for checking the user first login or
                            #  more than one times logged in

                            if chk == False:  # if the 19 coloumn for the user value is False the user log in this account for
                                            # first time so we have to gather data from the user

                                print(
                                    '\n------------------------------------------'
                                    '------------------------------------------------\n')
                                print('\n\t\t\t\tStudent Information\n')
                                array = []  # Array for store the data entered by the user
                                attributes = ['Name', 'Exam rollno', 'Gender']

                                for i in range(len(attributes)):
                                    c = input('Enter the ' + attributes[i] + ' :')  # to get the data from the user
                                    array.append(c)  # Adding the user entered data to the array

                                DOB = input('Enter your Date of Birth(dd-mm-yyyy): ')
                                array.append(DOB)  # appending the DOB to the array
                                age = agee(DOB)  # call age function to calculate the age from the DOB
                                attributes = ['Address', 'City','Postal Code', 'State', 'Nationality',
                                              'Cell_phone', 'Email', 'Father name', 'Father Occupation', 'Mother Name',
                                              'Mother Occupation', 'Annual income']

                                for i in range(len(attributes)):
                                    c = input('Enter the ' + attributes[i] + ' :')
                                    array.append(c)

                                wob = copy(rb)
                                w_sheet = wob.get_sheet(0)

                                for i in range(len(array)):
                                    val = array[i]  # (r,c,value) # to get the user entered data from the array
                                    w_sheet.write(main, i + 2, val)  # and  store to the excel sheet

                                w_sheet.write(main, 18, age)  # this to store the age in the excel sheet
                                w_sheet.write(main, 19, 'True')  # After getting the information from the first time signed in
                                # it change the value in 19 th column to true, so the next time the user signed in the informat-
                                # ion gathering code is skipped by using the condition sttmt in line 47
                                w_sheet.write(main, 20,  str(datetime.now()))
                                wob.save('studbase.xls')  # after storing the values the excel sheet saved
                                pro(user_name,pwd)

                            else:

                                dob = autoage(main)  # to auto calculate the age, the user did not want to change the age
                                # for every year
                                rb = xlrd.open_workbook('studbase.xls')
                                cb = copy(rb)
                                w_sheet = cb.get_sheet(0)
                                DOB = dob.split('-')
                                DOB = list(map(int, DOB))
                                Birth_year = DOB[2]
                                current_date = datetime.now()
                                current_year = current_date.year
                                age = current_year - Birth_year
                                w_sheet.write(main, 18, age)
                                cb.save('studbase.xls')
                                infor(main)  # to show the profle of the signed in user
                                pro(user_name, pwd)

                        elif opt == '2':

                            print(
                                '\n------------------------------------------'
                                '------------------------------------------------\n')

                            print('\n\t\t\t\tProfile Editing\n')
                            info = ['1, Name', '\t2, Exam Roll', '\t3, Gender', '\n4, DOB', '\t5, Address', '\t6, city',
                                    '\n7, Postal code',
                                    '\t8, State', '\t9, Nationality', '\n10, Phone No.', '\t11, E-mail', '\t12, Father name',
                                    '\n13, Father occupation',
                                    '\t14, Mother Name', '\t15, Mother occupation', '\n16, Annual Income', '\t17, Age']

                            print(''.join(map(str, info)))  # for printing the list without square brackets and comma
                            info = ['Name', 'Exam Roll', 'Gender', 'DOB', 'Address', 'city', 'Postal code',
                                    'State', 'Nationality', '10, Phone No.', 'E-mail', 'Father name', 'Father occupation',
                                    'Mother Name', 'Mother occupation', 'Annual Income', ' Age']
                            e = input('how many options that u want to edit: ')
                            ed = input('\nEnter the options that you want to edit(with space in numbers): ')  # this get the options in number
                            edar = ed.split()  # to split the user entered options
                            edr = list(map(int, edar))  # to convert that splitted words
                            # print(edr)
                            rb = xlrd.open_workbook('studbase.xls')
                            wb = copy(rb)
                            w_sheet = wb.get_sheet(0)
                            ar = []  # to store the changed values
                            # main = 9

                            for i in edr:

                                edi = input('Enter ' + info[i - 1] + ': ')  # to get the values to change
                                ar.append(edi)  # append that values in 'ar' list

                            n = 0

                            for i in edr:

                                w_sheet.write(main, i + 1, ar[n])
                                n += 1

                            wb.save('studbase.xls')
                            pro(user_name, pwd)

                        elif opt == '3':

                            print('Marks')
                            tests = []
                            l_file = ('studbase.xls')

                            marks = xlrd.open_workbook(l_file)
                            marks = marks.sheet_by_index(1)
                            ini = 1

                            while 1:

                                tname = marks.cell_value(0, ini + 3)  # to get the test names
                                tests.append(tname)  # append it in the tests list

                                if tname == 'NULL':
                                    break

                                ini += 6

                            print('\t\t\tList of tests')
                            print(tests)
                            etest = input('Enter the name of the test : ')

                            no = 4  # to place the column in 4 becoz in evey column 4 the test name is stored in xl

                            while 1:

                                t_name = marks.cell_value(0, no)  # to get the test names

                                if t_name == etest:

                                    sub = -3
                                    subs = []
                                    loc = ("studbase.xls")
                                    wb = xlrd.open_workbook(loc)
                                    sheet = wb.sheet_by_index(1)

                                    for j in range(6):
                                        oii = no + sub  # to move the pointer to nxt column(sub)
                                        bla = sheet.cell_value(1, oii)  # get the sub name
                                        subs.append(bla)  # store the subject name
                                        sub += 1

                                    print(subs)

                                    sub = -3

                                    for j in range(6):
                                        oii = no + sub
                                        # ma = input('Enter the ' + subs[j] + ' mark ')
                                        mo = sheet.cell_value(main, oii)
                                        print(subs[j], ':', mo)
                                        sub += 1

                                    break

                                no += 6
                            pro(user_name, pwd)

                        elif opt == '4':

                            pswd = input('Enter the New password: ')
                            rpswd = input('Re-Enter the New password: ')
                            print(main)

                            if pswd == rpswd:
                                rb = xlrd.open_workbook('studbase.xls')
                                wb = copy(rb)
                                w_sheet = wb.get_sheet(0)
                                w_sheet.write(main, 1, rpswd)
                                wb.save('studbase.xls')
                                pro(user_name, pwd)



                        else:
                            iput()
                            #break
            pro(user_name, pwd)

            if user_name == 'admin' and pwd == str(rsheet.cell_value(1, 0 + 1)):

                found = True

                def aprof(a, b):

                    print('\t\t\t\tWelcome Admin')
                    print('(1) View Student Profile\n(2) Enter Students Mark\n(3) View Students Marks\n'
                          '(4) Change Password' '\n(5) Log out\n')
                    ch = input('Enter the option: ')

                    if ch == '1':

                        loc = ("studbase.xls")
                        wb = xlrd.open_workbook(loc)
                        sheet = wb.sheet_by_index(0)
                        vi = input('Enter the student Username to view profile: ')

                        for i in range(sheet.nrows):  # for iterating the all userno.

                            te = sheet.cell_value(i, 0)  # to get the user number from the xl sheet
                            ro = (str(te)[:4])
                            main = i  # it is important to view the data of correct student data

                            if vi == ro:  # to check the userno.

                                infor(main)  # to view the details of the student
                                aprof(user_name, pwd)

                    elif ch == '2':
                        def menu():

                            print('(1) New Entry\n(2) Update\n(3) Main Menu\n')
                            mk = input('Enter the option: ')

                            if mk == '1':

                                loc = ("studbase.xls")
                                wb = xlrd.open_workbook(loc)
                                sheet = wb.sheet_by_index(1)
                                found = False
                                init = 0  # to initialize the column
                                stud = sheet.nrows
                                # print(stud)

                                while 1:  # this loop is to avoid over writing

                                    a = sheet.cell_value(1, init)

                                    if a == 'NULL':  # if it find the  Null to stop the iteration
                                        break

                                    init += 1  # to store the Null position

                                rb = xlrd.open_workbook('studbase.xls')
                                wb = copy(rb)
                                w_sheet = wb.get_sheet(1)
                                print('New Entry')
                                subjects = newmarks(w_sheet, init, wb, stud)
                                rb = xlrd.open_workbook('studbase.xls')
                                wb = copy(rb)
                                w_sheet = wb.get_sheet(1)
                                for i in range(stud):
                                    no = 4001 + i
                                    print('Enter the mark of Student ' + str(no))
                                    for j in range(6):
                                        ma = input('Enter the ' + subjects[j] + ' mark ')
                                        w_sheet.write(2 + i, j + init, ma)
                                    if no == 4040:
                                        break
                                wb.save('studbase.xls')
                                menu()


                            elif mk == '2':

                                print('Update')
                                tests = []
                                l_file = ('studbase.xls')

                                marks = xlrd.open_workbook(l_file)
                                marks = marks.sheet_by_index(1)
                                ini = 1

                                while 1:

                                    tname = marks.cell_value(0, ini + 3)  # to get the test names
                                    tests.append(tname)  # append it in the tests list

                                    if tname == 'NULL':

                                        break

                                    ini += 6

                                print('\t\t\tList of tests')
                                print(tests)
                                etest = input('Enter the name of the test to update: ')
                                sname = input('Enter the student roll no. to edit: ')

                                for i in range(marks.nrows):

                                    te = marks.cell_value(i, 0)
                                    ro = (str(te)[:4])  # to get the student user no. from the xl
                                    main = i  # it is important to save the data in correct student data

                                    if sname == ro:  # to check the stud user no.

                                        no = 4  # to place the column in 4 becoz in evey column 4 the test name is stored in xl

                                        while 1:

                                            t_name = marks.cell_value(0, no)  # to get the test names

                                            if t_name == etest:

                                                sub = -3
                                                subs = []
                                                loc = ("studbase.xls")
                                                wb = xlrd.open_workbook(loc)
                                                sheet = wb.sheet_by_index(1)

                                                for j in range(6):

                                                    oii = no + sub  # to move the pointer to nxt column(sub)
                                                    bla = sheet.cell_value(1, oii)  # get the sub name
                                                    subs.append(bla)  # store the subject name
                                                    sub += 1

                                                print(subs)
                                                rb = xlrd.open_workbook('studbase.xls')
                                                wb = copy(rb)
                                                w_sheet = wb.get_sheet(1)
                                                sub = -3

                                                for j in range(6):

                                                    oii = no + sub
                                                    ma = input('Enter the ' + subs[j] + ' mark ')
                                                    w_sheet.write(main, oii, ma)
                                                    sub += 1

                                                wb.save('studbase.xls')
                                                break

                                            no += 6  # move to next test
                                menu()
                            else:

                                print('Main menu')
                                aprof(user_name, pwd)
                        menu()


                    elif ch == '3':

                        print('View mark')
                        aprof(user_name, pwd)

                    elif ch == '4':


                        def cpwd():

                            pswd = input('Enter the New password: ')
                            rpswd = input('Re-Enter the New password: ')


                            if pswd == rpswd:
                                rb = xlrd.open_workbook('studbase.xls')
                                wb = copy(rb)
                                w_sheet = wb.get_sheet(0)
                                w_sheet.write(1, 1, rpswd)
                                wb.save('studbase.xls')
                            else:
                                print('Password does not matched, Re-enter the new password\n')
                                cpwd()
                        cpwd()
                        pro(user_name, pwd)


                    else:
                        iput()
                aprof(user_name,pwd)





        if found is False:

            print('User name or password or both are wrong!!!\n')
            print('If you want to renter the login details or  \'Exit\'\n')
            cont = input('Options\n\n1, Login\n2, Exit\n\nChoose your option: ')

            if cont == '1':

                iput()

            else:

                print('\n\t\t\t\tThank You, Good Bye')
                exit()


iput()
def exit():
    print('---x-----x----')

