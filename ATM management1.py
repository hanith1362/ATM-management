import xlwt
import xlrd
from xlutils.copy import copy
import smtplib
server=smtplib.SMTP('smtp.gmail.com',587)
server.starttls()
server.login('pythonmail222@gmail.com','RockyRocky222')
workbook=xlwt.Workbook()
sheet=workbook.add_sheet("admin sheet")
path="C:/Users/hanith/Documents/atmdata.xls"
path2="C:/Users/hanith/Documents/atmdatauser.xls"
print("press 1 to open admin sheet\n","press 2 to open user sheet")
n=int(input("enter the number for above operation:"))
if(n==1):
    print("you entered into admin sheet:")
    n1=input("enter the username:")
    dict1={'username':"password",'hanith':"9989","sai":"1234","krishna":"krishna123"}
    usernames=list(dict1.keys())
    password=list(dict1.values())
    row=0
    while(row<len(usernames)):
        column=0
        sheet.write(row,column,usernames[row])
        column+=1
        sheet.write(row,column,password[row])
        row+=1
    workbook.save(path)
    if n1 in usernames:
        m=usernames.index(n1)
        print("username is valid")
        n2=int(input("enter the admin number:"))
        w=xlrd.open_workbook(path)
        sheet=w.sheet_by_index(0)
        n3=sheet.cell_value(n2,1)
        n4=input("enter the password:")
        if n3==n4:
            print("password is valid")
            print("press 1 to add customer\n press 2 for Remove customer\n press 3 for details of customer\n")
            m=int(input("enter the number for above operation:"))
            if(m==1):
                print('Enter the details of customer below\n')
                username1=input('Enter the username:')
                userpassword=input('Enter the password:')
                balance=input("enter the initial amount: ")
                mail=input("enter the mail id:")
                workbook1=xlrd.open_workbook(path2)
                sheet1=workbook1.sheet_by_index(0)
                h=copy(workbook1)
                a=h.get_sheet(0)
                a.write((sheet1.nrows)+1,0,username1)
                a.write((sheet1.nrows)+1,1,userpassword)
                a.write((sheet1.nrows)+1,2,balance)
                a.write((sheet1.nrows)+1,3,mail)
                h.save(path2)
                print("details are added successfully")
            if(m==2):
                workbook1=xlrd.open_workbook(path2)
                sheet1=workbook1.sheet_by_index(0)
                usernames=[sheet1.cell_value(i,0) for i in range(1,sheet1.nrows)]
                r=input(('Enter the username of the to delete from data base:'))
                if r in usernames:
                    p=usernames.index(r)
                    p=p+1
                    d=copy(workbook1)
                    q=d.get_sheet(0)
                    q.write(p,0,'')
                    q.write(p,1,'')
                    q.write(p,2,'')
                    q.write(p,3,'')
                    d.save(path2)
                    print('Data about the user is deleted successfully\n')
                else:
                    print('entered username is not in the list')
            if(m==3):
                print('customers details \n')
                workbook1=xlrd.open_workbook(path2)
                sheet1=workbook1.sheet_by_index(0)
                for count in range(sheet1.nrows):
                    print(sheet1.row_values(count))
        else:
            print("password is invalid")
    else:
        print("username is invalid")
if(n==2):
    workbook1=xlrd.open_workbook(path2)
    sheet1=workbook1.sheet_by_index(0)
    usernames=[sheet1.cell_value(i,0) for i in range(1,sheet1.nrows)]
    username=input("enter the username:")
    if username in usernames:
        m=usernames.index(username)
        password=input('Enter the password:')
        if(password==sheet1.cell_value(m+1,1)):
            print('Loginsucessfully')
            print('1.Check Your Balance\n2.Withdraw\n3.Change Password\n')
            b=int(input('enter the number for above operations:'))
            if(b==1):
                bal=str(sheet1.cell_value(m+1,2))
                print('Your Account Balance:'+bal)
                server.sendmail('pythonmail222@gmail.com',sheet1.cell_value(m+1,3),'This is a Mail From sbi Bank !\nGenerated due to your action at ATM to Check Yourbalance.\n Your Balance is:RS '+str(bal))
            if(b==2):
                rd=copy(workbook1)
                d=rd.get_sheet(0)
                bal=sheet1.cell_value(m+1,2)
                bal1=int(bal)
                wd=float(input('Enter the amount you want to withdraw:'))
                c=bal1-wd
                d.write(m+1,2,c)
                rd.save(path2)
                print('Amount Withdrawn Sucessful!\n')
                server.sendmail('otisola222@gmail.com',sheet1.cell_value(m+1,3),'This is a mail from SBI bank!\n to intimate that Your account has been debited with amount '+str(wd)+'\nupdated balance is '+str(c))
            if(b==3):
                while(1):
                    pass1=input(('Enter Your old Password:'))
                    if(pass1==sheet1.cell_value(m+1,1)):
                        while(1):
                            pass2=input('Enter Your New Password:')
                            pass3=input('confirm new password:')
                            if(pass2==pass3):
                                rd=copy(workbook1)
                                d=rd.get_sheet(0)
                                d.write(m+1,1,pass2)
                                rd.save(path2)
                                print('Password Changed Successful!Now you Can Login With New Password\n')
                                server.sendmail('pythonmail222@gmail.com',sheet1.cell_value(m+1,3),'This is the mail from SBI Bank to intimate you that your password has been changed')
                                break
                            else:
                                print('Password do not match!')
                        break
                        
                    else:
                        print("your old password is incorrect")
                        
        else:
            workbook1=xlrd.open_workbook(path2)
            sheet1=workbook1.sheet_by_index(0)
            usernames=[sheet1.cell_value(i,0) for i in range(1,sheet1.nrows)]
            server.sendmail('pythonmail222@gmail.com',sheet1.cell_value(m+1,3),'This mail is from SBI bank\n some one is trying to login to your account please contact near sbi branch')
            print("password is incorrect!!!!")
            
    else:
        print("username is not in the sheet")