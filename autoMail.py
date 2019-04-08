import xlrd
import smtplib
from tkinter import *

def send_mail():
    file = file_name.get()
    send_mail = sender_mail.get()
    send_pass = sender_pass.get()
    
    wb = xlrd.open_workbook(file)
    sheet = wb.sheet_by_index(0)
    w = sheet.ncols 
    h = sheet.nrows 
    print(w)
    print(h)

    Matrix = [[0 for x in range(h)] for y in range(w)]
    sheet.cell_value(0,0)


    arr = []
    for i in range(h):
        arr.append([])
        for j in range(w):
            arr[i].append(sheet.cell_value(i,j))

    #for i in range(h):
    #    for j in range(w):
    #        print(arr[i][j])

    for i in range(1,h):
        print(arr[i][1])
        s = smtplib.SMTP('smtp.gmail.com', 587) 
        s.starttls() 
        s.login(send_mail,send_pass) 
        message = "This is the Test Mail " +arr[i][0]
        s.sendmail("pkasani1@gmail.com", arr[i][1], message) 
        s.quit()
        
        for i in range(1,h):
            label = Label(window,text="Mail to %s has been sent" %arr[i][0])
            label.grid(row=9+i,column=3,padx=5, pady=5)
        e1.delete(0,END)
        e2.delete(0,END)
        e3.delete(0,END)
            
window = Tk()
window.title('Mail Sender')

l1 = Label(window,text="Enter Excel file name (.xlsx) format")
l1.grid(row=2,column=2,padx=10, pady=10)

file_name = StringVar()
e1 = Entry(window,textvariable=file_name)
e1.grid(row = 2, column = 4,padx=10, pady=10)

l1 = Label(window,text="Sender email")
l1.grid(row=4,column=2,padx=10, pady=10)

sender_mail = StringVar()
e2 = Entry(window,textvariable=sender_mail)
e2.grid(row = 4, column = 4,padx=10, pady=10)

l1 = Label(window,text="Sender password")
l1.grid(row=6,column=2,padx=10, pady=10)

sender_pass = StringVar()
e3 = Entry(window,textvariable=sender_pass,show='*')
e3.grid(row = 6, column = 4,padx=10, pady=10)

b1 = Button(window,text="Submit",width=10,command=send_mail)
b1.grid(row = 8, column = 3,padx=10, pady=10)
window.mainloop()

print(file_name.get())