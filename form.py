import tkinter as tk
from tkinter import ttk,messagebox
import openpyxl
from openpyxl import Workbook
import pathlib

f=pathlib.Path('App.xlsx')
if f.exists():
    pass
else:
    file=Workbook()
    sh=file.active
    sh['A1']="First Name"
    sh['B1']="Last Name"
    sh['C1']="Phone No."
    sh['D1']="Email"
    sh['E1']="D.O.B"
    sh['F1']="Gender"
    sh['G1']="Guardian's Name"
    sh['H1']="Address"
    sh['I1']="Class 10"
    sh['J1']="Class 12"
    sh['K1']="Stream"

    file.save('App.xlsx')
    
def submit():
    fan=firstname.get()
    lan=lastname.get()
    p=phn.get()
    e=email.get()
    d=date.get()
    g=arr[gender.get()]
    gua=guardn.get()
    adr=address.get()
    c1m=c10.get()
    c2m=c12.get()
    s=st.get()

    l=[fan,lan,p,e,d,g,gua,adr,c1m,c2m,s]
    f=0
    for i in l:
        if i=='':
            f=1
    if f==1:
        messagebox.showwarning("Warning","Some places are left blank")
    else:

        file=openpyxl.load_workbook('App.xlsx')
        sh=file.active

        sh.append(l)

        file.save('App.xlsx')

        root.withdraw()
        messagebox.showinfo("Information","Details Added")
    

def clear():
    firstname.set('')
    lastname.set('')
    date.set('')
    gender.set(0)
    guardn.set('')
    address.set('')
    c10.set('')
    c12.set('')
    st.set('')

root=tk.Tk()
root.title("Application form")
root.resizable(False,False)
root.geometry('610x500')

l=tk.Label(root,text='Application Form',height=5,width=100,font=('arial',15,'bold','underline')).pack()

fn=tk.Label(root,text='First name',height=2,width=10,font=('arial',11)).place(x=10,y=100)
ln=tk.Label(root,text='Last name',height=2,width=10,font=('arial',11)).place(x=300,y=100)
phn=tk.Label(root,text='Phn No.',height=2,width=10,font=('arial',11)).place(x=10,y=150)
em=tk.Label(root,text='Email',height=2,width=10,font=('arial',11)).place(x=260,y=150)
dob=tk.Label(root,text='D.O.B',height=2,width=10,font=('arial',11)).place(x=10,y=200)
gen=tk.Label(root,text='Gender',height=2,width=10,font=('arial',11)).place(x=260,y=200)
gn=tk.Label(root,text="Guardiaan's name",height=2,width=15,font=('arial',11)).place(x=10,y=250)
ad=tk.Label(root,text='Present Address',height=2,width=14,font=('arial',11)).place(x=10,y=300)
m10=tk.Label(root,text='Class 10 Percentage',height=2,width=17,font=('arial',11)).place(x=10,y=350)
m12=tk.Label(root,text='Class 12 Percentage',height=2,width=17,font=('arial',11)).place(x=300,y=350)
s=tk.Label(root,text='Choose Stream',height=2,width=13,font=('arial',11)).place(x=10,y=400)


firstname = tk.StringVar()
lastname = tk.StringVar()
date=tk.StringVar()
gender=tk.IntVar()
guardn=tk.StringVar()
address=tk.StringVar()
c10=tk.IntVar()
c12=tk.IntVar()
st=tk.StringVar()
phn=tk.IntVar()
email=tk.StringVar()

e1=tk.Entry(root,textvariable=firstname,width=30).place(x=110,y=110)
e2=tk.Entry(root,textvariable=lastname,width=30).place(x=400,y=110)
e11=tk.Entry(root,textvariable=phn,width=25).place(x=110,y=160)
e22=tk.Entry(root,textvariable=email,width=38).place(x=355,y=160)
e3=tk.Entry(root,textvariable=date,width=15).place(x=110,y=210)

arr=['Male','Female','Others']
x1 = tk.Radiobutton(root,text=arr[0],variable=gender,value=0,font=('arial',10)).place(x=355,y=209)
x2 = tk.Radiobutton(root,text=arr[1],variable=gender,value=1,font=('arial',10)).place(x=425,y=209)
x3 = tk.Radiobutton(root,text=arr[2],variable=gender,value=2,font=('arial',10)).place(x=515,y=209)


e4=tk.Entry(root,textvariable=guardn,width=72).place(x=150,y=260)
e5=tk.Entry(root,textvariable=address,width=72).place(x=150,y=310)
e6=tk.Entry(root,textvariable=c10,width=10).place(x=175,y=360)
e7=tk.Entry(root,textvariable=c12,width=10).place(x=470,y=360)

c2=tk.ttk.Combobox(root,textvariable=st,width=12,values=['CSE','IT','ECE','EE'])
c2.place(x=150,y=410)
c2.current(1)

b1=tk.Button(root,text='Submit',command=lambda:submit(),font=('arial',10),bg='light blue').place(x=250,y=450)
b2=tk.Button(root,text='Clear',command=lambda:clear(),font=('arial',10),bg='light blue').place(x=340,y=450)


root.mainloop()
