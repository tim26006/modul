from tkinter import *

from tkinter.ttk import Radiobutton
from tkinter import scrolledtext
import openpyxl
from tkinter import messagebox

running=True
def run():
    def on_closing():

        global running
        running = False
        tk.destroy()  # Закрыть окно



    def clicked():


            notification = Label(tk, text = 'Какого сотрудника вы хотите добавить в список?\nВведите ФИО', font = ('Arial Bold',14))
            notification.grid(column =0,row=3)

            text = Entry(tk, width=75)
            text.grid(column =0, row =4)

            def clicked():
                #Имя
                name = '{}'.format(text.get())
                a = [name]

                notification = Label(tk, text='Введите ID', font=('Arial Bold', 14))
                notification.grid(column=0, row=6)
                text1 = Entry(tk, width=75)
                text1.grid(column=0, row=7)
                def clicked1():
                    id = '{}'.format(text1.get())
                    a.append(id)

                    for i in wb.worksheets:  # перебираю таблицы
                        i.append(a)


                    tk.destroy()

                btn = Button(tk, text="Добавить", command=clicked1)
                btn.grid(column=0, row=8)



            btn = Button(tk, text="Далее",command = clicked)
            btn.grid(column=0, row=5)


    def delete_member():
        notification = Label(tk, text='Какого сотрудника вы хотите удалить из списка?\nВведите ФИО',font=('Arial Bold', 14))
        notification.grid(column=0, row=3)

        tx = Entry(tk, width=75)
        tx.grid(column=0, row=4)

        def delete():
            people = '{}'.format(tx.get())
            # print(people)

            if people in names:
                for cell in sheet['A'][1:]:
                    if cell.value == people:
                        sheet.delete_rows(cell.row)

            else:
                messagebox.showinfo('Внимание!', 'Такого сотрудника не существует.')
            tk.destroy()



        btn = Button(tk, text="Удалить",command=delete)
        btn.grid(column=0, row=5)

    tk = Tk()

    tk.resizable(width=False,height=False)
    tk.title('Polymetall')


    tk.app_widht = 800
    tk.app_height = 500
    tk.screen_width = tk.winfo_screenwidth()
    tk.screen_height = tk.winfo_screenheight()

    tk.x = (tk.screen_width / 2) - (tk.app_widht / 2)
    tk.y = (tk.screen_height / 2) - (tk.app_height / 2)
    tk.geometry(f'{tk.app_widht}x{tk.app_height}+{int(tk.x)}+{int(tk.y)}')

    rad1 = Radiobutton(tk,text='Добавить сотрудника',value=1,command =clicked)
    rad2 = Radiobutton(tk,text='Удалить сотрудника',value=2,command = delete_member)
    #btn = Button(tk, text="Обновить",command=submit)

    #btn.place(x=550,y=135)
    #rad3 = Radiobutton(tk,text='Изменить данные сотрудника',value=3)
    rad1.place(x=550,y=85)
    rad2.place(x=550,y=110)

    #rad3.grid(column=2,row=1)

    wb = openpyxl.load_workbook(filename = 'test.xlsx')
    wb.active = 0
    sheet = wb.active

    lbl = Label(tk, text = 'Что вы хотите сделать?', font = ('Arial Bold',14))

    lbl.place(x=520,y=50)

    txt = scrolledtext.ScrolledText(tk, width=40, height=10,font = ('Arial Bold',14))
    txt.grid(column=0, row=2)
    names =[]
    for i in range(1,sheet.max_row+1):

        txt.insert(INSERT, '{} '.format(sheet['A'+str(i)].value))
        names.append(sheet['A'+str(i)].value)
        # print(names)
        txt.insert(INSERT, '{} \n'.format(sheet['B'+str(i)].value))

    our_image=PhotoImage(file="polymetall.png")
    our_image=our_image.subsample(2,2)
    our_label=Label(tk)
    our_label.image=our_image
    our_label["image"]=our_label.image
    our_label.place(x=485,y=360)
    tk.protocol("WM_DELETE_WINDOW", on_closing)
    tk.mainloop()


    wb.save('test.xlsx')

while running:
    run()



