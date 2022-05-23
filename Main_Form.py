import tkinter
import Main
import string
import AutoDO
from tkinter import Frame,BOTH,Tk,Button,Text,WORD,Label,END,Checkbutton,IntVar
class W_NIOSS_RRL(Frame):
    def __init__(self, parent):
        Frame.__init__(self, parent, background="white")
        self.parent = parent
        self.parent.title("MTS Transport")
        self.pack(fill=BOTH, expand=1)
        self.centerWindow()
        self.variable = tkinter.StringVar()
        self.StatusBar = Label(self, bd=1, relief=tkinter.SUNKEN,
                              textvariable=self.variable,
                              font=('arial', 10, 'normal'))
        self.variable.set('Скорпируйте ссылку из НИОСС на РРЛ двойным щелчком')
        self.StatusBar.grid(row=8,columnspan=2,sticky='w')
        self.isAGOS= IntVar()
        self.isCheckList = IntVar()
        self.isCheckAGOS = IntVar()
        self.go_button=Button(self,text='Запустить')
        self.go_button.grid(row=7,column=0,sticky='w')
        self.label_RRL_ID=Label(self,text='RRL_ID',background="white")
        self.label_RRL_ID.grid(row=2,column=0,sticky='e')
        self.text_RRL_ID = Text(self, width=40,height=1,wrap=WORD,font='Arial 10')
        self.text_RRL_ID.bind("<Double-1>", self.get_clipboard_ID)
        self.text_RRL_ID.insert(1.0,'Вставьте ссылку на RRL или obj_ID')
        self.text_RRL_ID.grid(row=2,column=1,padx=3, pady=3)
        self.label_RRL_ip = Label(self, text='ip', background="white")
        self.label_RRL_ip.grid(row=3, column=0, sticky='e')
        self.text_RRL_ip = Text(self, width=40, height=1, wrap=WORD, font='Arial 10')
        self.text_RRL_ip.bind("<Double-1>", self.get_clipboard_ip)
        self.text_RRL_ip.insert(1.0, 'Вставьте ip')
        self.text_RRL_ip.grid(row=3, column=1,padx=3, pady=3)
        self.check_bt_AGOS=Checkbutton(self,text="Завести АГОС к РРС",variable=self.isAGOS,background="white",onvalue=1, offvalue=0)
        self.check_bt_AGOS.select()
        self.check_bt_AGOS.grid(row=4, column=0,columnspan=2,padx=3,sticky='w')
        self.check_bt_CheckList = Checkbutton(self,text="Завести работу на проверку чеклиста", variable=self.isCheckList,background="white",onvalue=1, offvalue=0)
        self.check_bt_CheckList.select()
        self.check_bt_CheckList.grid(row=5, column=0,padx=3,columnspan=2,sticky='w')
        self.check_bt_Check_AGOS = Checkbutton(self,text="Завести работу на проверку АГОС", variable=self.isCheckAGOS,background="white",onvalue=1, offvalue=0)
        self.check_bt_Check_AGOS.select()
        self.check_bt_Check_AGOS.grid(row=6, column=0,padx=3,columnspan=2,sticky='w')




    def get_ID_RRL(self):
        captured = str(self.text_RRL_ID.get("1.0", tkinter.END))
        return captured

    def get_ip_RRL(self):
        captured = str(self.text_RRL_ip.get("1.0", tkinter.END))
        return captured

    def get_clipboard_ID(self,event):
        str1=self.clipboard_get()
        if self.text_RRL_ID.get(1.0, END)!='':
            self.text_RRL_ID.delete(1.0, END)
            self.text_RRL_ID.insert(1.0,str1)

    def get_clipboard_ip(self,event):
        str1=self.clipboard_get()
        if self.text_RRL_ip.get(1.0, END)!='':
            self.text_RRL_ip.delete(1.0, END)
            self.text_RRL_ip.insert(1.0,str1)

    def centerWindow(self):
        # w = 390
        # h = 250
        sw = self.parent.winfo_screenwidth()
        sh = self.parent.winfo_screenheight()
        # x = (sw - w) / 2
        # y = (sh - h) / 2
        # self.parent.geometry('%dx%d+%d+%d' % (w, h, x, y))


class W_NIOSS_new_RRL(Frame):
    def __init__(self, parent):
        Frame.__init__(self, parent, background="white")
        self.parent = parent
        self.parent.title("Создать РРЛ в НИОСС")

        self.pack(fill=BOTH, expand=1)
        self.variable = tkinter.StringVar()
        self.StatusBar = Label(self, bd=1, relief=tkinter.SUNKEN,
                              textvariable=self.variable,
                              font=('arial', 10, 'normal'))
        self.variable.set('Скорпируйте ссылку из НИОСС на РРЛ двойным щелчком')
        self.StatusBar.grid(row=8,columnspan=2,sticky='w')

        self.go_button = Button(self, text='Запустить')
        self.go_button.grid(row=7, column=0, sticky='w')

        self.isRRN_A = IntVar()
        self.isRRN_Z = IntVar()
        self.isRRN_A.set(1)
        self.isRRN_Z.set(1)
        self.check_RRN_A = Checkbutton(self, text="Создать РРУ А",
                                              variable=self.isRRN_A, background="white",onvalue=1, offvalue=0)
        # self.check_RRN_A.select()
        self.check_RRN_A.grid(row=3, column=0, padx=3, columnspan=2, sticky='w')

        self.check_RRN_Z = Checkbutton(self, text="Создать РРУ Z",
                                              variable=self.isRRN_Z, background="white",onvalue=1, offvalue=0)
        # self.check_RRN_Z.select()
        self.check_RRN_Z.grid(row=4, column=0, padx=3, columnspan=2, sticky='w')

        self.label_PL_A=Label(self,text='PL_A_',background="white")
        self.label_PL_A.grid(row=1,column=0,sticky='e')
        self.text_PL_A = Text(self, width=40,height=1,wrap=WORD,font='Arial 10')
        self.text_PL_A.insert(1.0,'23_')
        self.text_PL_A.grid(row=1,column=1,padx=3, pady=3)

        self.label_PL_Z = Label(self, text='PL_Z_', background="white")
        self.label_PL_Z.grid(row=2, column=0, sticky='e')
        self.text_PL_Z = Text(self, width=40, height=1, wrap=WORD, font='Arial 10')
        self.text_PL_Z.insert(1.0, '23_')
        self.text_PL_Z.grid(row=2, column=1, padx=3, pady=3)
    def go_button_click(self,event):
        import string


        AutoDO.RRL_birth('PL_'+self.text_PL_A.get(1.0, END)[:-1],'PL_'+self.text_PL_Z.get(1.0, END)[:-1],self.isRRN_A.get(),self.isRRN_Z.get())








class tool_bar(Frame):
    def __init__(self, parent):
        Frame.__init__(self, parent, background="white")
        self.parent = parent
        self.parent.title("Панель инструментов")
        self.pack(fill=BOTH, expand=1)
        self.variable = tkinter.StringVar()
        self.StatusBar = Label(self, bd=1, relief=tkinter.SUNKEN,
                              textvariable=self.variable,
                              font=('arial', 10, 'normal'))
        self.variable.set('Нажмите нужную кнопку')
        self.StatusBar.grid(row=8,columnspan=2,sticky='w')

        self.btRRL_to_NIOSS=Button(self,text='Заполнить РРЛ НИОСС')
        self.btRRL_to_NIOSS.grid(row=2,column=0,sticky='w')
        self.bt_birth_RRL_NIOSS=Button(self,text='Создать РРЛ НИОСС')
        self.bt_birth_RRL_NIOSS.grid(row=1,column=0,sticky='w')

    def btRRL_to_NIOSS_click(self,event):
        root = Tk()
        ex = W_NIOSS_RRL(root)
        ex.go_button.bind("<Button-1>", Main.RRL_NIOSS)
        root.mainloop()

    def bt_birth_RRL_NIOSS_click(self, event):
        root = Tk()
        ex = W_NIOSS_new_RRL(root)
        ex.go_button.bind("<Button-1>", ex.go_button_click)
        root.mainloop()



if __name__ == "__main__":

    root = Tk()
    ex = W_NIOSS_new_RRL(root)
    ex.go_button.bind("<Button-1>", ex.go_button_click)
    # ex.bt_birth_RRL_NIOSS.bind("<Button-1>", ex.bt_birth_RRL_NIOSS_click)
    root.mainloop()

    # root = Tk()
    # ex = tool_bar(root)
    # ex.btRRL_to_NIOSS.bind("<Button-1>", ex.btRRL_to_NIOSS_click)
    # ex.bt_birth_RRL_NIOSS.bind("<Button-1>", ex.bt_birth_RRL_NIOSS_click)
    # root.mainloop()
    # RRL_NIOSS()

