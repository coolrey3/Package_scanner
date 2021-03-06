

#import library modules
from tkinter import *
from tkinter import filedialog
from collections import Counter
from tkinter import Tk, Label, Button
from tkinter.filedialog import askopenfilename
import sys
import os
import datetime
import xlrd
import xlwt


#current time & date
now = datetime.datetime.now()
print (now.strftime("%m-%d-%Y"))

#tkinter window
root = Tk()
root.iconbitmap('icon.ico')
#tkinter window title
root.title("Package Scanner")
#tkinter window size
root.geometry("500x600")

#scrollbar code
scrollbar = Scrollbar(root)
scrollbar.grid( sticky = E)#side = RIGHT, fill = Y  )
#Scan In log
list_in = Listbox(root, yscrollcommand = scrollbar.set, height=30)
list_in.grid(sticky=NSEW, column=0, row=3)#side = LEFT, fill = BOTH )
scrollbar.config(command = list_in.yview)

#Scan Out Log
list_out = Listbox(root, yscrollcommand = scrollbar.set, height=30)
list_out.grid(sticky=W, column=7, row=3)#side = LEFT, fill = BOTH )

#Total scanned in
mylistTotalIn = Listbox(root, yscrollcommand = scrollbar.set,height=30 )
mylistTotalIn.grid(column=2,columnspan = 1,row = 3, sticky = NSEW )#side = LEFT, fill = BOTH )

#Total scanned Out
mylistTotalOut = Listbox(root, yscrollcommand = scrollbar.set )
mylistTotalOut.grid(column=8, columnspan=1, row=3, sticky=NSEW )#side = LEFT, fill = BOTH )



#Entrada Label
entradaLabel = Label(root, text="Entrada")
entradaLabel.grid(row = 2, column = 0, sticky = NSEW,columnspan = 1)

#Total in Label
quantityLabel = Label(root, text="Total In")
quantityLabel.grid(row = 2, column = 2 , sticky = NSEW, columnspan = 1)

#Total Out Label
quantityOutLabel = Label(root, text="Total Out",width=15)
quantityOutLabel.grid(row = 2, column = 8 , sticky = NSEW, columnspan = 1)

#Instructions to scan Label
entry1Label = Label(root, text="Scan Package:")
entry1Label.grid(row = 0, column = 0 , sticky = NSEW)

#Salida Label
salidaLabel = Label(root, text="Salida",width=15)
salidaLabel.grid(row = 2, column = 7 , sticky = NSEW, columnspan = 1)

#Scanner Entry Field
entry1 = Entry(root)
entry1.grid(row = 0,column = 2,columnspan = 3)
entry1.focus_set()

# Furgon Entry Field
furgonLabel = Label(root, text="Furgon: ")
furgonLabel.grid(row=0, column=7, sticky=E, columnspan=1)
furgon = Entry(root)
furgonNumber = furgon.get()
furgon.grid(row=0, column=8, columnspan=1, sticky=E)

#Define Variables and Arrays
scanIn = []
scanOut = []
startRowSave = 1
startRow = 3
inCount=0
outCount=0

#Menu Taskbar
menu = Menu(root)
root.config(menu=menu)
subMenu = Menu(menu,tearoff=False)
fileMenu = Menu(menu,tearoff=False)

selection_in = Menu(menu, tearoff=False)
selection_out = Menu(menu, tearoff=False)
menu.add_cascade(label="File", menu=fileMenu)
menu.add_cascade(label="Scan Mode", menu=subMenu)
# menu.add_cascade(label="Delete", menu=selectionIn)
# menu.add_cascade(label="Selection", menu=selectionOut)

currentSelection = Label(root, text = "Selection:  " ,fg = "white",bg = "black",width=20)
currentSelection.grid(row = 10,sticky='W')
#Scan Modes
def entradaMode():
    global mode
    print("Entrada Selected")
    mode = "Entrada"
    select = Label(root, text=" Scan mode:   " + mode, fg='green',width = 17)
    select.grid(row=0, columnspan=2, sticky=NSEW)

def salidaMode():
    global mode
    print("Salida Selected")
    mode = "Salida"
    select = Label(root, text="Scan mode: " + mode, fg='red',width = 17)
    select.grid(row=0, columnspan=2, sticky=NSEW)

class saveas:

    def save2Excel():

        style0 = xlwt.easyxf('font: name Times New Roman, color-index red, bold on',num_format_str='#,##0.00')
        style1 = xlwt.easyxf(num_format_str='D-MMM-YY')
        wb = xlwt.Workbook()
        ws = wb.add_sheet('Package Scanner Totals', cell_overwrite_ok=True)
        #wsLog = wb.add_sheet('Scan Log', cell_overwrite_ok=True)

        ws.write(0, 0, "Total In", style0)
        inCounter = Counter(scanIn)
        inCounter = inCounter.most_common()
        startRow=1
        try:
            for value, count in inCounter:
                storedIn = value,count
                #counterLabel = Label(root, text = storedIn)
                #counterLabel.grid(column=2,columnspan = 1,row = startRow, sticky = NSEW)
                ws.write(startRow,0 , str(storedIn) , style1)
                startRow = startRow + 1

        except:
            pass
        #ws.write(2, 0, )
        ws.write(0, 1, "Total Out", style0)

        startRow=1
        outCounter = (Counter(scanOut))
        outCounter = outCounter.most_common()
        try:
            for value, count in outCounter:
                storedOut = value, count
                # counterLabel = Label(root, text = storedIn)
                # counterLabel.grid(column=2,columnspan = 1,row = startRow, sticky = NSEW)
                ws.write(startRow, 1, str(storedOut))
                startRow = startRow + 1
        except:
            pass

        ws.write(0, 2, "Furgon", style0)
        try:
            text2saveFurgon = str(furgon.get())
            ws.write(1, 2, text2saveFurgon, style1)
        except:
            pass
        #ws.write(2, 2, xlwt.Formula("A3+B3"))
        #wsLog.write(0, 0, "Scan Log", style0)
        #wsLog.write(1, 0, timestamp, style1)
       #print (content)

        filename=  now.strftime("%H-%M_%m-%d-%Y")
        wb.save('Z:/Reciving Almacen/Entradas y Salidas/ '+ filename +'.xls')

    def saveScan(startRowIn, timestamp):


        style0 = xlwt.easyxf('font: name Times New Roman, color-index red, bold on',
                         num_format_str='#,##0.00')
        style1 = xlwt.easyxf(num_format_str='D-MMM-YY')

        wb = xlwt.Workbook()
        ws = wb.add_sheet('A Test Sheet', cell_overwrite_ok=False)
        wsLog = wb.add_sheet('Scan Log', cell_overwrite_ok=False)
        wsLog.write(0, 0, "Scan Log", style0)

        wsLog.write(startRowIn, 0, timestamp, style1)
        print("printed from savescan")


        wb.save('Z:/Reciving Almacen/Entradas y Salidas/Scan Log/' + now.strftime("%m-%d-%Y") + '.xls')

    # Menu Taskbar Commands
    # start file save
    def file_save():
        f = filedialog.asksaveasfile(mode='w', defaultextension=".txt", initialdir="%userprofile%/desktop/",
                                     title="Save file",
                                     filetypes=(("Text file", "*.txt"), ("Excel file", "*.xls"), ("all files", "*.*")))
        if f is None:  # asksaveasfile return `None` if dialog closed with "cancel".
            return
        f.write("Furgon")
        f.write("\n")
        try:
            text2saveFurgon = str(furgon.get())
            f.write(text2saveFurgon)
        except:
            pass
        f.write("\n")
        f.write("\n")
        f.write("Total In")
        # end check if list empty
        f.write("\n")
        # check if list is empty
        #if item in storedIn:

        inCounter = Counter(scanIn)
        inCounter = inCounter.most_common()

        startRow = 3
        for value, count in inCounter:
            storedIn = value,count
            #counterLabel = Label(root, text = storedIn)
            #counterLabel.grid(column=2,columnspan = 1,row = startRow, sticky = NSEW)
            startRow = startRow + 1

            try:
                text2saveIn = str(storedIn) + "\n"
                f.write(text2saveIn)

            except:
                pass
        f.write("\n")
        f.write("\n")
        f.write("Total Out")
        f.write("\n")
        # if item in storedOut:
        try:
            for value, count in outCounter:
                storedOut = value, count
            # counterLabel = Label(root, text = storedIn)
            # counterLabel.grid(column=2,columnspan = 1,row = startRow, sticky = NSEW)
                startRow = startRow + 1


                text2saveOut = str(storedOut) + "\n"
                f.write(text2saveOut)
        except:
            pass
        f.write("\n")
        # f.write("Scan Details")
        # f.write("\n")
        # f.write(timestamp)
        f.close()
        #print(storedIn, storedOut)
    # end file save

def open():
    print("opened file")
    filename = askopenfilename()
    print(filename)

#File Menu List
#fileMenu.add_command(label="Open",command = open)
fileMenu.add_command(label="Save As",command = saveas.file_save)
fileMenu.add_command(label="Save to Excel",command = saveas.save2Excel)
fileMenu.add_command(label="Exit", command=root.quit)

#Scan Menu List
subMenu.add_command(label="Entrada", command=entradaMode)
subMenu.add_command(label="Salida", command=salidaMode)




def restart_program():
    #Restarts the current program.
    #Note: this function does not return. Any cleanup action (like
    #saving data) must be done before calling this function.
    python = sys.executable
    os.execl(python, python, * sys.argv)

#root = Tk()
#fileMenu.add_command(label="Restart",command = restart_program)

mode = ""
content = entry1.get()
timestamp = "Package " + content + " Scanned at " + now.strftime("%m-%d-%Y %H:%M")
#subMenu.add_separator()

def func(event):
    global mode
    global inCount
    global outCount
    global storedIn
    global storedOut
    content = entry1.get()

    if content == "":
        print("No Package Scanned")
        print(mode)

    elif mode == "Entrada":
        timestamp = "Package " + content + " Scanned at " + now.strftime("%m-%d-%Y %H:%M")
        print( timestamp)
        scanIn.append(content)
        startRowIn = 2

        for i in scanIn:
            inLabel = Label(root, text=i)
            inCount += 1
            inLabel.grid_forget()
            startRowIn = startRowIn + 1
        list_in.insert(0, i)

       #inLabel.grid(sticky=NSEW, column=0,row=3)#startRowIn
        global inCounter
        inCounter = Counter(scanIn)
        inCounter = inCounter.most_common()
        mylistTotalIn.delete(0, END)

        startRow = 3
        for value, count in inCounter:
            storedIn = value, "-" ,count
            counterLabel = Label(root, text = storedIn)
           # counterLabel.grid(column=2,columnspan = 1,row = startRow, sticky = NSEW)
            mylistTotalIn.insert(END, storedIn)
            startRow = startRow + 1
            print(storedIn)
            #saveas.saveScan(startRowSave,timestamp)

        entry1.delete(0, 'end')

    else:
        print( "Package " + content + " Scanned at " + now.strftime("%m-%d-%Y %H:%M"))
        scanOut.append(content)
        mode = "Salida"

        startRowOut = 2
        for i in scanOut:
            outLabel = Label(root, text=i,width=15)
            outLabel.grid_forget()
            outCount += 1
            startRowOut = startRowOut + 1

        #outLabel.grid(sticky=W, column=7,row=3)#startRowOut
        list_out.insert(0, i)
        entry1.delete(0, 'end')

        global outCounter
        outCounter = (Counter(scanOut))
        outCounter = outCounter.most_common()
        mylistTotalOut.delete(0, END)
        startRow = 3

        for value, count in outCounter:
            storedOut = value, "-", count
            #counterOutLabel = Label(root, text=storedOut,width=15)
            #counterOutLabel.grid(column=9, columnspan=1, row=startRow, sticky=NSEW)
            startRow = startRow + 1
            mylistTotalOut.insert(END, storedOut)

            if scanIn == scanOut:
                counterOutLabel.config(fg="green")
                furgonLabel.config(fg="green")

        entry1.delete(0, 'end')


def deleteIn():
    try:

        print('printed from delete')
        #
        # widget = event.widget
        # selection=widget.curselection()
        # value = widget.get(selection[0])
        print(scanIn)
        if value in scanIn:
            print('value found :' + value)
            print("Now Removing")
            scanIn.remove(value)
            # mylistIn.remove(value)
            print(list_in.curselection())
            list_in.delete(list_in.curselection())




            # if event.widget.get(event.widget.curselection()[0]) == value:


            # inCount.update
            # outCount.update

            print(storedOut)
            print(list_in)
            # scanIn.sort




        else:
            print('value not found')
    except:
        pass
    # scanIn.sort


try:
    # Selection menu
    # selectionIn.add_command(label="Delete In", command=deleteIn)
    selection_out.add_command(label="Delete Out", command=deleteOut)


except:
    pass

def deleteOut():
    try:

        print('printed from delete Out')
        #
        # widget = event.widget
        # selection=widget.curselection()
        # value = widget.get(selection[0])
        print(scanIn)
        if value in scanOut:
            print('value found :' + value)
            print("Now Removing")
            scanOut.remove(value)
            # mylistIn.remove(value)
            print(list_out.curselection())
            list_out.delete(list_out.curselection())
            # mylistTotalOut.insert(0, " ")
            # mylistOut.delete(0, 0)
            # mylistTotalOut.insert(0, "filler")
            # mylistTotalOut.delete(0,0)




            # if event.widget.get(event.widget.curselection()[0]) == value:


            # inCount.update
            # outCount.update
            print('this is before')


            print(storedOut)
            print(list_out)
            # scanIn.sort


            func('<RETURN>')

            newtemp = mylistTotalOut

            # mylistTotalOut.delete(0, END)

            # for value, count in outCounter:
            #     storedOut = value, "-", count
            #     # counterOutLabel = Label(root, text=storedOut,width=15)
            #     # counterOutLabel.grid(column=9, columnspan=1, row=startRow, sticky=NSEW)
            #     startRow = startRow + 1
            #     mylistTotalOut.insert(END, storedOut)

            # mylistTotalOut.insert(0, 'test')

            print('this is it')







        else:
            print('value not found')

        # func('<RETURN>')

    except:
        pass
    # scanIn.sort


try:
    # Selection menu
    selection_in.add_command(label="Delete In", command=deleteIn)
except:
    pass

try:
    # Selection menu
    selection_out.add_command(label="Delete Out", command=deleteOut)
except:
    pass



def switchMode(event):
    print("Switching scan mode from " + mode)
    if mode == "Salida":
        entradaMode()
    else:
        salidaMode()


def onDouble(event):
    global value
    widget = event.widget
    selection=widget.curselection()
    value = widget.get(selection[0])
    print (value)


    currentSelection = Label(root, text="Selection: " + value, fg = "white",bg = "black",width=19)
    currentSelection.grid_forget()
    currentSelection.grid(row=10, sticky = 'W')




def do_popupIn(event):
    # display the popup menu
    try:
        selection_in.tk_popup(event.x_root, event.y_root, 0)
    finally:
        # make sure to release the grab (Tk 8.0a1 only)
        selection_in.grab_release()

def do_popupOut(event):
    # display the popup menu
    try:
        selection_out.tk_popup(event.x_root, event.y_root, 0)
    finally:
        # make sure to release the grab (Tk 8.0a1 only)
        selection_out.grab_release()


try:
    list_in.bind("<Double-Button-1>", onDouble)
    list_in.bind("<Button-3>", do_popupIn)

except:
    pass

# try:
#     mylistTotalIn.bind("<Double-Button-1>", onDouble)  #<<ListboxSelect>>
# except:
#     pass

try:
    list_out.bind("<Double-Button-1>", onDouble)

    list_out.bind("<Button-3>", do_popupOut)
except:
    pass
#
# try:
#     mylistTotalOut.bind("<Double-Button-1>", onDouble )
#
#     mylistTotalOut.bind("<Button-3>", do_popup)
# except:
#     pass



root.bind('<Return>', func)
root.bind('<Control_L>',switchMode)
entradaMode()
root.mainloop()

#save and send buttons
#button1 = Button(root,text = 'Save', fg='green')
#button2 = Button(root,text = 'Send', fg='red')
#button1.grid(row=3,column=0)
#button2.grid(row = 3 , column = 1)