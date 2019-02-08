

#import library modules
from tkinter import *
from tkinter import filedialog
from collections import Counter
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
root.geometry("550x500")


#Entrada Label
entradaLabel = Label(root, text="Entrada")
entradaLabel.grid(row = 2, column = 0, sticky = NSEW,columnspan = 1)

#Total in Label
quantityLabel = Label(root, text="Total In")
quantityLabel.grid(row = 2, column = 2 , sticky = NSEW, columnspan = 1)

#Total Out Label
quantityOutLabel = Label(root, text="Total Out",width=15)
quantityOutLabel.grid(row = 2, column = 9 , sticky = NSEW, columnspan = 1)

#Instructions to scan Label
entry1Label = Label(root, text="Scan Package:")
entry1Label.grid(row = 0, column = 0 , sticky = NSEW)

#Salida Label
salidaLabel = Label(root, text="Salida",width=15)
salidaLabel.grid(row = 2, column = 7 , sticky = NSEW, columnspan = 1)


#save and send buttons
button1 = Button(root,text = 'Save', fg='green')
button2 = Button(root,text = 'Send', fg='red')
#button1.grid(row=3,column=0)
#button2.grid(row = 3 , column = 1)

#Scanner Entry Field
entry1 = Entry(root)
entry1.grid(row = 0,column = 2,columnspan = 3)
entry1.focus_set()

quitButton = Button(root,text="Quit", command = root.quit )
#quitButton.grid(row=1,column=12)



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
subMenu = Menu(menu)
fileMenu = Menu(menu)
menu.add_cascade(label="File", menu=fileMenu)
menu.add_cascade(label="Scan Mode", menu=subMenu)


#Menu Taskbar Commands

#start file save
def file_save():
    f = filedialog.asksaveasfile(mode='w', defaultextension=".txt",initialdir = "%userprofile%/desktop/",title = "Save file",filetypes = (("Text file","*.txt"),("Excel file","*.xls"),("all files","*.*")))
    if f is None: # asksaveasfile return `None` if dialog closed with "cancel".
        return
    text2saveFurgon = str(furgon.get())
    f.write("Furgon")
    f.write("\n")
    f.write(text2saveFurgon)
    f.write("\n")
    f.write("\n")
    f.write("Total In")
    #end check if list empty

    f.write("\n")
    #check if list is empty
    #if item in storedIn:
    text2saveIn = str(inCounter)
    f.write(text2saveIn)
    f.write("\n")
    f.write("\n")
    f.write("Total Out")
    f.write("\n")
    #if item in storedOut:
    text2saveOut = str(outCounter)
    f.write(text2saveOut)
    f.write("\n")
    #f.write("Scan Details")
    #f.write("\n")
    #f.write(timestamp)
    f.close()
    print(storedIn,storedOut)
#end file save

def entradaMode():
    global mode
    print("Entrada Selected")
    mode = "Entrada"
    select = Label(root, text=" Scan mode:   " + mode, fg='green',width = 17)
    select.grid(row=0, columnspan=2, sticky=NSEW)

def salidaMode():
    print("Salida Selected")
    global mode
    global furgon
    global furgonNumber
    global furgonLabel

    mode = "Salida"
    select = Label(root, text="Scan mode: " + mode, fg='red',width = 17)
    select.grid(row=0, columnspan=2, sticky=NSEW)

    # Furgon Entry Field
    furgonLabel = Label(root, text="Furgon: ")
    furgonLabel.grid(row=0, column= 8, sticky=E,columnspan=1)
    furgon = Entry(root)
    furgonNumber = furgon.get()
    furgon.grid(row=0, column=9, columnspan=1,sticky=E)
    #

class savetoexcel:

    def save2Excel():

        style0 = xlwt.easyxf('font: name Times New Roman, color-index red, bold on',
                         num_format_str='#,##0.00')
        style1 = xlwt.easyxf(num_format_str='D-MMM-YY')

        wb = xlwt.Workbook()
        ws = wb.add_sheet('A Test Sheet', cell_overwrite_ok=True)
        wsLog = wb.add_sheet('Scan Log', cell_overwrite_ok=True)

        ws.write(0, 0, "Total In", style0)
        ws.write(1, 0, str(storedIn) , style1)
        ws.write(2, 0, )
        ws.write(0, 1, "Total Out", style0)
        ws.write(1, 1, str(storedOut))
        ws.write(0, 2, "Furgon", style0)
        ws.write(1, 2, furgonNumber, style1)
        ws.write(2, 2, xlwt.Formula("A3+B3"))
        wsLog.write(0, 0, "Scan Log", style0)
        wsLog.write(1, 0, timestamp, style1)
        print (content)


        wb.save('Z:/Reciving Almacen/Entradas y Salidas/example.xls')

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
        wb.save('Z:/Reciving Almacen/Entradas y Salidas/' + now.strftime("%m-%d-%Y") + '.xls')



subMenu.add_command(label="Entrada", command=entradaMode)
subMenu.add_command(label="Salida", command=salidaMode)

fileMenu.add_command(label="Save As",command = file_save)
#fileMenu.add_command(label="Save to Excel",command = savetoexcel.save2Excel)

fileMenu.add_command(label="Exit", command=root.quit)


mode = ""
content = entry1.get()
timestamp = "Package " + content + " Scanned at " + now.strftime("%m-%d-%Y %H:%M")

subMenu.add_separator()

def func(event):


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



        inLabel.grid(sticky=NSEW, column=0,row=startRowIn)
        global inCounter
        inCounter = Counter(scanIn)
        inCounter = inCounter.most_common()

        startRow = 3
        for value, count in inCounter:
            storedIn = value, "-" ,count
            counterLabel = Label(root, text = storedIn)
            counterLabel.grid(column=2,columnspan = 1,row = startRow, sticky = NSEW)
            startRow = startRow + 1

            print(storedIn)
            savetoexcel.saveScan(startRowSave,timestamp)

        entry1.delete(0, 'end')


    else:
        print( "Package " + content + " Scanned at " + now.strftime("%m-%d-%Y %H:%M"))
        scanOut.append(content)

        startRowOut = 2
        for i in scanOut:
            outLabel = Label(root, text=i,width=15)
            outLabel.grid_forget()
            outCount += 1
            startRowOut = startRowOut + 1

        outLabel.grid(sticky=W, column=7,row=startRowOut)

        entry1.delete(0, 'end')

        global outCounter
        outCounter = (Counter(scanOut))
        outCounter = outCounter.most_common()

        startRow = 3
        for value, count in outCounter:
            storedOut = value, "-", count
            counterOutLabel = Label(root, text=storedOut,width=15)
            counterOutLabel.grid(column=9, columnspan=1, row=startRow, sticky=NSEW)
            startRow = startRow + 1

            if scanIn == scanOut:
                counterOutLabel.config(fg="green")
                furgonLabel.config(fg="green")


        entry1.delete(0, 'end')





    #inCounter = Counter(scanIn)
    #inCounter = inCounter.most_common()
 #   for value, count in inCounter:

  #       print(value, count)


   # totalInLabel = Label(root, text=inCounter)
  #  totalInLabel.grid(row= 3, column = 8)



entradaMode()
root.bind('<Return>', func)



root.mainloop()