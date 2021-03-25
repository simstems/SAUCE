import sys
import time
import subprocess
from tkinter import Tk as tink
from tkinter import ttk as tinky
from tkinter import Frame, Menu, Entry, IntVar, Checkbutton, Button, Label, StringVar, END


# Here, we are creating our class, Window, and inheriting from the Frame
# class. Frame is a class from the tkinter module. (see Lib/tkinter/__init__)
class Window(Frame):

    # Define settings upon initialization. Here you can specify
    def __init__(self, master=None):
        
        # parameters that you want to send through the Frame class. 
        Frame.__init__(self, master)   

        #reference to the master widget, which is the tk window                 
        self.master = master

        #with that, we want to then run init_window, which doesn't yet exist
        self.init_window()

    #Creation of init_window
    def init_window(self):

        # changing the title of our master widget      
        self.master.title('SAUCE')

        # allowing the widget to take the full space of the root window
        self.grid()

        # creating a menu instance
        menu = Menu(self.master)
        self.master.config(menu=menu)

        # create the file object)
        file = Menu(menu)

        # adds a command to the menu option, calling it exit, and the
        # command it runs on event is client_exit
        file.add_command(label="Exit", command=self.client_exit)

        #added "file" to our menu
        menu.add_cascade(label="File", menu=file)

        # create the file object)
        edit = Menu(menu)

        # adds a command to the menu option, calling it exit, and the
        # command it runs on event is client_exit
        edit.add_command(label="Undo")

        #added "file" to our menu
        menu.add_cascade(label="Edit", menu=edit)

    
    def client_exit(self):
        exit()

class PowerShell:

    def _init_(self):
        return True

    #Functions
    def RunCommand(self, Command):
        process = subprocess.Popen(["powershell", Command], stdout=subprocess.PIPE, stderr=subprocess.PIPE, universal_newlines=True)
        output = process.stdout.readlines()
        return output
             
    def get_computerName(self):
        computerName = self.RunCommand("$env:computername")
        computerName = computerName[0].strip('\n')
        return computerName

    def get_Serial(self):
        return "wmic bios get serialnumber"
    
    def get_Model(self):
        return "wmic csproduct get name"
    
    def get_Brand(self):
        return "wmic csproduct get vendor"

    def get_Status(self, computerName):
        isOnline = self.RunCommand("Test-Connection -BufferSize 32 -Count 1 -ComputerName '" + computerName + "' -Quiet")
        isOnline = isOnline[0].strip('\n') #[i.strip('\n') for i in isOnline]
        if isOnline == "True":
            return "Online"
        else:
            return "Offline"

    def get_MAC(self):
        MAC = ""
        ethernet = 'Get-NetAdapter -Name "Ethernet" | Select-Object '
        wi_fi = 'Get-NetAdapter -Name "Wi-Fi" | Select-Object '
        e_status = ""
        w_status = ""

        try:
            w_status = self.RunCommand(str(wi_fi) + "Status")
            if(w_status.lower() == "online"):
                w_status = True
            else:
                w_status = False
        except:
            print("No MSFT_NetAdapter objects found with property 'Name' equal to 'Wi-Fi'")

        if (w_status == True):
            MAC = (wi_fi + "MacAddress")
        else:
            MAC = (ethernet + "MacAddress")     
        return MAC

    def get_computerInfo(self, computerName):
        result=[]
        commands = [self.get_Serial(), self.get_Model(), self.get_Brand(), self.get_MAC()]
        if computerName == "this":
            for command in commands:
                result.append(self.RunCommand(command))
        else:
            for command in commands:
                result.append(self.RunCommand("Invoke-command -ComputerName '" + str(computerName) + "' -ScriptBlock {" + command + "}")) 
        return result 

# MAIN, variables 
shell = PowerShell()
i=2
rows = []

# Data Grid functions here

# Create one row and add in to the rows list
def add_row(name,serial, model, brand, status, mac):
    computer = [name,serial,model,brand,status,mac]
    global i 
    i=i+1
    items = []
    var = IntVar()
    # Create and append the checkbox to the item list
    c = Checkbutton(center, variable = var)
    c.val = var
    items.append(c)
    c.grid(row = i, column = 0)
    # Next create and append the entry boxes to the items list
    for j in range(1,7): #Columns
        e = Entry(center)
        items.append(e)
        e.grid(row=i, column=j)

        data = items[j] # Locate entry
        data.delete(0, END) # Clear any text in entry
        data.insert(0, computer[j-1]) # Add a text value to entry
        data.config(state='readonly',justify='center') # set to read only

    # Append the new entry fields to the rows list
    rows.append(items)

# Remove selected row 
def delete_row():
    for rowno, row in reversed(list(enumerate(rows))):
        if row[0].val.get() == 1:
            for i in row:
                i.destroy()
            rows.pop(rowno)

# This is a function which increases the progress bar value by the given increment amount
def startProgress():
	progessBarOne['value']=progessBarOne['value'] + 1
	window.update_idletasks()

# This is a function which stops the progress bar
def stopProgress():
	progessBarOne.stop()

# This is a function which shows a popup message.

def popupmsg(msg):
    popup = tink()
    popup.geometry('{}x{}'.format(155, 75)) # Width x Height
    popup.wm_title("!")
    label = tinky.Label(popup,text=msg)
    label.configure(anchor="center")
    label.pack(side="top",fill="x",pady=10)
    B1 = tinky.Button(popup,text="Ok",command=popup.destroy)
    B1.pack()
    popup.mainloop()

# This is the function that gets the computer information
def show_ComputerInfo():
    startProgress()
    name = cnEntry.get().lower()
    if name == "this":
        results = shell.get_computerInfo(name)
        add_row(shell.get_computerName(), str(results[0][2]), str(results[1][2]), str(results[2][2]), "Online", str(results[3][3]))
       
    elif shell.get_Status(name) == "Online":
        results = shell.get_computerInfo(name)
        add_row(name.upper(), str(results[0][2]), str(results[1][2]), str(results[2][2]), "Online", "GHT:456")
    else:
        popupmsg("This device may be offline.")
        
    stopProgress()

window = tink()
# window.iconphoto(True, PhotoImage(file="SAS.png"))
window.geometry('{}x{}'.format(873, 350)) # Width x Height
window.resizable(False, False)
window.config(bg='#fff')

# create all of the main containers

top_frame = Frame(window, bg='#fff', width=450, height=80, pady=3)
center = Frame(window, bg='#fff', width=50, height=40, padx=3, pady=3)
btm_frame = Frame(window, bg='#fff', width=450, height=20, pady=3)
btm_frame2 = Frame(window, bg='#fff', width=450, height=30, pady=3)

# layout all of the main containers
window.grid_rowconfigure(1, weight=1)
window.grid_columnconfigure(0, weight=1)


top_frame.grid(row=0, sticky="ew", pady=10)
center.grid(row=1, sticky="nsew")
btm_frame.grid(row=3, sticky="ew")
btm_frame2.grid(row=4, sticky="ew", pady=1)

# create the widgets for the top frame
cName_label = Label(top_frame, background='#fff',text='Computer Name:')
cnEntry = Entry(top_frame, width=30, background="#fff")
cnButton = Button(top_frame, width=10, text='FIND', bg='#4EC5F1', font=('arial', 8, 'normal'), command=show_ComputerInfo)

# layout the widgets in the top frame
cName_label.grid(row=1, column=0)
cnEntry.grid(row=1, column=1)
cnButton.grid(row=1, column=2)

# create the center widgets, datagrid
v0 = StringVar()
e0 = Entry(center, textvariable = v0, state = 'readonly', justify = 'center')
v0.set('Select')
e0.grid(row = 1, column = 0 )

v1 = StringVar()
e1 = Entry(center, textvariable = v1, state = 'readonly', justify= 'center')
v1.set('Name')
e1.grid(row = 1, column = 1 )

v2 = StringVar()
e2 = Entry(center, textvariable = v2, state = 'readonly', justify= 'center')
v2.set('Serial#')
e2.grid(row = 1, column = 2 )

v3 = StringVar()
e3 = Entry(center, textvariable = v3, state = 'readonly', justify= 'center')
v3.set('Model#')
e3.grid(row = 1, column = 3)

v4 = StringVar()
e4 = Entry(center, textvariable = v4, state = 'readonly', justify= 'center')
v4.set('Brand')
e4.grid(row = 1, column = 4 )

v5 = StringVar()
e5 = Entry(center, textvariable = v5, state = 'readonly', justify= 'center')
v5.set('Status')
e5.grid(row = 1, column = 5 )

v6 = StringVar()
e6 = Entry(center, textvariable = v6, state = 'readonly', justify= 'center')
v6.set('MAC')
e6.grid(row = 1, column = 6 )

# Vertical (y) Scroll Bar
#scrolly=Scrollbar(window, command=center.yview)
#scrolly.pack(side=LEFT, fill=Y, pady=65)

# create the widgets for the bottom frames
dlButton = Button(btm_frame, text='Delete Row', bg='#4EC5F1', font=('arial', 8, 'normal'), command=delete_row)
progessBarOne=tinky.Progressbar(btm_frame2, style='progessBarOne.Horizontal.TProgressbar', orient='horizontal', length=873, mode='indeterminate', maximum=100, value=0)

# layout the widgets in the bottom frames
dlButton.grid(row =0, column=0)
progessBarOne.grid(row=1, column=0)

#creation of an instance for window menu
app = Window(window)

#creation of an instance for data grid
# data = DataGrid(center)

# END, Open Window
window.mainloop()