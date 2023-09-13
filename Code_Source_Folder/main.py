from tkinter import *
from tkinter import messagebox
import sqlite3
import pandas as pd

def readExcel(fileName):
    try:
        df = pd.read_excel(fileName)
    except:
        return None
    return df

class Database:
    def __init__(self, fileName):
        self.fileName = fileName
        try:
            self.df = readExcel(self.fileName)
            if(self.df is None):
                raise Exception
        except:
            print("File name entered doesn't exist, please enter valid name.")
            return
        try:
            self.conn = sqlite3.connect("Database.db")
        except:
            print("Connection to database unsuccessfull.")
        self.tableName = "student"
        self.columnNames = [columnName for columnName in self.df]
        return
    def generateCreateTableQuery(self):
        self.createTableQuery = "Create table "
        self.createTableQuery += (self.tableName + "(")
        count = 0
        for column in self.df:
            if(count==0):
                self.createTableQuery += (column+" text primary key not null unique, ")
            elif(count==(len(self.columnNames)-1)):
                self.createTableQuery += (column+" text)")
            else:
                self.createTableQuery += (column+" text, ")
            count += 1
        return self.createTableQuery
    def generateInsertTableQuery(self):
        self.insertTableQuery = "insert into "
        self.insertTableQuery += (self.tableName + "(")
        count = 0
        questionString = ""
        for column in self.columnNames:
            if(count==(len(self.columnNames)-1)):
                self.insertTableQuery += (column + ") values(")
                questionString += ("?)")
            else:
                self.insertTableQuery += (column + ", ")
                questionString += "?, "
            count += 1
        self.insertTableQuery += questionString
        return self.insertTableQuery
    def createDatabase(self):
        self.createTableQuery = self.generateCreateTableQuery()
        try:
            self.conn.execute(self.createTableQuery)
        except:
            print("Create table query execution error.")
        print("Table created succesfully.")
        self.insertTableQuery = self.generateInsertTableQuery()
        for index in self.df.index:
            try:
                self.conn.execute(self.insertTableQuery, [str(rowValue) for rowValue in self.df.loc[index]])
            except:
                print("Insert table query execution error.")
        print("Values inserted successfully into database from Excel.")
        try:
            self.conn.commit()
            self.conn.close()
        except:
            print("Commit/close execution error.")
        return
    def updateDatabase(self):
        id = self.df.columns[0]
        idColumn = self.df
        try:
            self.rows = self.conn.execute("Select * from " + self.tableName)
        except:
            print("Query execution error.")
        for row in self.rows:
            flag0 = 0
            for index in self.df.index:
                if(row[0]==str(self.df.loc[index][0])):
                    idColumn = idColumn.drop(idColumn.loc[idColumn['id']==(int(row[0]))].index)
                    flag0 = 1
                    count = 0
                    for i in self.df.loc[index]:
                        if(str(i)!=row[count]):
                            try:
                                updateTableQuery = "Update " + self.tableName + " set " + self.columnNames[count] + " = '" + str(i) + "' where " + id + " = " + row[0]
                                self.conn.execute(updateTableQuery)
                            except:
                                print("Update table query execution error.")
                        count += 1
            if(flag0==0):
                try:
                    deleteTableQuery = "Delete from " + self.tableName + " where " + id + " = " + row[0]
                    self.conn.execute(deleteTableQuery)
                except:
                    print("Delete table query execution error.")
        for index in idColumn.index:
            try:
                self.insertTableQuery = self.generateInsertTableQuery()
                self.conn.execute(self.insertTableQuery, [str(rowValue) for rowValue in idColumn.loc[index]])
            except:
                continue
        try:
            self.conn.commit()
            self.conn.close()
        except:
            print("Commit/close execution error.")
        return
    def sqlToExcel(self):
        self.rows = self.conn.execute("Select * from " + self.tableName)
        list0 = []
        list0.append(self.columnNames)
        for row in self.rows:
            list0.append(list(row))
        df = pd.DataFrame(list0[1:], columns = list0[0])
        df.to_excel(self.fileName, index = False)
        return
        
def driverCode(fileName, option):
    fileName = fileName.strip()
    db = Database(fileName)
    if(option==1):
        try:
            if(db.df is None):
                raise Exception
            db.createDatabase()
        except:
            print("Table cannot be created.")
    elif(option==2):
        try:
            if(db.df is None):
                raise Exception
            db.updateDatabase()
        except:
            print("Table cannot be updated.")
    elif(option==3):
        try:
            if(db.df is None):
                raise Exception
            db.sqlToExcel()
        except:
            print("Database cannot be converted to Excel.")

def SecondWindow():
    d2x = Tk()
    d2x.geometry("1600x300")
    d2x.title("D2X Application")
    title = Label(d2x, text = "Welcome to D2X", fg = "#000000", font = ("Arial", 24))
    subtitle = Label(d2x, text = "Please select the operation:", fg = "#808080", font = ("Arial", 24))
    file_label = Label(d2x, text = "Enter Excel file name:", bg = '#333333', fg = "#FFFFFF", font = ("Arial", 16))
    file_entry = Text(d2x, font = ("Arial", 16))
    file_entry.configure(height=1)
    sqltoexcel = Button(d2x, text = "Sync to Excel file from database", bg = "#CC8899", fg = "#FFFFFF", font = ("Arial", 24), command = lambda: driverCode(file_entry.get('1.0', END), 3))
    exceltosql0 = Button(d2x, text = "Create Database from Excel file", bg = "#CC8899", fg = "#FFFFFF", font = ("Arial", 24), command = lambda: driverCode(file_entry.get('1.0', END), 1))
    exceltosql1 = Button(d2x, text = "Sync to Database from Excel file", bg = "#CC8899", fg = "#FFFFFF", font = ("Arial", 24), command = lambda: driverCode(file_entry.get('1.0', END), 2))
    warning = Label(d2x, text = "NOTE: Any operations cannot be undone!", bg = '#333333', fg = "#808080", font = ("Arial", 14))
    title.grid(row = 0, column = 1)
    subtitle.grid(row = 1, column = 1)
    file_label.grid(row = 3, column = 0)
    file_entry.grid(row = 3, column = 1, padx = 25, columnspan = 3)
    sqltoexcel.grid(row = 5, column = 2, rowspan = 2,  padx = 20, pady = 25)
    exceltosql0.grid(row = 5, column = 0, rowspan = 2,  padx = 20, pady = 25)
    exceltosql1.grid(row = 5, column = 1, rowspan = 2,  padx = 20, pady = 25)
    warning.grid(row = 7, column = 1)

def validateLogin():
    if(username_entry.get()=="admin"and password_entry.get()=="admin"):
       SecondWindow()
       loginWindow.withdraw()
    else:
      messagebox.showerror("Login error!", "Please enter the correct credentials!")
    
loginWindow = Tk()
loginWindow.geometry('440x440')
loginWindow.title('D2X application')
frame = Frame(bg = '#333333')
login_label = Label(frame, text = "LOGIN", bg = '#333333', fg = "#FFFFFF", font = ("Arial", 30))
username_label = Label(frame, text = "Username", bg='#333333', fg="#FFFFFF", font = ("Arial", 16))
username_entry = Entry(frame, font = ("Arial", 16))
password_entry = Entry(frame, show = "*", font = ("Arial", 16))
password_label = Label(frame, text = "Password", bg = '#333333', fg = "#FFFFFF", font = ("Arial", 16))
login_button = Button(frame, text = "LOGIN", bg = "#CC8899", fg = "#FFFFFF", font = ("Arial", 24), command = validateLogin)
login_label.grid(row = 0, column = 0, columnspan = 2, pady = 40)
username_label.grid(row = 1, column = 0)
username_entry.grid(row = 1, column = 4, pady = 20)
password_label.grid(row = 2, column = 0)
password_entry.grid(row = 2, column = 4, pady = 20)
login_button.grid(row = 4, column = 4, rowspan = 4, columnspan = 2, pady = 50)
frame.pack()
loginWindow.mainloop()