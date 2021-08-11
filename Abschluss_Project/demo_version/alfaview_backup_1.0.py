"""PyTym is a software to create and manage Timesheet in a company.
It deals with database, plotting and creating personalized application.
Author : Priya Subramanian Version: 1.0 """
# --- import modules ---
import csv
import datetime
import tkinter as tk
from datetime import *
from tkinter import ttk, messagebox,filedialog

# --- matplotlib ---
import matplotlib
from PIL import ImageTk
# --- tkinter ---
from tkcalendar import DateEntry

matplotlib.use("TkAgg")  # choose backend
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
from matplotlib.pyplot import Figure

# --- Part 1 ---

class Login:
    """The Log-in Class covers log.in, regissteration, log-out, verification and
    inserting the data in database"""
    def __init__(self, root):
        self.root = root

        global employee_no_verify
        global password_verify
        global first_name
        global last_name
        global emp_id
        global login_time
        global login_date

        employee_no_verify = tk.StringVar()
        password_verify = tk.StringVar()

        self.company_name = "Demo Firma"                                # set company name as constant variable
        self.login_tuple = ()
        self.loggedin_now = ""

    def verify_login(self):
        """Verifies the Log-in data and imports the import"""
        emp_no_verify_data = employee_no_verify.get()
        pwd_verify_data = password_verify.get()

        if pwd_verify_data == "" or emp_no_verify_data == "":
            messagebox.showerror("Error", "All fields are required !", parent=self.root)
        else:
            with open('D:\Python\Abschluss_Project\demo_version\Employee_Registeration.csv',
                      newline="") as employee_file:                      # newline = "\ n" switches off the traditional line break.
                emp_data_writer = csv.reader(employee_file, delimiter=',')
                titlecolumn = next(emp_data_writer)                 # this moves the cursor to headerrow and rules out the title column !!!
                for row in emp_data_writer:
                    if emp_no_verify_data == row[0]:
                        verify_pwd = row[1]
                        if pwd_verify_data in verify_pwd:
                            today = date.today()  # set  the values to database
                            now = datetime.now()
                            self.loggedin_now = str(now)
                            self.login_date = today.strftime("%b-%d-%Y")
                            self.login_time = now.strftime("%H:%M:%S")
                            self.emp_id = (row[0])
                            self.first_name = row[3]
                            self.last_name = row[4]
                            self.login_tuple = (self.emp_id, self.login_date, self.login_time)

                            with open('D:\Python\Abschluss_Project\demo_version\Timesheet\Jun-2021\Employee_Timesheet_Tracker.csv', 'a', newline="") as employee_file:  # newline = "\ n" switches off the traditional line break.
                                emp_data_writer = csv.writer(employee_file, delimiter=',', quotechar='"',
                                                             quoting=csv.QUOTE_NONE, escapechar="'")
                                emp_data_writer.writerow([self.emp_id, self.login_date, self.login_time])

                            messagebox.showinfo(title="Log-in Success",
                                                message=" Welcome,  " + self.first_name + "  " + self.last_name + "  you have successfully logged in on  " + self.login_date + "   at   " + self.login_time)

                            root.destroy()
                            break
                        else:
                            messagebox.askretrycancel(title="Verify Password", message="Incorrect Password !")
                        break
                # else:
                # messagebox.askretrycancel(title="Verify Employee ID", message="Incorrect Employee ID !")

    def new_user_register(self):                                # Create and register New Employee details
        global set_emp_no                                        # set global variables
        global emp_pwd_entry
        global emp_dob_entry
        global emp_doj_entry
        global emp_dept_entry
        global emp_email_entry
        global newWindow
        global emp_firstname_entry
        global emp_lastname_entry

        with open('D:\Python\Abschluss_Project\demo_version\Employee_Registeration.csv', 'r',newline="") as employee_file:                                                      # newline = "\ n" switches off the traditional line break.
            set_emp_no = sum(1 for line in employee_file) + 99

        emp_pwd_entry = tk.StringVar()
        emp_dob_entry = tk.StringVar()
        emp_doj_entry = tk.StringVar()
        emp_dept_entry = tk.StringVar()
        emp_email_entry = tk.StringVar()
        emp_lastname_entry = tk.StringVar()
        emp_firstname_entry = tk.StringVar()

        newWindow = tk.Toplevel(self.root)                                                    # assign a new window
        newWindow.title("New Employee Registration")                                     # sets the title of the New Window widget
        newWindow.geometry("400x600+10+10")                                             # sets the geometry
        tk.Label(newWindow, text="Please enter the details: ", font=("Goudy old style", 13, "bold"), fg="red").pack()
        tk.Label(newWindow, text="Employee ID ").pack()

        tk.Label(newWindow, bd=5, text=set_emp_no, width=30, bg="yellow").pack()
        tk.Label(newWindow, text="Password ").pack()
        tk.Entry(newWindow, show='*', textvariable=emp_pwd_entry, bd=5, width=30, bg="dark grey").pack()
        tk.Label(newWindow, text="Date of Birth").pack()
        emp_dob_ent = DateEntry(newWindow, width=30, textvariable=emp_dob_entry, bd=5, background='darkblue', foreground='white', borderwidth=2).pack()
        tk.Label(newWindow, text="First Name ").pack()
        tk.Entry(newWindow, width=30, textvariable=emp_firstname_entry, bd=5, bg="dark grey").pack()
        tk.Label(newWindow, text="Last Name ").pack()
        tk.Entry(newWindow, width=30, textvariable=emp_lastname_entry, bd=5, bg="dark grey").pack()
        tk.Label(newWindow, text="E-mail ID ").pack()
        tk.Entry(newWindow, width=30, textvariable=emp_email_entry, bd=5, bg="dark grey").pack()
        tk.Label(newWindow, text="Date of Joining").pack()
        emp_doj_ent = DateEntry(newWindow, textvariable=emp_doj_entry, bd=5, width=12, background='darkblue', foreground='white', borderwidth=2).pack(padx=10, pady=10)
        tk.Label(newWindow, text="Department ").pack()
        tk.Entry(newWindow, width=30, textvariable=emp_dept_entry, bd=5, bg="dark grey").pack()
        tk.Button(newWindow, text="Save & Register", bg="yellow", bd=5, font=("lucidatypewriter", 15, "bold"), fg="red", command=self.confirm_registration).pack()

    def confirm_registration(self):  # Confirming info on click of register button
        emp_pwd_data = emp_pwd_entry.get()
        emp_dob_data = emp_dob_entry.get()
        emp_email_data = emp_email_entry.get()
        emp_doj_data = emp_doj_entry.get()
        emp_dept_data = emp_dept_entry.get()
        emp_firstname_data = emp_firstname_entry.get()
        emp_lastname_data = emp_lastname_entry.get()

        # importing the inputs to csv file
        with open('D:\Python\Abschluss_Project\demo_version\Employee_Registeration.csv', 'a',
                  newline="") as employee_file:  # newline = "\ n" switches off the traditional line break.
            emp_data_writer = csv.writer(employee_file, delimiter=',', quotechar='"', quoting=csv.QUOTE_NONE,
                                         escapechar="'")
            emp_data_writer.writerow(
                [set_emp_no, emp_pwd_data, emp_dob_data, emp_firstname_data, emp_lastname_data, emp_email_data,
                 emp_doj_data, emp_dept_data])

        newWindow.destroy()
        messagebox.showinfo(title="Registration Complete",
                            message="Congratulations, your Account has been registered successfully ! Please Enter Login details. ")
        pass

    def forgot_pwd(self):
        messagebox.showwarning("Forgot Password : ", "Please contact the Administrator to reset your account")

# This window creates a Login form, where new user can register themselves and can login.
# Further, existing user´s account is verified and authenticated
root = tk.Tk()
my_login_class = Login(root)

# creating personalized Login Window with Company Logo and Name
root.title(my_login_class.company_name)                                                  # sets personalized  title of the Login Window
bg_image = ImageTk.PhotoImage(file="company_background.png")         # Refer to Section5.9,“Images”(page.14).PIL documentation in Tkinter Documentation
login_bg_label = tk.Label(image=bg_image).pack()
login_label = tk.Label(root, text=my_login_class.company_name + " Employee Portal ", font=("Goudy old style", 18, "bold"), fg="#458B74").pack()

user_label = tk.Label(root, text="Employee ID", font=("lucidatypewriter", 15, "bold"), fg="magenta").pack()
user_entry = tk.Entry(root, bg="grey", font=("lucidatypewriter", 10, "bold"), textvariable=employee_no_verify).pack()
pwd_label = tk.Label(root, text="Password", font=("lucidatypewriter", 15, "bold"), fg="magenta").pack()
pwd_entry = tk.Entry(root, bg="grey", font=("lucidatypewriter", 10, "bold"), show="*", textvariable=password_verify).pack()

login_btn = tk.Button(root, text="Login", font=("lucidatypewriter", 15, "bold"), fg="blue", bg="light grey", command=my_login_class.verify_login).pack(ipadx=10)
new_reg_btn = tk.Button(root, text="New Registration", command=my_login_class.new_user_register, font=("lucidatypewriter", 12, "bold"), fg="brown").pack(side="right", padx=10, pady=10)
forget_btn = tk.Button(root, text="Forget Password", command=my_login_class.forgot_pwd, font=("lucidatypewriter", 12, "bold"), fg="brown").pack(side="left", padx=10, pady=10)

root.geometry("500x500+10+10")  # sets the geometry of Login Screen
root.resizable(0, 0)                                                                                    # disable Re-sizing option & background image isnt disturbed, we can also provide the value as "False" instead of "0"
root.mainloop()


def timeout():
    now = datetime.now()
    logout_time = (now.strftime("%H:%M:%S"),)  # assigning tuple variable
    loggedout_now = str(now)
    datetimeFormat = '%Y-%m-%d %H:%M:%S.%f'
    time_diff = datetime.strptime(loggedout_now, datetimeFormat) - datetime.strptime(my_login_class.loggedin_now, datetimeFormat)
    str_diff = str(time_diff)
    split_str = str_diff.split(".")

    hrs_worked = (split_str[0],)
    t = split_str[0].split(':')
    total_minutes = round((int(t[0]) * 60 + int(t[1]) * 1 + int(t[2]) / 60), 3)

    if total_minutes > 525:
        overtime = total_minutes - 525
        salary = round((525 * 0.5 + overtime * 1.2), 2)
        print(f"Hours Worked : {total_minutes} , Overtime hours : {overtime}, Salary : {salary}")
    else:
        overtime = 0
        salary = round((total_minutes * 0.5), 2)

    pause = "00:45:00"

    with open('D:\Python\Abschluss_Project\demo_version\Timesheet\Jun-2021\Employee_Timesheet_Tracker.csv',
              newline="") as employee_file:  # newline = "\ n" switches off the traditional line break.
        emp_data_writer = csv.reader(employee_file, delimiter=',')
        titlecolumn = next(emp_data_writer)  # this moves the cursor to headerrow and rules out the title column !!!
        for row in emp_data_writer:
            if my_login_class.emp_id == row[0]:
                if my_login_class.login_date == row[1]:
                    if my_login_class.login_time == row[2]:
                        with open(
                                'D:\Python\Abschluss_Project\demo_version\Timesheet\Jun-2021\Final_Timesheet_Tracker.csv',
                                'a', newline="") as my_file:
                            fill_data = csv.writer(my_file, delimiter=',', quotechar='"', quoting=csv.QUOTE_NONE,
                                                   escapechar="'")
                            fill_data.writerow(
                                [my_login_class.emp_id, my_login_class.login_date, my_login_class.login_time,
                                 logout_time, pause, hrs_worked, overtime, salary])
                        break
    messagebox.showinfo(title="Log out", message="Goodbye, you are successfully Logged out ! Have a nice time.")
    Main_Window.destroy()

# --- Part 2 ---

# This window is the Home Page of our Portal Software, where various kinds of applications can be provided.

Main_Window = tk.Tk()
Main_Window.title(f" ***       Employee Portal :  {my_login_class.emp_id}  ,  {my_login_class.first_name}  {my_login_class.last_name} - [  Logged in on :  {my_login_class.login_date} at {my_login_class.login_time} ]")
Main_Window.geometry("900x900")  # sets the geometry of Login Screen
Main_Window['bg'] = '#fb0'

tabControl = ttk.Notebook(Main_Window)
tab1 = ttk.Frame(tabControl)
tab2 = ttk.Frame(tabControl)
tab3 = ttk.Frame(tabControl)
tab4 = ttk.Frame(tabControl)
tab5 = ttk.Frame(tabControl)

tabControl.add(tab1, text='My Profile')
tabControl.add(tab2, text='Timesheet Check')
tabControl.add(tab3, text='Log out')


tabControl.pack(fill="both", padx=30, pady=30, ipadx=10, ipady=10)

def my_edit():
    messagebox.showinfo("Edit Window", "You have reached the Profile Edit Window")

def my_email():
    messagebox.showinfo("E-mail Box ", "E-mail here for any changes ! ")

with open('D:\Python\Abschluss_Project\demo_version\Employee_Registeration.csv', newline="") as employee_file:
    file_reader = csv.reader(employee_file, delimiter=',')
    titlecolumn = next(file_reader)                                         # this moves the cursor to header row and rules out the title column !!!
    for row in file_reader:
        if my_login_class.emp_id == row[0]:
            ttk.Label(tab1, text="Employee ID : ", style='LB.TButton').grid(column=0, row=0, padx=30, pady=30)
            ttk.Label(tab1, text=row[0], style='LB.TButton').grid(column=2, row=0, padx=30, pady=30)
            ttk.Label(tab1, text="First Name : ").grid(column=0, row=2, padx=30, pady=30)
            ttk.Label(tab1, text=row[3]).grid(column=2, row=2, padx=30, pady=30)
            ttk.Label(tab1, text="Last Name : ").grid(column=0, row=4, padx=30, pady=30)
            ttk.Label(tab1, text=row[4]).grid(column=2, row=4, padx=30, pady=30)
            ttk.Label(tab1, text="Date of Birth : ").grid(column=0, row=6, padx=30, pady=30)
            ttk.Label(tab1, text=row[2]).grid(column=2, row=6, padx=30, pady=30)
            ttk.Label(tab1, text="Date of Joining :").grid(column=0, row=8, padx=30, pady=30)
            ttk.Label(tab1, text=row[6]).grid(column=2, row=8, padx=30, pady=30)
            ttk.Label(tab1, text="E-mail ID : ").grid(column=0, row=10, padx=30, pady=30)
            ttk.Label(tab1, text=row[5]).grid(column=2, row=10, padx=30, pady=30)
            ttk.Label(tab1, text="Department : ").grid(column=0, row=12, padx=30, pady=30)
            ttk.Label(tab1, text=row[7]).grid(column=2, row=12, padx=30, pady=30)

if my_login_class.emp_id == "102":
    photo = tk.PhotoImage(file='profile_bild.png')                  # Refer to Section5.9,“Images”(page.14).PIL documentation in Tkinter Documentation
    image_label = ttk.Label(tab1, image=photo)
    image_label.grid(row=0, column=40)
elif my_login_class.emp_id =="101":
    photo = tk.PhotoImage(
        file='profile_pic2.png')  # Refer to Section5.9,“Images”(page.14).PIL documentation in Tkinter Documentation
    image_label = ttk.Label(tab1, image=photo)
    image_label.grid(row=0, column=40)
ttk.Button(tab1, text="Edit Profile", command=my_edit).grid(row=14, column=20)

l1 = ttk.LabelFrame(tab2, text="Track and Plot Timesheet : ")
l1.grid(row=10, column=0, padx=30, pady=30, ipadx=10, ipady=10)
mycombo = ttk.Combobox(l1, width=20)
mycombo["values"] = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October",  "November", "December"]
mycombo.grid(row=14, column=2, padx=10, pady=10)
mycombo.current(0)

# ttk.LabelFrame(tab2, text="Year : ").grid(row=5, column=0, columnspan=2)
mycombo_1 = ttk.Combobox(l1, width=20)
mycombo_1["values"] = ["2021", "2020", "2019", "2018", "2017", "2016", "2015"]
mycombo_1.grid(row=14, column=8, padx=10, pady=10)
mycombo_1.current(0)
date_list, worked_list, overtime = [], [], []
my_choice = tk.StringVar()
my_choice = mycombo.get()

label3 = ttk.LabelFrame(tab2, text="Recent Login / Logout Details : ")
label3.grid(row=15, column=20, padx=30, ipadx=10, ipady=10)

tree = ttk.Treeview(label3)
tree["columns"] = ("one", "two", "three", "four", "five", "six")
tree.column("#0", width=0, minwidth=0, stretch=tk.NO)
tree.column("one", width=90, minwidth=80, stretch=tk.NO)
tree.column("two", width=120, minwidth=90, stretch=tk.NO)
tree.column("three", width=120, minwidth=90, stretch=tk.NO)
tree.column("four", width=120, minwidth=90, stretch=tk.NO)
tree.column("five", width=120, minwidth=90, stretch=tk.NO)
tree.column("six", width=120, minwidth=90, stretch=tk.NO)

tree.heading("#0", text="", anchor=tk.W)
tree.heading("one", text="Employee ID ", anchor=tk.W)
tree.heading("two", text="Log-in Date ", anchor=tk.W)
tree.heading("three", text="Login Time", anchor=tk.W)
tree.heading("four", text="Logout Time", anchor=tk.W)
tree.heading("five", text="Total Hours Worked", anchor=tk.W)
tree.heading("six", text="Overtime in Mins", anchor=tk.W)

tree_id = 0
with open('D:\Python\Abschluss_Project\demo_version\Timesheet\Jun-2021\Final_Timesheet_Tracker.csv',
          newline="") as employee_file:
    file_reader = csv.reader(employee_file, delimiter=',')
    titlecolumn = next(file_reader)  # this moves the cursor to header row and rules out the title column !!!
    for row in file_reader:
        if my_login_class.emp_id == row[0]:
            tree.insert(parent="", index=tree_id, values=(row[0], row[1], row[2], row[3], row[5], row[6]))
            tree_id += 1
            tree.grid(row=18, column=0, padx=30, pady=20)

def my_print():
    messagebox.showinfo("Print Page", "You have reached the Print Page Window")

def open_file():
    filedialog.askopenfilename(filetypes=[("csv files", "*.csv")], initialdir="D:\Python\Abschluss_Project\demo_version\Timesheet\Sample_Database", title="Open File")

label2 = ttk.LabelFrame(tab2, text="Plot Window:  ")
label2.grid(row=15, column=0, padx=30)

class matplot_in_tkinter:
    """Plot Window"""
    def __init__(self, masterframe):
        global top
        self.figure = matplotlib.pyplot.Figure(facecolor="beige", dpi=120, constrained_layout =True)
        self.axis = self.figure.add_subplot(111)
        # create canvas as matplotlib drawing area - using `fig` and assign to widget `tab2`
        self.canvas = FigureCanvasTkAgg(self.figure, master=masterframe)
        # get canvas as reference to tkinter widget and put in widget `tab2`
        self.canvas.get_tk_widget().grid()
        # create toolbar
        self.toolbar = NavigationToolbar2Tk(self.canvas, window= masterframe, pack_toolbar=False)
        self.toolbar.grid()
        self.canvas._tkcanvas.grid()

    def my_clear(self):
        """clear figure"""
        self.axis.cla()
        self.canvas.draw()

    def view_plot(self):
        with open('D:\Python\Abschluss_Project\demo_version\Timesheet\Sample_Database\Employee_Timesheet.csv',
                  newline="") as csv_file:
            csv_reader = csv.reader(csv_file, delimiter=",")
            x, y, z = [], [], []
            Title_Col = next(csv_reader)
            for row in csv_reader:
                if my_login_class.emp_id == row[0]:
                    if my_choice == row[1]:
                        d = datetime.strptime(row[2], "%d.%m.%Y")
                        x.append(d)
                        y.append(float(row[6]))
                        z.append(float(row[7]))
        return x, y, z

    def tot_hrs_bar_graph(self):
        self.my_clear()
        a, b,c = self.view_plot()
        # draw on this plot
        self.axis.bar(a, b)
        self.axis.set_xlabel('Dates')
        self.axis.set_ylabel('Total Hours Worked')

        self.axis.tick_params(axis='x', rotation=45, colors='r', color='b')
        self.axis.tick_params(axis='y',  colors='g')

        self.axis.grid()
        self.canvas.draw()

    def overtime_bar_graph(self):
        self.my_clear()
        a, b, c = self.view_plot()
        # draw on this plot
        self.axis.bar(a, c)
        self.axis.tick_params(axis='x', rotation=45, colors='r', color='b')
        self.axis.tick_params(axis='y', colors='g')
        self.axis.set_xlabel('Dates')
        self.axis.set_ylabel('Overtime in mins')
        self.axis.grid()
        self.canvas.draw()

def future_enhance():
    messagebox.showinfo("Message", "You have reached this option, under future enhancements. Stay logged-in to Experience this Service")

# --- canvas and toolbar in my Notebook:(TAB2 Frame) ---
# create figure
#top frame for canvas and toolbar
top = ttk.Frame(label2)
top.grid(row=22, column=2,  padx=30, pady=20)
my_plot_class = matplot_in_tkinter(top)
# bottom frame for other widgets
bottom = ttk.Frame(label2)
bottom.grid(row=38, column=5)

# --- other widgets in bottom ---
ttk.Button(bottom, text='Print', command=my_print).grid(row=40, column=12)
ttk.Button(bottom, text='Clear', command=my_plot_class.my_clear).grid(row=40, column=16)

mb = ttk.Menubutton(l1, text="Plot Sheet", width=30)
mb.menu = tk.Menu(mb)
mb["menu"] = mb.menu

mb.menu.add_command(label="Pie Chart", command=future_enhance)
mb.menu.add_command(label="Stacked Chart", command=future_enhance)
mb.menu.add_command(label="Grouped Bar Chart", command=future_enhance)
mb.menu.add_separator()
mb.menu.add_command(label="Total Hours Bar Graph", command=my_plot_class.tot_hrs_bar_graph)
mb.menu.add_command(label="Overtime Bar Chart", command=my_plot_class.overtime_bar_graph)
mb.menu.add_separator()
mb.menu.add_command(label="2D Chart", command=future_enhance)
mb.grid(row=14, column=14, padx=10, pady=10)

ttk.Button(l1, text="Send E-mail", command=my_email).grid(row=14, column=22)
ttk.Button(l1, text="Open Sheet", command=open_file).grid(row=14, column=18)

# This will be adding style, and naming that style variable
my_style = ttk.Style()
# Widget.Tbutton (TButton is used for ttk.Button).
my_style.configure('LB.TButton', font=('calibri', 18, 'bold'), foreground='red')

"""Log-out"""
b1 = btn_logout = ttk.Button(tab3, text="Log Out", command=timeout, style='LB.TButton')
b1.grid(row=5, column=5, sticky="ew", padx=15, pady=15)

Main_Window.mainloop()
