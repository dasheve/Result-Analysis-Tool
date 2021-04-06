
#Importing necessary libraries
# details of each column
# screenshot of format

import tkinter as tk
from tkinter import ttk
from PIL import ImageTk, Image
import pandas as pd
import openpyxl
from matplotlib.backends.backend_tkagg import (FigureCanvasTkAgg, NavigationToolbar2Tk)
from tkscrolledframe import ScrolledFrame
from Data import getdetails, getcourse, get_rank, get_studentReport, studentreport, pdfupload
from tkinter.filedialog import askopenfilename
import sys
import os
import webbrowser

url="https://github.com/dasheve/Result-Analysis-Tool#readme"
def openweb():
    webbrowser.open(url, new=1)


#function for all comboboxes on tha pages
def combobox(loc, val, default, x1, y1, call):    
    '''
    Parameters----------
    loc : object
        the parent frame 
    val : list
        list of values in combobox
    default : string
        deafult value of combobox
    x1 : int
        x-coordinate for position
    y1 : int
        y-coordinate for position
    call : object
        function to retrieve value of combobox
    Returns :
        combo : tkinter object
            the comobox with all its properties
    '''
    combo=ttk.Combobox(loc, values=val)
    combo.set(default)
    combo.configure(width=30, font=("Helvetica",14))
    combo.place(x=x1, y=y1)
    combo.bind("<<ComboboxSelected>>", call)
    return combo

#function for buttons on frames
def buttons(loc, path, x1,y1,width,height, go):
    '''
    Parameters----------
    path : string
        the path of image file
    x1 : int
        x-coordinate
    y1 : TYPE
        y-coordinate
    width : int/float
        width of image
    height : int/float
        height of image
    go : object
        destination frame
    Returns : None
    '''
    img=Image.open("Buttons/"+path).resize((width, height))
    img_op=ImageTk.PhotoImage(img, master=loc)
    button=tk.Button(loc, image=img_op, bd=0, bg='white', command=go)
    button.image=img_op
    button.place(x=x1,y=y1)
    
#function for icons on top and left
def icons(loc, path, x1,y1, go):
    '''
    Parameters----------
    path : string
        the path of image file
    x1 : int
        x-coordinate
    y1 : TYPE
        y-coordinate
    go : object
        destination frame
    Returns : None
    '''
    img=Image.open("Icons/"+path)
    img_op=ImageTk.PhotoImage(img, master=loc)
    button=tk.Button(loc, image=img_op, bd=0, bg='white', command=go)
    button.image=img_op
    button.place(x=x1,y=y1)
    
def TopLeftFrame(loc, x1, y1, pdx, pdy, wi, hi):
    """
    Parameters----------
    loc : object
        root window name
    x1 : int
        x-coordinate for plaing the frame on the window
    y1 : int
        y-coordinate for plaing the frame on the window
    pdx : int
        x-padding
    pdy : int
        y-padding
    wi : int
        width of the object
    hi : int
        height of the object
    Returns :
    -------
    frame_icon : tkinter.labelframe object
        labelframe for putting elements over it
    """
    frame_icon=tk.LabelFrame(loc, padx=pdx, pady=pdy, width=wi, height=hi, bg='white', bd=0)
    frame_icon.place(x=x1, y=y1)
    return frame_icon

def bgimage(path, w, h, root, x1, y1 ):
    """
    Parameters----------
    path : string
           path of the image file
    w : int
        width of image in pixels
    h : int
        height of image in pixels
    root : object
           tkinter window 
    x1 : int
         x-coordinate 
    y1 : int
         y-coordinate
    Returns : None
    """
    img=Image.open(path).resize((w,h))
    my_img=ImageTk.PhotoImage(img, master=root)
    my_label=tk.Label(root, image=my_img, bd=0, bg="#F6F8FB")
    my_label.sample=my_img
    my_label.place(x=x1, y=y1)

class tkinterApp(tk.Tk): 

    # __init__ function for class tkinterApp 
    def __init__(self): 
        
        # __init__ function for class Tk 
        tk.Tk.__init__(self) 
        
        # creating a container 
        self.container = tk.Frame(self)
        self.container.pack(side = "top", fill = "both", expand = True)
        self.geometry("1440x1024")
        
        self.container.grid_rowconfigure(0, weight = 1)
        self.container.grid_columnconfigure(0, weight = 1) 
        
        # initializing frames to an empty array
        self.frames = {}
        
        # iterating through a tuple consisting 
        # of the different page layouts 
        for F in (HomePage, Upload): 
            
            frame = F(self.container, self) 
            # initializing frame of that object from 
            # HomePage, Chart respectively with 
            # for loop 
            self.frames[F] = frame 
            
            frame.grid(row = 0, column = 0, sticky ="nsew") 
            self.show_frame(HomePage)
            
    # to display the current frame passed as 
	# parameter 
    def show_frame(self, cont): 
        frame=self.frames[cont]
        frame.tkraise()
        
    # to show the Student Data Page
    def showStudent(self, args, cont):
        frame=cont(self.container, self, args)
        frame.grid(row = 0, column = 0, sticky ="nsew")
        frame.tkraise()
        
    # to show the Student Selection Page
    def showStudent2(self, args):
        frame=Student2(self.container, self, args)
        frame.grid(row = 0, column = 0, sticky ="nsew")
        self.frames[Student2]=frame
        frame.tkraise()
        
    # to destroy the current frame and show the previous one    
    def back_frame(self, args):
        args.destroy()
                    
"""
This class denotes the Home Page or Landing Page
which enables the user to select the correct dataframe
and then proceed towards the desired visualisation.
"""
class  HomePage(tk.Frame):
    # init function of class HomePage
    def __init__(self, parent, controller): 
        """
        Parameters
        ----------
        parent : object
                 The parent window for the frame
        controller : object
                     Enables navigation between pages
        args : object, optional
               Values to be taken from one class to another

        Returns : None
        """
        
        # initialising the frame 
        tk.Frame.__init__(self, parent)
        self.parent=parent
        self.controller=controller
        self.configure(bg='#F6F8FB')  
        
        
        self.c_values={'course':'G582_Physical Science', 'year':2019, 'college':'Acharya Narendra Dev College'}
        self.student_combo_values=['16001582057']
        
        # the default vertical toolbar which enables to select from the given options        
        frame_icon=TopLeftFrame(self, 0, 0, 10, 10, 60, 1024)

        # button for menu page
        icons(frame_icon, "menu.png", 4, 10, lambda : controller.show_frame(StudentPage))

        # button for user details
        icons(frame_icon, "user.png", 4, 60, None)

        # button to navigate to the homepage
        icons(frame_icon, "home_alt_fill.jpg", 4, 110, None)

        # button to navigate to the Upload page
        buttons(frame_icon, "Upload.png", 4, 160, 23, 23, lambda : controller.show_frame(Upload))

        # button for settings page 
        icons(frame_icon, "settings.png", 4, 210, None)
        
        # image placement
        bgimage("Images/backimage.png", 600, 720, self, 680, 0)
        
        # the default horizontal toolbar at the top
        frame_top=TopLeftFrame(self, 60, 0, 10, 0, 1220, 55)
        
        # button for showing notifications
        icons(frame_top, "notification.png", 1160, 15, None)

        # button for showing information about the page
        icons(frame_top, "info.png", 1110, 15, None)
        
        # setting default text on homepage
        my_label6=tk.Label(self, text="Welcome to", font=("Times New Roman", 25), bg='#F6F8FB' )
        my_label6.place(x=130, y=110)
        
        my_label7=tk.Label(self, text="Result Analyzer", font=("Times New Roman", 25), bg='#F6F8FB', fg='#5DB447')
        my_label7.place(x=300, y=110)
        
        my_label8=tk.Label(self, text="Choose", font=("Times New Roman", 30, 'bold'), bg='#F6F8FB', fg='black')
        my_label8.place(x=130, y=150)
        
        
        # setting styles for combobox
        combostyle=ttk.Style()
        combostyle.theme_create('Art', parent='clam', settings = {'TCombobox':{'configure':{'selectbackground': 'Blue','fieldbackground': '#F6F8FB', 'background': '#F6F8FB', 'arrowsize':20}}})
        combostyle.theme_use('Art')
        combostyle.configure('TCombobox', relief='flat', foreground='Black', padding=5, bordercolor='#032D23')
        combostyle.configure('Vertical.TScrollbar', arrowsize=10)
        
        # creating comboboxes to select from given values
        
        # Combobox for selection of College
        self.colleges=('Bhagini Nivedita College',
                        'Hans Raj College',
                        'Hindu College',    
                        'Kalindi College',
                        'Keshav Mahavidyalaya',
                        'Kirori Mal College',
                        'Maharaja Agrasen College',
                        'Miranda House ',
                        'Rajdhani College',
                        'Ramjas College ',
                        'S.G.T.B. Khalsa College',
                        'Shivaji College ',
                        'Shyam Lal College (Day)',
                        'St. Stephens College',
                        'Swami Shraddhanand College',
                        'Zakir Husain Delhi College (Day',
                        'I.P.College For Women',
                        'Keshav Mahavidyalaya',
                        'Shyama Prasad Mukherjee College',
                        'Sri Guru Gobind Singh College',
                        'Acharya Narendra Dev College',
                        'Atma Ram Sanatan Dharam College',
                        'Deen Dayal Upadhyaya College',
                        'Deshbandhu College (Day) ',
                        'Dyal Singh College (Day)',
                        'Gargi College',
                        'Maitreyi College   ',
                        'Moti Lal Nehru College (Day) ',
                        'Sri Aurobindo College (Day)',
                        'Bhaskaracharya College of Appli',
                        'College Of Vocational Studies',
                        'Ramanujan College',
                        'P.G.D.A.V. College (Day)',
                        'Ram Lal Anand College (Day)',
                        'Aryabhatta College Formerly Ram',
                        'Shaheed Rajguru College of Appl',
                        'Shaheed Sukhdev College of Busi')
            
        
        def open_file(txt):
                    file1=open(txt,"r")
                    lst=list()
                    for i in file1.readlines():
                        lst.append(i.rstrip())
                    return lst
            
        # Combobox for selection of College 
        self.college_combo=combobox(self, open_file("Files/Colleges.txt"), "Choose College", 135, 240, self.callback1)
        
        # Combobox for selection of Course           
        self.course_combo=combobox(self, open_file("Files/Courses.txt"), "Choose Course", 135, 300, self.callback2)
        
        # Combobox for selection of Year
        self.year_combo=combobox(self, ['2019'], "Choose Year", 135, 360, self.callback3)

        # Combobox for selection of Part
        self.part_combo=combobox(self, ['First Year', 'Second Year','Passout'], "Choose Course Year", 135, 420, self.callback4)
        
        # button to navigate to student selection page after selecting
        # values from the comboxes
        buttons(self, "student_button.png", 135, 480, 170, 50, lambda: self.warning('stu'))
            
        # button to navigate to Course Details page
        buttons(self, "course_button.png", 328, 480, 170, 50, lambda: self.warning('cou'))

        # button to navigate to Ranking page
        buttons(self, "ranking_button.png", 135, 540, 170, 50, lambda: controller.showStudent(args=self.c_values, cont=Ranking))

        # setting some default text
        info_label1=tk.Label(self, text="Wanna Know More?", font=("Helvetica", 12), bg='#F6F8FB', fg='Black' )
        info_label1.place(x=195, y=620)
        
        # button to know more about the functioning of the app
        info_label1_button=tk.Button(self, text="Click Here", font=("Helvetica", 12), bg='#F6F8FB', fg='#FF6600', bd=0, command=openweb)
        info_label1_button.place(x=340, y=618.45)
        
    def warning(self, txt):
            if getdetails(self.c_values['course'],self.c_values['year'], self.c_values['part'],'')=="FileError":
                tk.messagebox.showwarning("File Not Found", "File does not exist. Please upload the required file.")
            else:
                if self.c_values['college'] not in getdetails(self.c_values['course'],self.c_values['year'], self.c_values['part'],''):
                    tk.messagebox.showwarning("Course Not Found", "The selected course is not in the selected college.")
                else:
                    if txt=='cou':
                        self.controller.showStudent(args=self.c_values, cont=Course)
                    elif txt=='stu':
                        self.controller.showStudent2(args=self.c_values)
            
    def callback1(self, *args): 
        """
        Class Method
        Receives the value from the College Selection Combobox
        Returns: None
        """
        self.college=self.college_combo.get()
        self.c_values['college']=self.college 
        
    def callback2(self, *args):
        """
        Class Method
        Receives the value from the Course Selection Combobox
        Returns: None
        """
        self.course=self.course_combo.get()
        self.c_values['course']=self.course
        
    def callback3(self, *args):
        """
        Class Method
        Receives the value from the Year Selection Combobox
        Returns: None
        """
        self. year=self.year_combo.get()
        self.c_values['year']=self.year
        
    def callback4(self, *args):
        """
        Class Method
        Receives the value from the Part Selection Combobox
        Returns: None
        """
        self.part=self.part_combo.get()
        part_dict={'First Year':1, 'Second Year':2, 'Passout':3}
        self.c_values['part']=part_dict[self.part]
                
"""
The Student Page showing all the result details
of the selected student.
"""
class StudentPage(tk.Frame):
  
    # init function for the class StudentPage
    def __init__(self, parent, controller, args=None):
        """
        Parameters
        ----------
        parent : object
                 The parent window for the frame
        controller : object
                     Enables navigation between pages
        args : object, optional
               Values to be taken from one class to another

        Returns : None
        """
        
        # initialising the Frame
        tk.Frame.__init__(self, parent)
        self.controller=controller
        self.parent=parent
        
        # setting the background color of the frame
        self.configure(bg='#F6F8FB')
        
        # taking the selected values from the HomePage and student2 Page
        # as argument and storing it in a variable
        values=args
        values['student']=int(values['student'])
        
        # calling the required imported function and storing the
        # returned values in a variable
        s_r=get_studentReport(values["course"],values["year"],values["part"],values["college"],int(values["student"]) )
        if s_r[0]['Fail']==True:
                top1=tk.Toplevel()
                top1.title("Error")
                top1.geometry("700x500")
                top1.configure(bg='white')
                top1.resizable(0,0)
                
                tk.Label(top1, text="Error IN:", bg="white", font=("Helvetica", 15)).place(x=20, y=50)
                for i in values:
                    tk.Label(top1, text=str(i)+" - ", bg='white', font=("Helvetica", 14)).place(x=20, y=100+list(values.values()).index(values[i])*40)
                    tk.Label(top1, text=values[i], bg='white', font=("Helvetica", 14)).place(x=140, y=100+list(values.values()).index(values[i])*40)
                tk.Label(top1, text="file - ", bg='white', font=("Helvetica", 14)).place(x=20, y=300)
                tk.Label(top1, text=s_r[1]['path'], bg='white', font=("Helvetica", 14)).place(x=140, y=300)
            
        
        # the default vertical toolbar which enables to select from the given options        
        frame_icon=TopLeftFrame(self, 0, 0, 10, 10, 60, 1024)

        # button for menu page
        icons(frame_icon, "menu.png", 4, 10, None)

        # button for user details
        icons(frame_icon, "user.png", 4, 60, None)

        # button to navigate to the homepage
        icons(frame_icon, "home_alt_fill.jpg", 4, 110, lambda : controller.show_frame(HomePage))

        # button to navigate to the Upload page
        buttons(frame_icon, "Upload.png", 4, 160, 23, 23, lambda : controller.show_frame(Upload))

        # button for settings page 
        icons(frame_icon, "settings.png", 4, 210, None)
                    
        # the default horizontal toolbar at the top
        frame_top=TopLeftFrame(self, 60, 0, 10, 0, 1220, 55)
        
        # button for showing notifications
        icons(frame_top, "notification.png", 1160, 15, None)

        # button for showing information about the page
        icons(frame_top, "info.png", 1110, 15, None)
        
        # button to navigate to the previous page        
        icons(frame_top, "back.png", 10, 15, lambda: controller.back_frame(self))
    
        # setting default labels of name, roll number,
        # college, course and status
        name_label=tk.Label(self, text='Name', bg='#F6F8FB', font=('Helvetica',12), fg='#9FA2B4').place(x=100, y=80)
        roll_label=tk.Label(self, text='Roll Number', bg='#F6F8FB', font=('Helvetica',12), fg='#9FA2B4').place(x=100, y=105)
        college_label=tk.Label(self, text='College', bg='#F6F8FB', font=('Helvetica',12), fg='#9FA2B4').place(x=100, y=130)
        course_label=tk.Label(self, text='Course', bg='#F6F8FB', font=('Helvetica',12), fg='#9FA2B4').place(x=100, y=155)
        status_label=tk.Label(self, text='Status', bg='#F6F8FB', font=('Helvetica',12), fg='#9FA2B4').place(x=100, y=180)
        
        # showing the details of the selected student
        name=str(s_r[0]["name"])
        if '\n' in name:
            names=name.split('\n')[0]
        else:
            names=name.split(' ')[0:2]
            
        enter_name_label=tk.Label(self, text=names, bg='#F6F8FB', font=('Helvetica',16)).place(x=200, y=80)
        enter_roll_label=tk.Label(self, text=str(s_r[0]["rollno"]), bg='#F6F8FB', font=('Helvetica',16)).place(x=200, y=105)
        enter_college_label=tk.Label(self, text=values['college'], bg='#F6F8FB', font=('Helvetica',16)).place(x=200, y=130)
        enter_course__label=tk.Label(self, text=values['course'], bg='#F6F8FB', font=('Helvetica',16)).place(x=200, y=155)
        enter_status_label=tk.Label(self, text='Passed', bg='#F6F8FB', font=('Helvetica',16)).place(x=200, y=180)
                                    
        # getting the plot from the imported functions from other
        # file and setting its width and height
        ax1=s_r[0]["plot"].figure  
        ax1.set_figheight(3)
        ax1.set_figwidth(6)
        # creating a plot to set the plot on
        frame_plot1=tk.LabelFrame(self, padx=0, pady=0, width=437, height=220, bg='white', bd=2)
        # creating a canvas object and setting plot on it
        canvas = FigureCanvasTkAgg(ax1, master=frame_plot1)   
        canvas.draw() 
        # placing the canvas on the frame
        canvas.get_tk_widget().place(x=0, y=0)
        # placing the canvas on the frame
        frame_plot1.place(x=100, y=430)
        
        ax2=s_r[0]["percplot"]
        ax2.set_figheight(2.5)
        ax2.set_figwidth(3)
        frame_plot2=tk.LabelFrame(self, padx=0, pady=0, width=220, height=185, bg="white", bd=2)
        canvas2 = FigureCanvasTkAgg(ax2, master=frame_plot2)  
        canvas2.draw() 
        canvas2.get_tk_widget().place(x=0, y=0)
        frame_plot2.place(x=100, y=230)
        
        title_frame=tk.LabelFrame(self, padx=3, pady=2, width=500, height=30, bg="#F6F8FB", bd=1, relief=tk.RAISED)
        title_frame.place(x=700, y=260)
        
        # creating LabelFrames and labels to show yearwise 
        # SGPA distribution of each object
        
        #print(len(s_r[0]['ER'].split(',')))

        tk.Label(title_frame, text="Year              Semester              SGPA              CGPA", font=('Helvetica', 11), bg="#F6F8FB").place(x=10, y=3)
        
        year1_frame=tk.LabelFrame(self, padx=3, pady=2, width=500, height=100, bg="#F6F8FB", bd=1, relief=tk.RAISED)
        year1_frame.place(x=700, y=305)
        
        year2_frame=tk.LabelFrame(self, padx=3, pady=2, width=500, height=100, bg="#F6F8FB", bd=1, relief=tk.RAISED)
        year2_frame.place(x=700, y=425)
        
        year3_frame=tk.LabelFrame(self, padx=3, pady=2, width=500, height=100, bg="#F6F8FB", bd=1, relief=tk.RAISED)
        year3_frame.place(x=700, y=545)
        
        year1_label=tk.Label(year1_frame, padx=7, pady=7, text="I", bg="#F6F8FB", font=("Helvetica", 16))
        year1_label.place(x=15, y=27)
        
        year2_label=tk.Label(year2_frame, padx=7, pady=7, text="II", bg="#F6F8FB", font=("Helvetica", 16))
        year2_label.place(x=15, y=27)
        
        year3_label=tk.Label(year3_frame, padx=7, pady=7, text="III", bg="#F6F8FB", font=("Helvetica", 16))
        year3_label.place(x=15, y=27)
        
        # displaying percentage of each student
        def sem_perc(frame, sem, txt, x1, y1):
             sem_perc=str(float("{:.2f}".format((s_r[sem]['perc'])/10)))
             sem=str(txt+sem_perc)
             label_sem_perc=tk.Label(frame, text=sem, bg="#F6F8FB", font=('Helvetica, 13')).place(x=x1, y=y1)
            
        sem_perc(year1_frame, 1, "I                   ", 125, 20)
        sem_perc(year1_frame, 2, "II                  ", 124, 55)
        sem_perc(year2_frame, 3, "III                 ", 125, 20)
        sem_perc(year2_frame, 4, "IV                 ", 124, 55)
        sem_perc(year3_frame, 5, "V                  ", 125, 20)
        sem_perc(year3_frame, 6, "VI                 ", 124, 55)
        
        tk.Label(year1_frame, text=str(float("{:.2f}".format((s_r[1]['perc']/10+s_r[2]['perc']/10)/2))), bg="#F6F8FB", font=("Helvetica", 15)).place(x=317, y=30)
        tk.Label(year2_frame, text=str(float("{:.2f}".format((s_r[3]['perc']/10+s_r[4]['perc']/10)/2))), bg="#F6F8FB", font=("Helvetica", 15)).place(x=317, y=30)
        tk.Label(year3_frame, text=str(float("{:.2f}".format((s_r[5]['perc']/10+s_r[6]['perc']/10)/2))), bg="#F6F8FB", font=("Helvetica", 15)).place(x=317, y=30)
        
        # defining the function for a toplevel window on the student page
        # to show grade distribution plots in each semester
        def toplevel(a,b):
            """
            Parameters
            ----------
            a : int
                Semester number to show the plot for distribution of grade points among 
                each subject for that semester.
            b : int
                Semester number to show the plot for distribution of grade points among 
                each subject for that semester.
            Returns : None
            """
            # defining the window
            top1=tk.Toplevel()
            title=str()
            if a==1:
                title='Year 1'
            elif a==3:
                title="Year 2"
            elif a==5:
                title="Year 3"
            # setting the title for the window
            top1.title(title)
            # setting the geometry
            top1.geometry("830x600")
            # setting parameters 
            top1.configure(bg='white')
            # disabling resize option
            top1.resizable(0,0)
            #top1.grab_set()     # opens up only one toplevel window at a time
            
            sem_a="Semester "+str(a)
            tk.Label(top1, text=sem_a, font=("Helvetica", 12), bg='white').place(x=30, y=132)
    
            ax3=s_r[a]["plot"].figure
            ax3.set_dpi=100
            canvas3 = FigureCanvasTkAgg(ax3, master=top1)   
            canvas3.draw() 
            canvas3.get_tk_widget().place(x=150, y=0)
            
            sem_b="Semester "+str(b)
            tk.Label(top1, text=sem_b, font=("Helvetica", 12), bg='white').place(x=30, y=422)
            
            ax4=s_r[b]["plot"].figure
            ax4.set_dpi=100
            canvas4 = FigureCanvasTkAgg(ax4, master=top1)   
            canvas4.draw() 
            canvas4.get_tk_widget().place(x=150, y=290)
            
            top1.mainloop()

        # creating buttons to open semester plots        
        buttons(year1_frame, "expand1.png", 410, 60, 75, 25, lambda : toplevel(1,2))
        buttons(year2_frame, "expand1.png", 410, 60, 75, 25, lambda : toplevel(3,4))
        buttons(year3_frame, "expand1.png", 410, 60, 75, 25, lambda : toplevel(5,6))
        
        
"""
The frame which opens after choosing the college, course, year,
for student result evaluation which enables the user to select
the particular student.
"""              
class Student2(tk.Frame):
    
    def __init__(self, parent, controller, args=None):
        
        """
        Parameters
        ----------
        parent : object
                 The parent window for the frame
        controller : object
                     Enables navigation between pages
        args : object, optional
               Values to be taken from one class to another

        Returns : None
        """

        # initialising the Frame
        tk.Frame.__init__(self, parent)
        self.controller=controller
        self.parent=parent
        self.configure(bg='#F6F8FB')
        
        # taking the selected values from the HomePage as argument
        # and storing it in a variable
        self.c_values=args

        # the default vertical toolbar which enables to select from the given options       
        frame_icon=tk.LabelFrame(self, padx=10, pady=10, width=60, height=1024, bg='white', bd=0)
        frame_icon.place(x=0, y=0)
        
        # the default vertical toolbar which enables to select from the given options        
        frame_icon=TopLeftFrame(self, 0, 0, 10, 10, 60, 1024)

        # button for menu page
        icons(frame_icon, "menu.png", 4, 10, None)

        # button for user details
        icons(frame_icon, "user.png", 4, 60, None)

        # button to navigate to the homepage
        icons(frame_icon, "home_alt_fill.jpg", 4, 110, lambda : controller.show_frame(HomePage))

        # button to navigate to the Upload page
        buttons(frame_icon, "Upload.png", 4, 160, 23, 23, lambda : controller.show_frame(Upload))

        # button for settings page 
        icons(frame_icon, "settings.png", 4, 210, None)
                    
        # the default horizontal toolbar at the top
        frame_top=TopLeftFrame(self, 60, 0, 10, 0, 1220, 55)
        
        # button for showing notifications
        icons(frame_top, "notification.png", 1160, 15, None)

        # button for showing information about the page
        icons(frame_top, "info.png", 1110, 15, None)
        
        # button to navigate to the previous page        
        icons(frame_top, "back.png", 10, 15, lambda: controller.back_frame(self))

        # combobox to select the name of student,
        # the name of college and course taken from
        # the HomePage
        names=getdetails(self.c_values['course'], self.c_values['year'],self.c_values['part'], self.c_values['college'])
        new_names=list()
        for i in names:
            if '\n' in i:
                new_names.append(i.split('\n')[0])
            else:
                new_names.append(i.split(' ')[0:4])
                
        # student selection combobox
        self.student_combo=combobox(self, new_names, "Choose Student", 135, 329, self.callback4)
        
        # button to take to the student page
        buttons(self, "login_button.png", 135, 400, 360, 50, lambda: controller.showStudent(args=self.c_values, cont=StudentPage))

        
    def callback4(self, *args):
        """
        Parameters
        ----------
        *args : object
        adds the student name to the dictionary
        Returns : None
        """
        self.student=self.student_combo.get()
        self.c_values['student']=self.student.split('.')[0]   
 
"""
The frame showing the result of entire course,
on clicking the course button.
"""

class Course(tk.Frame):
    
    def __init__(self, parent, controller, args=None):
        
        """
        Parameters
        ----------
        parent : object
                 The parent window for the frame
        controller : object
                     Enables navigation between pages
        args : object, optional
               Values to be taken from one class to another

        Returns : None
        """
        # initialising the frame
        tk.Frame.__init__(self, parent)
        self.controller=controller
        self.parent=parent
        self.configure(bg='#F6F8FB')
        
        # getting the values from HomePage
        # from the argument
        self.values=args
        cr=getcourse(self.values['course'],self.values['year'],self.values['part'],self.values['college'])

        if cr['Fail']==True:
                top1=tk.Toplevel()
                top1.title("Error")
                top1.geometry("700x500")
                top1.configure(bg='white')
                top1.resizable(0,0)
                
                tk.Label(top1, text="DataError:", bg="white", font=("Helvetica", 14)).place(x=20, y=50)
                tk.Label(top1, text=cr["DataError"], bg="white", font=("Helvetica", 14)).place(x=20, y=90)

        
        # the default vertical toolbar which enables to select from the given options       
        frame_icon=tk.LabelFrame(self, padx=10, pady=10, width=60, height=1024, bg='white', bd=0)
        frame_icon.place(x=0, y=0)
        
        # the default vertical toolbar which enables to select from the given options        
        frame_icon=TopLeftFrame(self, 0, 0, 10, 10, 60, 1024)

        # button for menu page
        icons(frame_icon, "menu.png", 4, 10, None)

        # button for user details
        icons(frame_icon, "user.png", 4, 60, None)

        # button to navigate to the homepage
        icons(frame_icon, "home_alt_fill.jpg", 4, 110, lambda : controller.show_frame(HomePage))

        # button to navigate to the Upload page
        buttons(frame_icon, "Upload.png", 4, 160, 23, 23, lambda : controller.show_frame(Upload))

        # button for settings page 
        icons(frame_icon, "settings.png", 4, 210, None)
                    
        # the default horizontal toolbar at the top
        frame_top=TopLeftFrame(self, 60, 0, 10, 0, 1220, 55)
        
        # button for showing notifications
        icons(frame_top, "notification.png", 1160, 15, None)

        # button for showing information about the page
        icons(frame_top, "info.png", 1110, 15, None)
        
        # button to navigate to the previous page        
        icons(frame_top, "back.png", 10, 15, lambda: controller.back_frame(self))        
        
        # displaying course details
        
        college_name="College :  "+self.values['college']
        course_name="Course :  "+self.values['course']
        part_dict={1:'Second Year', 2:'Third Year', 3:'Passout'}
        year_he="Year     :  "+str(self.values['year'])+" , "+str(part_dict[self.values['part']])
        no_students="Students :  "+str(len(cr['Stu_W']))
        er_students="Students with ER: "+str(len(cr['ER']))
        
        tk.Label(self, text=college_name, bg="#F6F8FB", font=("Helvetica", 16)).place(x=80, y=350)
        tk.Label(self, text=course_name, bg="#F6F8FB", font=("Helvetica", 16)).place(x=80, y=385)
        tk.Label(self, text=year_he, bg="#F6F8FB", font=("Helvetica", 16)).place(x=81, y=420)
        tk.Label(self, text=no_students, bg="#F6F8FB", font=("Helvetica", 16)).place(x=80, y=455)
        tk.Label(self, text=er_students, bg="#F6F8FB", font=("Helvetica", 16)).place(x=80, y=490)
        
        # a frame to show plots for the year wise plots for 
        # subject wise analysis
        chart_frame=tk.LabelFrame(self, width=935, height=270, bg="#F6F8FB")
                    
        # function for plots of subject wise analysis on the frame
        
        def year_plots(sem, x1, y1):
            ax1=cr["Sub_W"][sem].figure    # year 1 plot
            canvas = FigureCanvasTkAgg(ax1, master=chart_frame)  
            canvas.draw() 
            canvas.get_tk_widget().place(x=x1, y=y1)
            
        year_plots(1, 0,0)      # year 1 plot
        year_plots(2, 300,0)    # year 2 plot
        year_plots(3, 600, 0)  # year 3 plot 
        chart_frame.place(x=336, y=55)    # placing the frame at its place     
        
        # toplevel windows to show the year wise subject analysis plots
        # to get a clearer view of each subject
        def Toplevel1(hehe):
            """
            Parameters
            ----------
            hehe : int
                   year number
            Returns : None.
            """
            
            top1=tk.Toplevel()
            title=str()
            if hehe==1:
                title="Year 1"
            elif hehe==2:
                title="Year 2"
            elif hehe==3:
                title="Year 3"
            f_title="Subject Wise Aggregate ("+title+")"
            top1.title(f_title)
            top1.geometry("970x650")
            top1.configure(bg='white')
            top1.resizable(0,0)
            
            
            cr["Sub_W"][hehe].figure.set_figheight(8.6)
            cr["Sub_W"][hehe].figure.set_figwidth(13)
            cr["Sub_W"][hehe].figure.set_dpi(75)
            canvas1 = FigureCanvasTkAgg(cr["Sub_W"][hehe].figure, master=top1)   
            canvas1.draw()
            canvas1.get_tk_widget().place(x=0, y=0)
            
            tk.Button(top1, text="Exit", relief=tk.FLAT, width=6, height=1, bg="white", fg="#3333ff", font=("Arial Black", 10),  command=lambda: top1.destroy()).place(x=891,y=15)
            
            top1.mainloop()
            
        #print(cr['ER'])
            
        def Toplevel2(txt):
            """
            Parameters
            ----------
            txt : string
                  denotes for what the plot is for
                  value : "percplot"
            Returns : None
            """
            top1=tk.Toplevel()
            top1.title("Aggregate Marks of Entire Course")
            top1.geometry("970x650")
            top1.configure(bg='white')
            top1.resizable(0,0)
            
            
            cr[txt].set_figheight(8.6)
            cr[txt].set_figwidth(13)
            cr[txt].set_dpi(75)
            canvas1 = FigureCanvasTkAgg(cr[txt], master=top1)   
            canvas1.draw()
            canvas1.get_tk_widget().place(x=0, y=0)
            
            tk.Button(top1, text="Exit", relief=tk.FLAT, width=6, height=1, bg="white", fg="#3333ff", font=("Arial Black", 10),  command=lambda: top1.destroy()).place(x=891,y=15)
            
            top1.mainloop()
            
        def Toplevel3(txt):
            """
            Parameters
            ----------
            txt : string
                  denotes for what the plot is for
                  value : "Sem_W"
            Returns : None
            """
            top1=tk.Toplevel()
            top1.title("Semester Wise Aggregate")
            top1.geometry("970x650")
            top1.configure(bg='white')
            top1.resizable(0,0)
            
            cr[txt].figure.set_figheight(8.6)
            cr[txt].figure.set_figwidth(13)
            cr[txt].figure.set_dpi(75)
            canvas1 = FigureCanvasTkAgg(cr[txt].figure, master=top1)   
            canvas1.draw()
            canvas1.get_tk_widget().place(x=0, y=0)
            
            tk.Button(top1, text="Exit", relief=tk.FLAT, width=6, height=1, bg="white", fg="#3333ff", font=("Arial Black", 10),  command=lambda: top1.destroy()).place(x=891,y=15)
            
            top1.mainloop()
          
        # displaying the pie plot on the page    
        chart_frame1=tk.LabelFrame(self, width=273, height=270, bg="white")
        ax4=cr["percplot"]
        canvas4 = FigureCanvasTkAgg(ax4, master=chart_frame1)  
        canvas4.draw() 
        canvas4.get_tk_widget().place(x=15, y=14)
        chart_frame1.place(x=60, y=55)
        
        icons(chart_frame1, "plus_square.png", 243, 5, lambda: Toplevel2("percplot"))
        icons(chart_frame, "plus_square.png", 305, 5, lambda: Toplevel1(1))
        icons(chart_frame, "plus_square.png", 605, 5, lambda: Toplevel1(2))
        icons(chart_frame, "plus_square.png", 905, 5, lambda: Toplevel1(3))
        
        # displaying bar chart for subject wise aggregate
        chart_frame2=tk.LabelFrame(self, width=320, height=223, bg="white")
        ax5=cr["Sem_W"].figure
        canvas5 = FigureCanvasTkAgg(ax5, master=chart_frame2)  
        canvas5.draw() 
        canvas5.get_tk_widget().place(x=0, y=0)
        chart_frame2.place(x=951, y=326)
        
        # expand button for chart
        icons(chart_frame2, "plus_square.png", 285, 5, lambda: Toplevel3("Sem_W"))
        
        # creating a scrolled frame
        sf1 = ScrolledFrame(self, width=1193, height=130, bg="#F6F8FB")
        sf1.place(x=60, y=550)
        
        # binding scrollbars to the scrolledframe
        sf1.bind_arrow_keys(self)
        sf1.bind_scroll_wheel(self)
        
        # creating a frame in the scrolled frame
        left_frame = sf1.display_widget(tk.Frame)
        
        """
        Creating the table of students of the selected 
        course of the selected college sorted in ascending
        order of their overall performance
        
        We have set roll numbers in each row as buttons
        so as to get the result data of that particular 
        student directly.
        """
        
        # setting column names
        col_names=cr['Stu_W'].columns
        col_length=len(cr['Stu_W'].columns)
        for i,j in zip(col_names, range(col_length)):
            if i=='Name':
                tk.Label(left_frame, text=i, bg="#F6F8FB", bd=2, padx=30, pady=3, width=20, font=("Helvetica", 11), relief=tk.GROOVE).grid(row=2, column=j, pady=1, padx=1)
            else:
                tk.Label(left_frame, text=i, bg="#F6F8FB", bd=2, padx=30, pady=3, width=14, font=("Helvetica", 11), relief=tk.GROOVE).grid(row=2, column=j, pady=1, padx=1)
            
        # creating roll no column and setting each roll no as button
        len_roll=len(cr['Stu_W']['Roll No'])
        list_roll=cr['Stu_W']['Roll No']
        for i,j in zip(list_roll, range(1,len_roll+1)):
            butt=tk.Button(left_frame, text=i, bg="white", relief=tk.FLAT, width=25, command=lambda i=i: self.callback(i))
            butt.grid(row=j+2, column=0, pady=1, padx=1)
            
        # function for other columns
        def table(cat, wid, col):
            """
            Parameters----------
            cat : str
                column name
            wid : int
                width of label
            col : int
                

            Returns : None
            """
            length=len(cr['Stu_W'][cat])
            list1=cr['Stu_W'][cat];
            for i,j in zip(list1, range(1,length+1)):
                if cat=='Name':
                    i=i.split('\n')[0] if '\n' in i else i.split(' ')[0:2]
                tk.Label(left_frame, text=i, bg="white", pady=4, relief=tk.FLAT, width=wid, font=('Helvetica', 8)).grid(row=j+2, column=col, pady=1, padx=1)
        
        table("Name", 40, 1)
        table("FirstYear", 30, 2)
        table("SecondYear", 30, 3)
        table("ThirdYear", 30, 4)
        table("Total", 30, 5)
                
    def callback(self,i):
        """
        Parameters
        ----------
        i : int
            roll no of student
        Returns : None.
        """
        self.values['student']=i
        self.controller.showStudent(args=self.values, cont=StudentPage)
   
"""
The page showing the list of rankers in top 100
of the selected course, in selected college.
"""         
class  Ranking(tk.Frame):
    
    def __init__(self, parent, controller, args=None):
        """
        Parameters
        ----------
        parent : object
                 The parent window for the frame
        controller : object
                     Enables navigation between pages
        args : object, optional
               Values to be taken from one class to another

        Returns : None
        """
        
        # initialising the frame 
        tk.Frame.__init__(self, parent)
        self.controller=controller
        self.parent=parent
        self.configure(bg='#F6F8FB')
        
        # creating a dictionary to store values
        self.c_values=dict()
        # setting default values in the dictionary
        self.c_values['campus']='All'
        self.c_values['college']=None
        self.c_values.update(args)
        
        # background image        
        bgimage("Images/Ranking_Page.jpg", 1210, 700, self, 60, 55)
        
        # the default vertical toolbar which enables to select from the given options       
        frame_icon=tk.LabelFrame(self, padx=10, pady=10, width=60, height=1024, bg='white', bd=0)
        frame_icon.place(x=0, y=0)
        
        # the default vertical toolbar which enables to select from the given options        
        frame_icon=TopLeftFrame(self, 0, 0, 10, 10, 60, 1024)

        # button for menu page
        icons(frame_icon, "menu.png", 4, 10, None)

        # button for user details
        icons(frame_icon, "user.png", 4, 60, None)

        # button to navigate to the homepage
        icons(frame_icon, "home_alt_fill.jpg", 4, 110, lambda : controller.show_frame(HomePage))

        # button to navigate to the Upload page
        buttons(frame_icon, "Upload.png", 4, 160, 23, 23, lambda : controller.show_frame(Upload))

        # button for settings page 
        icons(frame_icon, "settings.png", 4, 210, None)
                    
        # the default horizontal toolbar at the top
        frame_top=TopLeftFrame(self, 60, 0, 10, 0, 1220, 55)
        
        # button for showing notifications
        icons(frame_top, "notification.png", 1160, 15, None)

        # button for showing information about the page
        icons(frame_top, "info.png", 1110, 15, None)
        
        # button to navigate to the previous page        
        icons(frame_top, "back.png", 10, 15, lambda: controller.back_frame(self))
        
        # getting the list of colleges
        college_list=getdetails(self.c_values['course'], self.c_values['year'], self.c_values["part"],"")
        college_list.insert(0, None) 
        
        # combobox to select the college name to show the ranking of students
        # in top 100
        self.college_combo=combobox(self, college_list, "Choose College", 120, 80, self.callback1)

        # combobox to select the campus in which to show the ranking
        self.campus_combo=combobox(self, ['South','North','All'], "Choose Campus", 530, 80, self.callback2)

        # button to show the ranking
        buttons(self, "go_button.png", 940, 79, 100, 36, lambda:(controller.showStudent(self.c_values, Ranking), controller.back_frame(self)))

        # getting the ranking data and setting the 
        # no. of rows and columns to be shown 
        # in the table
        
        data=get_rank(self.c_values["course"],self.c_values['year'],3,self.c_values['college'], self.c_values['campus'])
        
        if data["Fail"]==True:
            top1=tk.Toplevel()
            top1.title("Error")
            top1.geometry("700x500")
            top1.configure(bg='white')
            top1.resizable(0,0)
            
            tk.Label(top1, text="FileError:", bg="white", font=("Helvetica", 14)).place(x=20, y=50)
            tk.Label(top1, text=data["FileError"], bg="white", font=("Helvetica", 14)).place(x=20, y=90)
            tk.Label(top1, text="Check and correct the uploaded result file in the path.", bg="white", font=("Helvetica", 14)).place(x=20, y=130)

        
        rank_data=data["df"]
        ro=len(rank_data)
        colum=len(rank_data.columns)+2
        
        # creating a scrolled frame to show the ranking
        sf1 = ScrolledFrame(self, width=690, height=500)
        sf1.place(x=120, y=170)
        
        # creating the scroll bars
        sf1.bind_arrow_keys(self)
        sf1.bind_scroll_wheel(self)
        
        # creating the frame of the scrolled frame
        sf1_frame = sf1.display_widget(tk.Frame)
        sf1_frame.configure(bg='#8DB6BA')
        
        # setting column headers
        tk.Label(sf1_frame, text='College', width=25, anchor='w', bd=5, fg="white", bg='#4d79ff', font=('Helvetica', 10)).grid(row=0, column=0, padx=1, pady=1)
        tk.Label(sf1_frame, text='Roll No.', width=15, anchor='w', fg="white", bg='#4d79ff', bd=5, font=('Helvetica', 10)).grid(row=0, column=2, padx=1, pady=1)
        tk.Label(sf1_frame, text='Name', width=20, anchor='w', fg="white", bg='#4d79ff', bd=5, font=('Helvetica', 10), padx=1).grid(row=0, column=3, padx=1, pady=1)
        tk.Label(sf1_frame, text='Gr.CGPA', width=10, anchor='w', fg="white", bg='#4d79ff', bd=5, font=('Helvetica', 10)).grid(row=0, column=5, padx=1, pady=1)
        tk.Label(sf1_frame, text='Rank',width=7, anchor='w', bd=5, fg="white", bg='#4d79ff', font=('Helvetica', 10)).grid(row=0, column=6, padx=1, pady=1)
        
        
        # setting the row data
        for index, row in rank_data.iterrows():
            base = rank_data.index.get_loc(index)
            name=row[1]
            tk.Label(sf1_frame, text=index[0], width=25, anchor='w', bg="#F6F8FB", bd=5, font=('Helvetica', 10)).grid(row=base+1, column=0, padx=1, pady=1)
            tk.Label(sf1_frame, text=row[0], width=15, anchor='w', bg="#F6F8FB", bd=5, font=('Helvetica', 10)).grid(row=base+1, column=2, padx=1, pady=1)
            if '\n' in name:
                tk.Label(sf1_frame, text=row[1].split('\n')[0], width=20, anchor='w', bg="#F6F8FB", bd=5, font=('Helvetica', 10)).grid(row=base+1, column=3, padx=1, pady=1)
            else:
                tk.Label(sf1_frame, text=row[1].split(' ')[0:2], width=20, anchor='w', bg="#F6F8FB", bd=5, font=('Helvetica', 10)).grid(row=base+1, column=3, padx=1, pady=1)
            tk.Label(sf1_frame, text="{:.3f}".format(float(row[2])), bg="#F6F8FB", width=10, anchor='w', bd=5, font=('Helvetica', 10)).grid(row=base+1, column=5, padx=1, pady=1)
            tk.Label(sf1_frame, text=row[3],width=7, anchor='w', bg="#F6F8FB", bd=5, font=('Helvetica', 10)).grid(row=base+1, column=6, padx=1, pady=1)
    
        # calling and setting plots
        # pie plot to show no. of students of selected
        # college in top 100
        # and bar plot to show the distribution of colleges
        # in top 100 ranking
        if self.c_values['college']==None:
            rank_plot1=get_rank(self.c_values['course'],self.c_values["year"],3,self.c_values["college"], self.c_values['campus'])["plot"].figure
            rank_frame1=tk.LabelFrame(self, padx=0, pady=0, width=380, height=430, bg='white', bd=2)
            canvas1 = FigureCanvasTkAgg(rank_plot1, master=rank_frame1)   
            canvas1.draw()
            canvas1.get_tk_widget().place(x=0, y=0)
            rank_frame1.place(x=860, y=250)
        else:
            rank_plot2=get_rank(self.c_values['course'],self.c_values["year"],3,self.c_values["college"], self.c_values['campus'])["plot"]
            rank_frame2=tk.LabelFrame(self, padx=0, pady=0, width=380, height=300, bg='white', bd=2)
            canvas1 = FigureCanvasTkAgg(rank_plot2, master=rank_frame2)   
            canvas1.draw()
            canvas1.get_tk_widget().place(x=0, y=0)
            tk.Label(rank_frame2, text="No. of Students in top 100", bg="white", font=('Helvetica', 12)).place(x=97, y=250)
            tk.Label(rank_frame2, text=self.c_values["college"], bg="white", font=('Helvetica', 12)).place(x=10, y=10)
            rank_frame2.place(x=860, y=270)
            
            
    def callback1(self, *args):
        self.college=self.college_combo.get()   
        if self.college=='None':
            self.college=None
        self.c_values['college']=self.college
        
    def callback2(self, *args):
        self.campus=self.campus_combo.get()
        self.c_values['campus']=self.campus
    
"""
Extra Page for some additional information.
"""
class Upload(tk.Frame): 
    def __init__(self, parent, controller): 
        tk.Frame.__init__(self, parent) 
        self.parent = parent
        self.main = self.master
        self.controller=controller
        self.configure(bg="#F6F8FB")
        
        self.c_values=dict()
                
        bgimage("Images/Frame 1.png", 500, 600, self, 70, 40)
        
        # the default vertical toolbar which enables to select from the given options       
        frame_icon=tk.LabelFrame(self, padx=10, pady=10, width=60, height=1024, bg='white', bd=0)
        frame_icon.place(x=0, y=0)
        
        # the default vertical toolbar which enables to select from the given options        
        frame_icon=TopLeftFrame(self, 0, 0, 10, 10, 60, 1024)

        # button for menu page
        icons(frame_icon, "menu.png", 4, 10, None)

        # button for user details
        icons(frame_icon, "user.png", 4, 60, None)
        
        def res(i):
            if i=="home":
                tk.messagebox.showinfo("Restart Required","The application will restart")
            python=sys.executable
            os.execl(python, python, * sys.argv)

        # button to navigate to the homepage
        icons(frame_icon, "home_alt_fill.jpg", 4, 110, lambda: res("home"))

        # button to navigate to the Upload page
        buttons(frame_icon, "Upload.png", 4, 160, 23, 23, lambda : controller.show_frame(Upload))

        # button for settings page 
        icons(frame_icon, "settings.png", 4, 210, None)
                    
        # the default horizontal toolbar at the top
        frame_top=TopLeftFrame(self, 60, 0, 10, 0, 1220, 55)
        
        # button for showing notifications
        icons(frame_top, "notification.png", 1160, 15, None)

        # button for showing information about the page
        icons(frame_top, "info.png", 1110, 15, None)
        
        # button to navigate to the previous page        
        icons(frame_top, "back.png", 10, 15, res)
        
        tk.Label(self, text="Instructions - ", font=("Helvetica",16), bg="#F6F8FB").place(x=600,y=80)
        tk.Label(self, text="The file must contain following columns: ", font=("Helvetica",14), bg="#F6F8FB").place(x=620,y=110)
        
        ls="Roll No	SEM	Name	Sub	GR	GP	CRP	Sub.1	GR.1	GP.1	CRP.1	Sub.2	GR.2	GP.2	CRP.2	Sub.3	GR.3	GP.3	CRP.3	Sub.4	GR.4	GP.4	CRP.4	TOT CR	TOT CRP	SGPA	CGPA	Result	GR.CGPA	DIV"
        a=ls.split("\t")
        def instr(start,stop,x1):
            for i in range(0, len(a[start:stop])):
                tk.Label(self, text=a[i+start], font=("Helvetica",10), bg="#F6F8FB").place(x=x1,y=150+20*i)
        
        instr(0,10, 623)
        instr(10,20, 723)
        instr(20,30, 823)
        
        def sample():
            top=tk.Toplevel()
            top.title("Sample Form of File")
            top.configure(bg='white')
            top.resizable(0,0)
            top.geometry("1250x600")
            
            bgimage("Screenshots/Sample.png", 1250, 800, top, 0, 0)
            
            top.mainloop()
            
        
        tk.Label(self, text="To see a sample file format:", font=("Helvetica",14),bg="#F6F8FB").place(x=623, y=370)
        tk.Button(self, text="Click here", font=("Helvetica",10),bg="#F6F8FB", fg="Blue", relief=tk.FLAT, command=sample).place(x=623, y=395)
        
        tk.Label(self, text="Upload will take some time depending on the file size.", font=("Helvetica",14),bg="#F6F8FB").place(x=623, y=450)

        
        tk.Label(self, text='Upload Files', bg='white', font=("Roboto", 25), fg="black").place(x=135, y=110)
        
        self.v1 = tk.StringVar()
        tk.Label(self, text='Enter Course Name', bg='white', font=("Roboto", 10), fg="black").place(x=135, y=175)
        
        entry1 = tk.Entry(self, textvariable=self.v1, width=30, relief=tk.GROOVE, bg="white", font=("Helvetica",16), exportselection=0).place(x=135, y=200)        
        
        self.year_combo=combobox(self, ['2019'], "Choose Year", 135, 260, self.callback1)

        self.part_combo=combobox(self, ['First Year', 'Second Year','Passout'], "Choose Course Part", 135, 320, self.callback2)

        self.v = tk.StringVar()
        
        buttons(self, "browse_button.png", 240, 380, 150, 50, lambda: self.browse(self.v))

        entry = tk.Entry(self, textvariable=self.v, width=40, relief=tk.FLAT, bg="white", font=("Helvetica",12), exportselection=0).place(x=135, y=440)
        
        buttons(self, "upload_butt.png", 240, 500, 150, 50, self.upload)
                
    def browse(self,v):
        csv_file_path = askopenfilename()
        self.v.set(csv_file_path)
        self.c_values['course']=self.v1.get()
        
    def upload(self):
        self.c_values['course']=self.v1.get()
        aa=pdfupload(self.c_values['course'], self.c_values['year'], self.c_values['part'], self.v.get())
        dict1={list(aa.keys())[0]:list(aa.values())[0], list(aa.keys())[-1]:list(aa.values())[-1]}
        for i in dict1:
                top1=tk.Toplevel()
                top1.geometry("700x500")
                top1.configure(bg='white')
                top1.resizable(0,0)
                
                if 'Error' not in i:
                    top1.title("Error!")
                    tk.Label(top1, text=i, bg='white', font=("Helvetica", 14)).place(x=20, y=50)
                    tk.Label(top1, text="Following columns are missing:", bg='white', font=("Helvetica", 13)).place(x=25, y=100)
                    for j in aa[i]:
                        tk.Label(top1, text=j, bg='white', font=("Helvetica", 12)).place(x=20, y=150+aa[i].index(j)*40)
                    tk.Label(top1, text="Check the entire file for missing columns.", bg='white', font=("Helvetica", 13)).place(x=300, y=100)

                else:
                    top1.title("File Creation Error")
                    tk.Label(top1, text="Rank file cannot be created", bg='white', font=("Helvetica", 14)).place(x=20, y=50)
                    for j in aa[i]:
                        txt=str(int(aa[i].index(j)+1))+". "+j
                        tk.Label(top1, text=txt, bg='white', font=("Helvetica", 13)).place(x=20, y=50+int(aa[i].index(j)+1)*50)
                        
  
        tk.messagebox.showinfo("Success","File uploaded successfully with errors! It needs manual corrrection. Search for the uploaded excel file in folder of the application and edit it. Check the opened top-level windows of the application for possible errors. Restart required at this stage.")
        python=sys.executable
        os.execl(python, python, * sys.argv)
                
            
        
    def callback1(self, *args): 
        self.year=self.year_combo.get()
        self.c_values['year']=self.year
        
    def callback2(self, *args):
        self.part=self.part_combo.get()
        part_dict={'First Year':1, 'Second Year':2, 'Passout':3}
        self.c_values['part']=part_dict[self.part]
        
def main():
    app = tkinterApp()
    app.mainloop()
    
if __name__=='__main__':
    main()