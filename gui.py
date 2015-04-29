__author__ = 'Amit_saroj'
from tkinter import *
import xlwt
import xlrd



# main frame

class Example(Frame):
    def __init__(self, root):
        root.minsize(width=666, height=700)
        Frame.__init__(self, root)
        self.canvas = Canvas(root, borderwidth=0, background="pink")
        self.frame = Frame(root, background="red",)
        self.vsb = Scrollbar(root, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.vsb.set)

        self.vsb.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)
        self.canvas.create_window((10,10), window=self.frame, anchor="nw",
                                  tags="self.frame")

        self.frame.bind("<Configure>", self.OnFrameConfigure)

        self.populate()


    def populate(self):



        def save():
            book_r = xlrd.open_workbook("text.xls")
            first_sheet = book_r.sheet_by_name("University Form")
            row_count = 1

            for row_index in range(first_sheet.nrows):
                value = first_sheet.cell(rowx=row_index, colx=0).value
                if value != "Sr.No." or value != range(1, 200):
                   sheet1.write(row_index + 1, 0, row_count)
            row_count += 1


        def sel():
          selection = str(var.get())
          label_display.config(text=selection)
        def bani():
          selection2 = str(ban.get())
          label_display_2.config(text=selection2)

        def gender():
            selection3 = str(gen.get())
            label_gender.config(text=selection3)

        def cate():
            selection4 = str(cas.get())
            label_cate.config(text=selection4)

        def handi():
            selection5 = str(han.get())
            label_handi.config(text=selection5)

        def nati():
            selection6 = str(nat.get())
            label_Nationality.config(text=selection6)

        def htown():
            str(town.get())
            return
#left
        labelframe_left = LabelFrame(self.frame, text="1.", width=300)
        labelframe_left.pack(fill="both", expand="yes", side=LEFT)
#right
        labelframe_right = LabelFrame(self.frame,text="2")
        labelframe_right.pack(fill="both",expand="yes", side=RIGHT)
#widgit for university applying for
        label_apply = Label(labelframe_left, text="University Applying For", relief=RAISED, bd=4, bg="Blue")
        label_apply.grid(row=1, column=0)
        var = IntVar()
        R1_apply = Radiobutton(labelframe_left, text="Shobhit University,Meerut,Delhi,NCR", variable=var, value=1,
                  command=sel)
        R1_apply.grid(row=2, column=0)
        R2_apply = Radiobutton(labelframe_left, text="Shobhit University, Gangoh,Saharanpur,UP", variable=var, value=2,
                  command=sel)
        R2_apply.grid(row=3, column=0)

        label_display = Label(labelframe_left)
        label_display.grid(row=4)



# widget for program applying for

        label_program = Label(labelframe_left, text="PROGRAM APPLYING FOR", relief=RAISED, bd=4, bg="Blue")
        label_program.grid(row=4, column=0)
        label_program_detail = Label(labelframe_left, text="Program Name")
        label_program_detail.grid(row=5)
        entry_program_apply = Entry(labelframe_left)
        entry_program_apply.grid(row=5, column=1)



# Branch preference

        label_Branch = Label(labelframe_left, text="BRANCH PREFERENCE", relief=RAISED, bd=4, bg="Blue")
        label_Branch.grid(row=6)

        label_Branch_Number = Label(labelframe_left, text="1.")
        label_Branch_Number.grid(row=7)
        entry_pre_branch_1 = Entry(labelframe_left)
        entry_pre_branch_1.grid(row=7, column=1)

        label_Branch_Number = Label(labelframe_left, text="2.")
        label_Branch_Number.grid(row=8)
        entry_pre_branch_2 = Entry(labelframe_left)
        entry_pre_branch_2.grid(row=8, column=1)

        label_Branch_Number = Label(labelframe_left, text="3.")
        label_Branch_Number.grid(row=9)
        entry_pre_branch_3 = Entry(labelframe_left)
        entry_pre_branch_3.grid(row=9, column=1)

        label_Branch_Number = Label(labelframe_left, text="4.")
        label_Branch_Number.grid(row=10)
        entry_pre_branch_4 = Entry(labelframe_left,)
        entry_pre_branch_4.grid(row=10, column=1)

        label_Branch_Number = Label(labelframe_left, text="5.")
        label_Branch_Number.grid(row=11)
        entry_pre_branch_5 = Entry(labelframe_left)
        entry_pre_branch_5.grid(row=11, column=1)

# ADMISSION CRITERIA
        label_admission = Label(labelframe_left, text="ADMISSION CRITERIA", relief=RAISED, bd=4)
        label_admission.grid(row=12)
        ban = IntVar()
        R1_admission = Radiobutton(labelframe_left, text="Marks in Qualifying Exams", variable=ban, value=1,
                  command=bani)
        R1_admission.grid(row=13)

        R2_admission =Radiobutton(labelframe_left, text="GD/PI/Test Conducted by SU", variable=ban, value=2,
                  command=bani)
        R2_admission.grid(row=14)

        R3_admission =Radiobutton(labelframe_left, text="Test Conducted by National/State", variable=ban, value=3,
                  command=bani)
        R3_admission.grid(row=15)
        label_display_2 = Label(labelframe_left)
        label_display_2.grid(row=16)

# PERSONAL INFORMATION

        label = Label(labelframe_left, text="PERSONAL INFORMATION", relief=RAISED, bd=4, bg="Blue" )
        label.grid(row=16)

        label_Student_Name = Label(labelframe_left, text="Student Name:")
        label_Student_Name.grid(row=17)
        entry_stu_name = Entry(labelframe_left)
        entry_stu_name.grid(row=17, column=1)

        label_Mother_Name = Label(labelframe_left, text="Mother;s Name")
        label_Mother_Name.grid(row=18)
        entry_mother_name = Entry(labelframe_left)
        entry_mother_name.grid(row=18, column=1)

        label_Father_Name = Label(labelframe_left, text="Father's Name:")
        label_Father_Name.grid(row=19)
        entry_father_name = Entry(labelframe_left)
        entry_father_name.grid(row=19, column=1)
        label_Guardian_Name = Label(labelframe_left, text="Legal Guardian's Name:")
        label_Guardian_Name.grid(row=20)
        entry_guard_name = Entry(labelframe_left)
        entry_guard_name.grid(row=20, column=1)
        label_relation = Label(labelframe_left, text="relation with guardian if there:")
        label_relation.grid(row=21)
        entry_relation = Entry(labelframe_left)
        entry_relation.grid(row=21, column=1)


        label_DOB = Label(labelframe_left, text="Date of Birth:", bg="Green")
        label_DOB.grid(row=22)
        entry_DOB = Entry(labelframe_left, text="DD/MM/YYY")
        entry_DOB.grid(row=23)

        label = Label(labelframe_left, text="Gender:", bg="Green")
        label.grid(row=24)
        gen = IntVar()
        R1_Gender = Radiobutton(labelframe_left, text="Male", variable=gen, value=1,
                  command=gender)
        R1_Gender.grid(row=25)
        R2_Gender = Radiobutton(labelframe_left, text="Female", variable=gen, value=2,
                                command=gender)
        R2_Gender.grid(row=26)
        label_gender = Label(labelframe_left)
        label_gender.grid(row=27)




        label_category=Label(labelframe_left, text="Category:", bg="Green")
        label_category.grid(row=28)
        cas = IntVar()
        R1_Category = Radiobutton(labelframe_left, text="General", variable=cas, value=1,
                                  command=cate)
        R1_Category.grid(row=29)

        R2_Category = Radiobutton(labelframe_left, text="ST", variable=cas, value=2,
                                  command=cate)
        R2_Category.grid(row=30)

        R3_Category = Radiobutton(labelframe_left, text="SC", variable=cas, value=3,
                                  command=cate)
        R3_Category.grid(row=31)

        R4_Category = Radiobutton(labelframe_left, text="Minority", variable=cas, value=4,
                                  command=cate)
        R4_Category.grid(row=32)

        R5_Category = Radiobutton(labelframe_left, text="OBC", variable=cas, value=5,
                                  command=cate)
        R5_Category.grid(row=33)

        label_cate = Label(labelframe_left)
        label_cate.grid(row=34)


        label = Label(labelframe_left, text="If,other specify")
        label.grid(row=35)
        entry_cas = Entry(labelframe_left)
        entry_cas.grid(row=35, column=1)


        label_handicapped = Label(labelframe_left, text="Physically Handicapped:", bg="Green")
        label_handicapped.grid(row=37)
        han = IntVar()
        R1_handicapped = Radiobutton(labelframe_left, text="Yes", variable=han, value=1,
                                     command=handi)
        R1_handicapped.grid(row=38)

        R2_handicapped = Radiobutton(labelframe_left, text="No", variable=han, value=2,
                                     command=handi)
        R2_handicapped.grid(row=39)

        label_handi=Label(labelframe_left)
        label_handi.grid(row=40)

        label_nation = Label(labelframe_left, text="Nationality:", bg="Green")
        label_nation.grid(row=41)
        nat = IntVar()
        R1_Nationality = Radiobutton(labelframe_left, text="Indian", variable=nat, value=1,
                                     command=nati)
        R1_Nationality.grid(row=42)

        R2_Nationality = Radiobutton(labelframe_left, text="Others", variable=nat, value=2,
                                     command=nati)
        R2_Nationality.grid(row=43)
        label_nat_other = Label(labelframe_left, text="If,other specify")
        label_nat_other.grid(row=44)
        entry_nat = Entry(labelframe_left)
        entry_nat.grid(row=44, column=1)
        label_Nationality = Label(labelframe_left)
        label_Nationality.grid(row=46)


# Contact Details

        label = Label(labelframe_right, text="Contact Details", bd=4, bg="Yellow")
        label.grid(row=0)
        label_correspondence_address = Label(labelframe_right, text="Correspondence Address:", bg="green")
        label_correspondence_address.grid(row=1)
        text_con_address = Text(labelframe_right, height=4)
        text_con_address.grid(row=2)


        label_permanent = Label(labelframe_right, text="Permanent Address", bg="green")
        label_permanent.grid(row=3)
        text_permanent_address = Text(labelframe_right, height=4)
        text_permanent_address.grid(row=4)


        label_local_address=Label(labelframe_right, text="LocalGuardian Address", bg="green")
        label_local_address.grid(row=5)
        text_local_address = Text(labelframe_right, height=4)
        text_local_address.grid(row=6)

        label_tele_detail = Label(labelframe_right, text="Telephone number(with STD code):-", bg="green", bd="4", relief=GROOVE)
        label_tele_detail.grid(row=7)
        entry_tele_number = Entry(labelframe_right)
        entry_tele_number.grid(row=7, column=1)

        label_mobile_detail = Label(labelframe_right, text="Mobile number(with country code):-", bg="green", bd="4", relief=GROOVE)
        label_mobile_detail.grid(row=8)
        entry_mobile_number = Entry(labelframe_right)
        entry_mobile_number.grid(row=8, column=1)

        label_pmob_detail = Label(labelframe_right, text="Mobile number of Parent/Guardian(with country code):-", bg="green", bd="4", relief=GROOVE)
        label_pmob_detail.grid(row=9)
        entry_pmob_number = Entry(labelframe_right)
        entry_pmob_number.grid(row=9, column=1)

        label_email_id = Label(labelframe_right, text= "Email_id:-", bg="green", bd=4, relief=GROOVE)
        label_email_id.grid(row=10)
        entry_email_id = Entry(labelframe_right)
        entry_email_id.grid(row=10, column=1)


        label_hometown = Label(labelframe_right, text="HomeTown", relief=GROOVE, bd=4, bg="green")
        label_hometown.grid(row=11)

        town = IntVar()
        R1_homeTown = Radiobutton(labelframe_right, text="Rural", variable=town, value=1, command=htown)
        R1_homeTown.grid(row=12)

        R2_homeTown = Radiobutton(labelframe_right, text="Urban(Town)", variable=town, value=2, command=htown)
        R2_homeTown.grid(row=13)

        R3_homeTown = Radiobutton(labelframe_right, text="Urban(Metrotown)", variable=town, value=3, command=htown)
        R3_homeTown.grid(row=14)



  # Academic Information
        label_Academic_info = Label(labelframe_right, text="Academic Information", bg="Yellow", bd="4", relief=GROOVE)
        label_Academic_info.grid(row=15)

        label_Qualifying_info = Label(labelframe_right, text="Qualifying Exams", bg="green", bd="4", relief=GROOVE)
        label_Qualifying_info.grid(row=16)

        Check_qualify_1 = IntVar()
        Check_qualify_2 = IntVar()
        Check_qualify_3 = IntVar()
        Check_qualify_4 = IntVar()
        Check_qualify_5 = IntVar()
        C1_Qualify= Checkbutton(labelframe_right, text = "10th", variable = Check_qualify_1)
        C2_Qualify = Checkbutton(labelframe_right, text = "10+2(12th)", variable = Check_qualify_2)
        C3_Qualify = Checkbutton(labelframe_right, text = "3 yrs. Diploma after 10th", variable = Check_qualify_3)
        C4_Qualify = Checkbutton(labelframe_right, text = "Graduation", variable = Check_qualify_4)
        C5_Qualify = Checkbutton(labelframe_right, text = "Post Graduation", variable = Check_qualify_5)


        C1_Qualify.grid(row=17)
        C2_Qualify.grid(row=18)
        C3_Qualify.grid(row=19)
        C4_Qualify.grid(row=20)
        C5_Qualify.grid(row=21)



        master_Button = Button(labelframe_right, text="Save", command=save)
        master_Button.grid(row=23, column=2)



    def OnFrameConfigure(self, event):

        '''Reset the scroll region to encompass the inner frame'''
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

if __name__ == "__main__":
	
    root = Tk()
    root.title("University Form")
    root.geometry("1250x680+60+0")
    title_cols = 1






# write in excel file
    book_w = xlwt.Workbook()
    sheet1 = book_w.add_sheet("University Form", cell_overwrite_ok=TRUE)
    style = xlwt.easyxf('font: bold 1')
# detail_list = ['entry_stu_name', 'entry_mother_name', 'entry_father_name', 'entry_guard_name', 'var', 'entry_program_apply', 'entry_pre_branch_1', 'entry_DOB', 'gen', 'cas', 'han', 'nat', 'ban']

    sheet1_list = ['Name', "Mother's Name", "Father's Name", "Guardians's Name and relation", 'Location', 'Program', 'Branch', 'DOB', 'Gender', 'Category', 'Physiscally handicapped', 'Nationality', 'Adminission Criteria']
    sheet1.write(0, 0, 'Sr.No.') # row, column, value
    for n in sheet1_list:
        sheet1.write(0, title_cols, n, style)
        sheet1.col(title_cols).width = 256 * (len(n) + 2)
        title_cols += 1



#read in excel file


    book_w.save("text.xls")
    Example(root).pack(side="top", fill="both", expand=True)

    root.mainloop()
