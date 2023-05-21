from fpdf import FPDF
from fpdf.enums import XPos, YPos
import pandas as pd
from PIL import Image


def analysis_pdf(file, fname):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("helvetica", "", 10)

    data = pd.read_excel(file)
    title = data.iloc[0:1,:].values[0]
    name_array = data.iloc[:, 1:2].values
    count = 0
    absent = 0
    failed = 0
    female_count = 0
    male_count = 0
    female_appeared = 0
    male_appeared = 0
    female_failed = 0
    male_failed = 0

    for sgp in range(0, len(title)):
        if(title[sgp] == 'SGPI'):
            sgp_col = data.iloc[:, sgp:sgp+1].values

    for name in range(0, len(name_array)):
        if(name_array[name] != 'Name'):
            count += 1
        name_div = name_array[name][0].split(" ")
        if (name_div[0] == '/'):
            female_count += 1
            if sgp_col[name] == 'Ab':
                pass
            else:
                female_appeared += 1
                if sgp_col[name] == 'F':
                    female_failed += 1
        else:
            male_count+=1
            if sgp_col[name] == 'Ab':
                pass
            else:
                male_appeared += 1
                if sgp_col[name] == 'F':
                    male_failed += 1

    for absent_stud in range(0, len(sgp_col)):
            if(sgp_col[absent_stud] == 'Ab'):
                absent += 1
            if(sgp_col[absent_stud] == 'F'):
                failed += 1

    student_appeared = count - absent
    passed_student = student_appeared - failed

    passing_percent_cal = (passed_student/student_appeared) * 100
    passing_percent = "%.2f" %passing_percent_cal + "%"

    male_passed = male_appeared - male_failed
    female_passed = female_appeared - female_failed

    female_passing_percentage_cal = (female_passed/female_appeared)*100
    female_passing_percentage = "%.2f" %female_passing_percentage_cal + "%"

    male_passing_percentage_cal = (male_passed/male_appeared)*100
    male_passing_percentage = "%.2f" %male_passing_percentage_cal + "%"
    

    logo = Image.open('logo.jpeg')
    logo_rgb = logo.convert('RGB')
    logo_display =  logo_rgb.resize((550, 100), resample=Image.NEAREST)
    pdf.image(logo_display, y=15)


    #Department and year information
    pdf.set_xy(70,50)
    pdf.cell(w=150, h=10, txt="Department of Computer Engineering", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_x(90)
    pdf.cell(w=150, h=10, txt="Result Analysis", new_x=XPos.LMARGIN, new_y=YPos.NEXT)


    #General Analysis
    pdf.set_xy(65 ,100)
    pdf.cell(w=55, h=5, txt = "    Total No. of Students ", border=1)
    pdf.cell(w=15, h=5, txt=str(count), border=1, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_xy(65 ,105)
    pdf.cell(w=55, h=5, txt="    Total Student Appeared", border=1)
    pdf.cell(w=15, h=5, txt=str(student_appeared), border=1, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_xy(65 ,110)
    pdf.cell(w=55, h=5, txt="    Total Student Passed", border=1)
    pdf.cell(w=15, h=5, txt=str(passed_student), border=1, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_xy(65 ,115)
    pdf.cell(w=55, h=5, txt="    Total Student Failed", border=1)
    pdf.cell(w=15, h=5, txt=str(failed), border=1 , new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_xy(65 ,120)
    pdf.cell(w=55, h=5, txt="    Passing Percentage", border=1)
    pdf.cell(w=15, h=5, txt=str(passing_percent), border=1, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_xy(65 ,125)
    pdf.cell(w=55, h=5, txt="    Total Not Appeared Student", border=1)
    pdf.cell(w=15, h=5, txt=str(absent), border=1 , new_x=XPos.LMARGIN, new_y=YPos.NEXT)


    #Female Analysis
    pdf.set_xy(75 ,140)
    pdf.cell(w=35, h=5, txt="      Total Female ", border=1)
    pdf.cell(w=15, h=5, txt=str(female_count), border=1, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_xy(75 ,145)
    pdf.cell(w=35, h=5, txt="    Female Appeared ", border=1)
    pdf.cell(w=15, h=5, txt=str(female_appeared), border=1, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_xy(75 ,150)
    pdf.cell(w=35, h=5, txt="    Female Passed", border=1)
    pdf.cell(w=15, h=5, txt=str(female_passed), border=1, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_xy(75 ,155)
    pdf.cell(w=35, h=5, txt="     Female Failed", border=1)
    pdf.cell(w=15, h=5, txt=str(female_failed), border=1, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_xy(75 ,160)
    pdf.cell(w=35, h=5, txt=" Female Pass %", border=1)
    pdf.cell(w=15, h=5, txt=str(female_passing_percentage), border=1, new_x=XPos.LMARGIN, new_y=YPos.NEXT)


    #Male Analysis
    pdf.set_xy(75, 175)
    pdf.cell(w=35, h=5, txt="      Total Male  ", border=1)
    pdf.cell(w=15, h=5, txt=str(male_count), border=1, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_xy(75 ,180)
    pdf.cell(w=35, h=5, txt="    Male Appeared ", border=1)
    pdf.cell(w=15, h=5, txt=str(male_appeared), border=1, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_xy(75 ,185)
    pdf.cell(w=35, h=5, txt="    Male Passed ", border=1)
    pdf.cell(w=15, h=5, txt=str(male_passed), border=1, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_xy(75 ,190)
    pdf.cell(w=35, h=5, txt="     Male Failed", border=1)
    pdf.cell(w=15, h=5, txt=str(male_failed), border=1, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_xy(75 ,195)
    pdf.cell(w=35, h=5, txt=" Male Pass %", border=1)
    pdf.cell(w=15, h=5, txt=str(male_passing_percentage), border=1, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    

    pdf.set_y(250)
    pdf.set_x(20)
    pdf.cell(w=20, h=10, txt="Result Analysis Incharge")
    pdf.set_x(140)
    pdf.cell(w=20, h=10, txt="Computer Engineering Dept.")

    pdf.output(fname)
