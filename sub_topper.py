from fpdf import FPDF
from fpdf.enums import XPos, YPos
import pandas as pd
from PIL import Image


def subject_topper_pdf(file, fname):
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
    passing_marks = 35
    subject_array = ['Subject']
    subject_student_total = ['Total Students']
    subject_student_present = [' Present']
    subject_student_passed = ['Student Passed']
    subject_student_failed = ['Student Failed']
    subject_student_passing_percentage = ['Passing %']
    top_marks_array = ['Highest Marks']
    topper_name_array = ['Topper']
    subject_teacher = ['Sub. Teacher']
    sub_topper_name_array = []

    def analysis_table(arr):
        for cell in range(0, len(arr)):
            if cell == len(arr) - 1:
                pdf.cell(w=23, h=5, txt=str(arr[cell]), border=1, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
            elif cell == 0:
                pdf.cell(w=25, h=5, txt=str(arr[cell]), border=1)
            else:
                pdf.cell(w=25.3, h=5, txt=str(arr[cell]), border=1)



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





    #Finding no. of subject in excel file
    for subject in range(0, len(title)):
        temp_subject = title[subject]
        temp_subject_arr = temp_subject.split(' ')
        if(temp_subject_arr[-1] == 'NT'):
            temp_subject = 'nt'
        if(temp_subject.lower() == 'seat no.' or temp_subject.lower() == 'name' or temp_subject.lower() == 'sgpi' or temp_subject.lower() == 'mini project' or temp_subject.lower() == 'nt'):
            pass
        else:
            subject_array.append(temp_subject.upper())
    



    #Finding total number of student appeared subject wise
    for subject in range(0, len(title)):
        temp_subject = title[subject]
        temp_subject_arr = temp_subject.split(' ')
        if(temp_subject_arr[-1] == 'NT'):
            temp_subject = 'nt'
        if(temp_subject.lower() == 'seat no.' or temp_subject.lower() == 'name' or temp_subject.lower() == 'sgpi' or temp_subject.lower() == 'mini project' or temp_subject.lower() == 'nt'):
            pass
        else:
            subject_col = data.iloc[1:, subject:subject+1].values
            student_count = 0
            for student in range(0, len(subject_col)):
                if (subject_col[student] == 'NC'):
                    pass
                else:
                    student_count += 1
            subject_student_total.append(student_count)



    #Finding present students subject wise
    for subject in range(0, len(title)):
        temp_subject = title[subject]
        temp_subject_arr = temp_subject.split(' ')
        if(temp_subject_arr[-1] == 'NT'):
            temp_subject = 'nt'
        if(temp_subject.lower() == 'seat no.' or temp_subject.lower() == 'name' or temp_subject.lower() == 'sgpi' or temp_subject.lower() == 'mini project' or temp_subject.lower() == 'nt'):
            pass
        else:
            subject_col = data.iloc[1:, subject:subject+1].values
            present_count = 0
            for student in range(0, len(subject_col)):
                if (subject_col[student] == 'NC' or subject_col[student] == 'AB'):
                    pass
                else:
                    present_count += 1
                    
            subject_student_present.append(present_count)


    #Finding students passed exams
    for subject in range(0, len(title)):
        temp_subject = title[subject]
        temp_subject_arr = temp_subject.split(' ')
        if(temp_subject_arr[-1] == 'NT'):
            temp_subject = 'nt'
        if(temp_subject.lower() == 'seat no.' or temp_subject.lower() == 'name' or temp_subject.lower() == 'sgpi' or temp_subject.lower() == 'mini project' or temp_subject.lower() == 'nt'):
            pass
        else:
            subject_col = data.iloc[1:, subject:subject+1].values
            passed_count = 0
            for student in range(0, len(subject_col)):
                if(subject_col[student] == 'NC' or subject_col[student] == 'AB'):
                    pass
                else:
                    if (int(subject_col[student]) >= passing_marks):
                        passed_count += 1
            subject_student_passed.append(passed_count)
    

    


    #Finding students students failed subject wise
    for failed_student in range(1, len(subject_student_passed)):
        fail = subject_student_total[failed_student] - subject_student_passed[failed_student]
        subject_student_failed.append(fail)
    


    #Finding Passing %
    for pass_percent in range(1, len(subject_student_passed)):
        percent_passing = (subject_student_passed[pass_percent] / subject_student_total[pass_percent]) * 100
        percent_passing_adjust = '%.2f' %percent_passing + "%"
        subject_student_passing_percentage.append(percent_passing_adjust)
    

    
    #Finding Highest marking
    for subject in range(0, len(title)):
        temp_subject = title[subject]
        temp_subject_arr = temp_subject.split(' ')
        if(temp_subject_arr[-1] == 'NT'):
            temp_subject = 'nt'
        if(temp_subject.lower() == 'seat no.' or temp_subject.lower() == 'name' or temp_subject.lower() == 'sgpi' or temp_subject.lower() == 'mini project' or temp_subject.lower() == 'nt'):
            pass
        else:
            subject_col = data.iloc[1:, subject:subject+1].values
            top_marks = 0
            for student in range(0, len(subject_col)):
                if(subject_col[student] == 'NC' or subject_col[student] == 'AB'):
                    pass
                else:
                    if (int(subject_col[student]) == top_marks):
                        sub_topper_name_array.append(name_array[student+1][0])
                    if (int(subject_col[student]) > top_marks):
                        sub_topper_name_array = []
                        top_marks = subject_col[student][0]
                        sub_topper_name_array.append(name_array[student+1][0])
                    
                    
            top_marks_array.append(top_marks)
            topper_name_array.append(sub_topper_name_array)


    #Finding Subject teacher array
    for _ in range(1, len(subject_student_passing_percentage)):
        subject_teacher.append('')

    
                
            
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

    pdf.set_xy(70,50)
    pdf.cell(w=150, h=10, txt="Department of Computer Engineering", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_x(90)
    pdf.cell(w=150, h=10, txt="Subject Wise Toppers", new_x=XPos.LMARGIN, new_y=YPos.NEXT)



    analysis_table(subject_array)
    analysis_table(subject_student_total)    
    analysis_table(subject_student_present)
    analysis_table(subject_student_passed)
    analysis_table(subject_student_failed)
    analysis_table(subject_student_passing_percentage)
    analysis_table(top_marks_array)

    for top in range(0, len(topper_name_array)):
        pdf.set_font_size(7)
        top_name = ''
        pdf.set_y(105)
        set_x = (10 + 25*top)
        if top == 1:
            set_x = (10+ 25*top)
        pdf.set_x(set_x)
        if top == 0:
            pdf.multi_cell(w=25, h=50, txt=str(topper_name_array[top]), align='J', border=1)
        else:
            for name_str in range(0, len(topper_name_array[top])):
                if name_str == 3:
                    top_name += 'FOR OTHER TOPPERS PLEASE REFER PDF'
                    break
                top_name += str(name_str+1) 
                top_name += topper_name_array[top][name_str]
                top_name += '\n'
            pdf.multi_cell(w=25, h=5, txt=top_name, align='J', border=1)

        
            

    pdf.set_xy(10, 190)

    for cell in range(0, len(subject_teacher)):
            if cell == len(subject_teacher) - 1:
                pdf.cell(w=23, h=15, txt=str(subject_teacher[cell]), border=1, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
            elif cell == 0:
                pdf.cell(w=25, h=15, txt=str(subject_teacher[cell]), border=1)
            else:
                pdf.cell(w=25.3, h=15, txt=str(subject_teacher[cell]), border=1)
    
    
    pdf.set_y(250)
    pdf.set_x(20)
    pdf.cell(w=20, h=10, txt="Result Analysis Incharge")
    pdf.set_x(140)
    pdf.cell(w=20, h=10, txt="Computer Engineering Dept.")


    pdf.output(fname)





