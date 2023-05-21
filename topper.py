import pandas as pd
from PIL import Image
from fpdf import FPDF
from fpdf.enums import XPos, YPos

def topper(file, fname):
    data = pd.read_excel(file)

    title = data.iloc[0:1,:].values[0]

    topper_marks = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
    topper_name = ['', '', '', '', '', '', '', '', '', '']
    temp_marks = 0
    temp_name = ''
    topper_marks_array = ['SGPI']
    topper_name_array = ['Name']
    topper_rank_array = ['Sr. no']

    for col in range(0, len(title)):
        if(title[col] == 'Name'):
            name = data.iloc[:,col:col+1].values
        if(title[col] == 'SGPI'):
            sgpi = data.iloc[:,col:col+1].values
            for pointer in range(0, len(sgpi)):
                if(sgpi[pointer][0] == 'SGPI' or sgpi[pointer][0] == 'F'):
                    pass
                else:
                    for rank in range(0, len(topper_marks)):
                        if(topper_marks[rank] < sgpi[pointer][0]):
                            temp_marks = sgpi[pointer][0]
                            sgpi[pointer][0] = topper_marks[rank]
                            topper_marks[rank] = temp_marks

                            temp_name = name[pointer][0]
                            name[pointer][0] = topper_name[rank]
                            topper_name[rank] = temp_name

    for new_arr in range(0, len(topper_marks)):
        topper_rank_array.append(str(new_arr+1))
        topper_marks_array.append(str(topper_marks[new_arr]))
        topper_name_array.append(str(topper_name[new_arr]))


    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("helvetica", "", 10)

    logo = Image.open('logo.jpeg')
    logo_rgb = logo.convert('RGB')
    logo_display =  logo_rgb.resize((550, 100), resample=Image.NEAREST)
    pdf.image(logo_display, y=15)

    pdf.set_xy(70,50)
    pdf.cell(w=150, h=10, txt="Department of Computer Engineering", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_x(90)
    pdf.cell(w=150, h=10, txt="Result Analysis", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_x(87)
    pdf.cell(w=150, h=10, txt="OVERALL TOPPERS", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_x(75)
    pdf.cell(w=150, h=10, txt="Class: SE        SEM: IV(C SCHEME)", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_x(77)
    pdf.cell(w=150, h=10, txt="EXAMINATION:- MAY 2022 FH", new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    pdf.set_y(100)
    for cell in range(0, len(topper_marks_array)):
        pdf.set_x(20)
        pdf.cell(w=20, h=10, txt=topper_rank_array[cell], border=1)
        pdf.cell(w=110, h=10, txt=topper_name_array[cell], border=1)
        pdf.cell(w=30, h=10, txt=topper_marks_array[cell], border=1, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    
    pdf.set_y(250)
    pdf.set_x(20)
    pdf.cell(w=20, h=10, txt="Result Analysis Incharge")
    pdf.set_x(140)
    pdf.cell(w=20, h=10, txt="Computer Engineering Dept.")
    pdf.output(fname)

