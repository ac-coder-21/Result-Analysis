import tkinter as tk
import datetime
import ttkbootstrap as tb
from tkinter import filedialog
import os
import shutil
import webbrowser
from analysis import analysis_pdf
from sub_topper import subject_topper_pdf
from topper import topper


main_frame = tk.Tk()
main_frame.geometry("1400x900")
bg = tk.PhotoImage(file='bg_rules.png')
my_bg = tk.Label(main_frame, image=bg)
my_bg.place(x=0, y=0, relwidth=1, relheight=1)
main_frame.title("Result Analysis - Lokmanya Tilak College of Engineering")
main_frame.iconbitmap('logo_solo_res.png')

func =''

def overall_topper():
    global func
    func = 'topper'
    func_selected_text.config(text='OverAll Topper Function')

def analysis():
    global func
    func = 'analysis'
    func_selected_text.config(text='Analysis Function')

def subject_topper():
    global func
    func = 'subject_topper'
    func_selected_text.config(text='Subject Topper Function')


overall_topper_btn = tk.PhotoImage(file='btn_overall_topper.png')
overall_topper_btn_label = tk.Button(image=overall_topper_btn, highlightthickness=0, borderwidth=0, command=overall_topper)
overall_topper_btn_label.place(x=330, y=165)

analysis_btn = tk.PhotoImage(file='btn_analysis.png')
analysis_btn_label = tk.Button(image=analysis_btn, highlightthickness=0, borderwidth=0, command=analysis)
analysis_btn_label.place(x=700, y=165)

sub_topper = tk.PhotoImage(file='btn_sub_topper.png')
sub_topper_label = tk.Button(image=sub_topper, highlightthickness=0, borderwidth=0, command=subject_topper)
sub_topper_label.place(x=1000, y=165)

first_year = 1994
current_year = datetime.datetime.now().year
diff_year = current_year-first_year

all_years = []
for year in range(0, diff_year+2):
    if year == 0:
        academic_year = str(first_year) +'-' +str(first_year+year)
        next_year = first_year + year
    else:
        academic_year = str(next_year) +'-' +str(first_year+year)
        all_years.append(academic_year)
        next_year = first_year + year


class_list = ['FE', 'SE', 'TE', 'BE']

year_combo_label = tb.Label(main_frame, text="Select Year", borderwidth=0, background='#96baca', foreground='#694535', font='Helvetica 10 bold')
year_combo_label.place(x=625, y=300)
year_combo = tb.Combobox(main_frame, bootstyle='success', values=all_years, width=50, state='readonly')
year_combo.place(x=800, y=290)

class_combo_label = tb.Label(main_frame, text='Select Class', borderwidth=0, background='#96baca', foreground='#694535', font='Helvetica 10 bold')
class_combo_label.place(x=622, y=370)
class_combo = tb.Combobox(main_frame, bootstyle='success', values=class_list, width=50, state='readonly')
class_combo.place(x=800, y=360)

sem = ''
sem_arr = ['ODD', 'EVEN']

sem_combo_label = tb.Label(main_frame, text='Select Semester', borderwidth=0, background='#96baca', foreground='#694535', font='Helvetica 10 bold')
sem_combo_label.place(x=595, y=440)
sem_combo = tb.Combobox(main_frame, bootstyle='success', values=sem_arr, width=50, state='readonly')
sem_combo.place(x=800, y=430)

file_dict_label = tb.Label(main_frame, text='Select File', borderwidth=0, background='#96baca', foreground='#694535', font='Helvetica 10 bold')
file_dict_label.place(x=625, y=510)
file_dict_text = tb.Label(main_frame, width=46, text='Select File', borderwidth=2, background='#ffffff', foreground='#694535', font='Helvetica 10 bold', padding=4 )
file_dict_text.place(x=800, y=500)

file_save_label = tb.Label(main_frame, text='Select Directory', borderwidth=0, background='#96baca', foreground='#694535', font='Helevetica 10 bold')
file_save_label.place(x=595, y=580)
file_save_text = tb.Label(main_frame, width=46, text='Select File', borderwidth=2, background='#ffffff', foreground='#694535', font='Helvetica 10 bold', padding=4 )
file_save_text.place(x=800, y=570)

func_selected_label = tb.Label(main_frame, text='Function Selected', borderwidth=0, background='#96baca', foreground='#694535', font='Helevetica 10 bold')
func_selected_label.place(x=585, y=650)
func_selected_text = tb.Label(main_frame, width=46, text='Select Function', borderwidth=2, background='#ffffff', foreground='#694535', font='Helvetica 10 bold', padding=4 )
func_selected_text.place(x=800, y=640)

generated_label = tb.Label(main_frame, text='', borderwidth=0, background='#96baca', foreground='#694535', font='Helvetica 10 bold')
generated_label.place(x=650, y=790)

def select_file():
    global file_details
    filedailouge = filedialog.askopenfilename(title='Select File', filetypes=(('xlsx files', '*.xlsx'),('all files', '*.*')))
    file_dict_text.config(text=filedailouge)
    file_details = filedailouge

def select_destiantion_directory():
    global destination_folder
    destination_folder = filedialog.askdirectory(title='Select Destination Folder')
    file_save_text.config(text=destination_folder)

file_png = tk.PhotoImage(file='file.png')
file_png_label = tk.Button(image=file_png, highlightthickness=0, borderwidth=0, command=select_file)
file_png_label.place(x=1150, y=500, height=26, width=26)

folder_png = tk.PhotoImage(file='folder.png')
folder_png_label = tk.Button(image=folder_png, highlightthickness=0, borderwidth=0, command=select_destiantion_directory)
folder_png_label.place(x=1150, y=570, height=26, width=26)

def generate():
    global sem_details
    sem_details =  sem_combo.get()
    class_details = class_combo.get()
    year_details = year_combo.get()

    path = destination_folder + '/' + year_details + '/'

    year_flag = os.path.exists(path=path)
    if year_flag == False:
        os.mkdir(path=path)

    path += class_details + '/'
    class_flag = os.path.exists(path=path)
    if class_flag == False:
        os.mkdir(path=path)

    path += sem_details +'/'
    sem_flag = os.path.exists(path=path)
    if sem_flag == False:
        os.mkdir(path=path)

    if func == 'analysis':
        name_file = year_details + '-' + class_details + '-' + sem_details + '-' + func + '.pdf'
        analysis_pdf(file=file_details, fname=name_file)
        shutil.copy(name_file, path)
        os.remove(name_file)
        
    if func == 'subject_topper':
        name_file = year_details + '-' + class_details + '-' + sem_details + '-' + func + '.pdf'
        subject_topper_pdf(file=file_details, fname=name_file)
        shutil.copy(name_file, path)
        os.remove(name_file)
    
    if func == 'topper':
        name_file = year_details + '-' + class_details + '-' + sem_details + '-' + func + '.pdf'
        topper(file=file_details, fname=name_file)
        shutil.copy(name_file, path)
        os.remove(name_file)

    generated_label.config(text=name_file + ' generated')
    def open_folder():
        webbrowser.open(path)
    open_btn = tb.Button(main_frame, text='Open Folder', style='primary', command=open_folder)
    open_btn.place(x= 1000, y=780)
    
    
    
generate_button = tb.Button(text='Generate', bootstyle='success, outline', command=generate)
generate_button.place(x=740, y=730)


#Rules


main_frame.resizable(False, False)
main_frame.mainloop() 