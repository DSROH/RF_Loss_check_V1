# %%
import glob

import ttkbootstrap as ttkbst
from ttkbootstrap.constants import *
import tkinter as tk
import _Function as func
import Check_Daseul as daseul
import Check_Pathloss as path
import tkinter.messagebox as msgbox


def condition(value, filename):
    # try:
        if value == 1:  # Daseul
            daseul.Daseul_plot_figure(filename)
        elif value == 2:  # Pathloss
            path.Pathloss_Plot_figure(filename, Result_var)
    # except Exception as e:
    #     msgbox.showwarning("Warning", e)



Win_GUI = ttkbst.Window(title="Pathloss Cal PGM V230427", themename="cosmo")
Win_GUI.attributes("-topmost", True)
Win_GUI.geometry("625x165")

Left_frame = ttkbst.Frame(Win_GUI)
Left_frame.place(x=0, y=0, width=620, height=165)

Result_var = ttkbst.BooleanVar()
Option_var = ttkbst.IntVar()

# 리스트 프레임
list_frame = ttkbst.Frame(Left_frame)
list_frame.place(x=5, y=50, width=620, height=75)

scrollbar = tk.Scrollbar(list_frame)
scrollbar.place(x=588, y=0, width=20, height=75)

list_file = tk.Listbox(list_frame, height=5, yscrollcommand=scrollbar.set)
list_file.place(x=5, y=0, width=580, height=75)

# Cal log : log 폴더의 CSV 파일 자동 입력
for fname in glob.glob("C:\\DGS\\LOGS\\*.csv"):
    list_file.insert(tk.END, fname)

scrollbar.config(command=list_file.yview)

# %%
# 경로 프레임
Selected_lossfile = ttkbst.BooleanVar()
Selected_lossfile.set(False)

path_frame = ttkbst.Frame(Left_frame)
path_frame.place(x=5, y=125, width=630, height=40)

lossfile_chkbox = ttkbst.Checkbutton(path_frame, style="info.TCheckbutton", variable=Selected_lossfile)
lossfile_chkbox.place(x=5, y=5, width=20, height=30)

path_lossfile = ttkbst.Entry(path_frame)
# spc 파일 경로 사전입력
path_lossfile.insert(0, "D:\\")
path_lossfile.place(x=25, y=5, width=345, height=30)

btn_spc = ttkbst.Button(
    path_frame,
    text="Browse lossfile (F8)",
    style="info.TButton",
    command=lambda: [func.browse_lossfile(path_lossfile, Selected_lossfile)],
)
btn_spc.place(x=375, y=5, width=130, height=30)

btn_transf = ttkbst.Button(
    path_frame,
    text="Atten File (F9)",
    style="info.TButton",
    command=lambda: [
        func.transf_to_attentable(list_file.get(0, tk.END), path_lossfile.get(), Result_var.get(), Selected_lossfile.get())
    ],
)
btn_transf.place(x=510, y=5, width=100, height=30)

# %%
# Cal log 파일 선택
file_frame = ttkbst.Frame(Left_frame)
file_frame.place(x=5, y=0, width=615, height=50)

btn_add_file1 = ttkbst.Button(
    file_frame, text="Daseul log 추가 (F1)", command=lambda: [Option_var.set(1), func.add_file("Daseul", list_file)]
)
btn_add_file1.place(x=5, y=10, width=140, height=30)

btn_add_file2 = ttkbst.Button(
    file_frame, text="Pathloss log 추가 (F2)", command=lambda: [Option_var.set(2), func.add_file("Path", list_file)]
)
btn_add_file2.place(x=150, y=10, width=145, height=30)

btn_Plot = ttkbst.Button(file_frame, text="Plot (F5)", command=lambda: [condition(Option_var.get(), list_file.get(0, tk.END))])
btn_Plot.place(x=490, y=10, width=120, height=30)

chkbox1 = ttkbst.Checkbutton(file_frame, text="Include Failed log", variable=Result_var)
# chkbox1.deselect()
chkbox1.place(x=360, y=20)

# %%
Win_GUI.bind("<F1>", lambda event: [func.add_file("Daseul", list_file)])
Win_GUI.bind("<F2>", lambda event: [func.add_file("Path", list_file)])
Win_GUI.bind("<F5>", lambda event: [condition(Option_var.get(), list_file.get(0, tk.END))])
Win_GUI.bind("<F8>", lambda event: [func.browse_lossfile(path_lossfile, Selected_lossfile)])
Win_GUI.bind(
    "<F9>",
    lambda event: [
        func.transf_to_attentable(list_file.get(0, tk.END), path_lossfile.get(), Result_var.get(), Selected_lossfile.get())
    ],
)

Win_GUI.resizable(False, False)
Win_GUI.mainloop()