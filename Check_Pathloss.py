import os
import matplotlib.pyplot as plt
from math import ceil, floor

plt.rc("font", family="Malgun Gothic")
plt.rc("axes", unicode_minus=False)

import pandas as pd
import numpy as np

from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment
from openpyxl.styles.numbers import builtin_format_code

import tkinter.messagebox as msgbox

from _RF_loss_Spec import *
import _Function as func
from datetime import datetime


def Pathloss_Plot_figure(filename, Result_var):
    font_style = Font(
        name="Calibri",
        size=10,
        bold=False,
        italic=False,
        vertAlign=None,  # 첨자
        underline="none",  # 밑줄
        strike=False,  # 취소선
        color="00000000",  # 블랙, # 00FF0000 Red, # 000000FF Blue
    )

    # Plot
    plt.figure(figsize=(10, 6), dpi=150)
    plt.ylabel("Measured loss", fontsize=12)
    df_BtoB1 = pd.DataFrame()

    if filename:
        for FileNumber, file in enumerate(filename):
            fname_only = os.path.basename(file).split(".")[0]
            # Import Data
            my_cols = [str(i) for i in range(10)]  # create some col names
            df_Data = pd.read_csv(file, sep="\t|,", names=my_cols, header=None, engine="python")
            Current_Type = df_Data[df_Data["0"].str.contains("Current Cable Type", na=False)]
            Current_Type = Current_Type["0"].str.split(":", expand=True).iloc[0, 1].strip()
            df_Test = df_Data.index[(df_Data["0"] == "#TEST")].to_list()

            if df_Data["0"].str.contains("SVC", na=False).any():
                df_Cabletype = df_Data[df_Data["0"].str.contains("RF Cable Type BtoB : ", na=False)].iloc[:, :1]
                Type_SVC = True
            elif Current_Type == "BtoB":
                df_Cabletype = df_Data[df_Data["0"].str.contains("RF Cable Type BtoB : ", na=False)].iloc[:, :1]
                Type_SVC = False
            else:
                df_Cabletype = df_Data[df_Data["0"].str.contains("RF Cable Type : ", na=False)].iloc[:, :1]
                Type_SVC = False

            Jig_list = df_Data[df_Data["0"].str.contains("JIG :", na=False)].iloc[:, :1]
            Result = df_Data[df_Data["0"].str.contains("RESULT :", na=False)].iloc[:, :1]
            lineip_list = df_Data[df_Data["0"].str.contains("RDM_LOT :", na=False)].iloc[:, :1]

            for Count, index in enumerate(df_Test):
                # for index in BtoB1:
                LossCal_Result = Result.iloc[Count].to_list()[0].split(":")[1]
                lineip = lineip_list.iloc[Count].to_list()[0].split(":")[1]
                lineip = lineip.split("_")[0].strip()

                # Type = int(df_Cabletype.iloc[Count].to_list()[0].split(":")[1].strip())
                # if (Type == 18) or (Type == 19):
                #     Size = 98
                # elif Type == 7:
                #     Size = 98
                # else:
                #     Size = 58

                # Loss_Start = index + 3
                # Loss_Stop = Loss_Start + Size
                # df_BtoB_1st = df_Data.iloc[Loss_Start:Loss_Stop, :4]
                # BtoB1st_Value = df_BtoB_1st.iloc[:, 1:2].reset_index(drop=True)
                # BtoB1st_Value = BtoB1st_Value.astype(float)
                # BtoB1st_Value.columns = [f"loss_{FileNumber + Count}"]
                # # Variance 측정을 위한 Dataframe
                # Var_Table = BtoB1st_Value[f"loss_{FileNumber + Count}"][0:18]
                # Variance = np.var(Var_Table)
                # BtoB1st_Item = (
                #     df_BtoB_1st["0"].str.split(" ", expand=True).reset_index(drop=True)
                # )
                # BtoB1st_Item.columns = Columns
                # BtoB1st_Item = BtoB1st_Item["Frequency"]
                # # BtoB1st_Item.drop(columns=["Meas", "Path"], inplace=True)
                # df_BtoB1 = pd.concat([df_BtoB1, BtoB1st_Value], axis=1)
                # # fill_between 사용할 수 있도록 np로 변경
                # np_BtoB1 = df_BtoB1[f"loss_{FileNumber + Count}"].to_numpy(
                #     dtype="float"
                # )

                # X_index = np.arange(0, len(np_BtoB1), 1)

                # plt.plot(
                #     X_index,
                #     np_BtoB1,
                #     marker=".",
                #     label="{}_{}".format(os.path.splitext(file)[0], Count + 1),
                #     lw=0.7,
                # )
                if Result_var.get():
                    if Type_SVC:
                        BtoB_Type = "SVC"
                    else:
                        BtoB_Type = "BtoB"

                    Type = int(df_Cabletype.iloc[Count].to_list()[0].split(":")[1].strip())
                    Jig = Jig_list.iloc[Count].to_list()[0].strip()

                    if (Type == 18) or (Type == 19):
                        Size = 98
                    elif Type == 7:
                        Size = 98
                    elif Type == 62:
                        Size = 129
                    else:
                        Size = 58

                    Loss_Start = index + 3
                    Loss_Stop = Loss_Start + Size
                    df_BtoB_1st = df_Data.iloc[Loss_Start:Loss_Stop, :4]
                    BtoB1st_Value = df_BtoB_1st.iloc[:, 1:2].reset_index(drop=True)
                    BtoB1st_Value = BtoB1st_Value.astype(float)
                    BtoB1st_Value.columns = [f"{fname_only}_IP_{lineip}_{Jig}"]
                    # Variance 측정을 위한 Dataframe
                    Var_Table = BtoB1st_Value[f"{fname_only}_IP_{lineip}_{Jig}"][0:18]
                    Variance = np.var(Var_Table)
                    BtoB1st_Item = df_BtoB_1st["0"].str.split(" ", expand=True).reset_index(drop=True)

                    if BtoB1st_Item[0].str.contains("SVC", na=False).any():
                        BtoB1st_Item.columns = ["SVC", "Meas", "BtoB_No", "Path", "Frequency"]
                    elif Current_Type == "BtoB":
                        BtoB1st_Item.columns = ["Meas", "BtoB_No", "Path", "Frequency"]
                    else:
                        BtoB1st_Item.columns = ["Meas", "Path", "Frequency"]

                    BtoB1st_Item = BtoB1st_Item["Frequency"]
                    # BtoB1st_Item.drop(columns=["Meas", "Path"], inplace=True)
                    df_BtoB1 = pd.concat([df_BtoB1, BtoB1st_Value], axis=1)
                    # fill_between 사용할 수 있도록 np로 변경

                else:
                    if LossCal_Result == "FAIL":
                        msgbox.showwarning("Warning", f"Losscal Result : Fail\nOr\nCheck 'Include Failed log' Button")
                        plt.close()
                        return
                    else:
                        if Type_SVC:
                            BtoB_Type = "SVC"
                        else:
                            BtoB_Type = "BtoB"

                        Type = int(df_Cabletype.iloc[Count].to_list()[0].split(":")[1].strip())
                        Jig = Jig_list.iloc[Count].to_list()[0].strip()

                        if (Type == 18) or (Type == 19):
                            Size = 98
                        elif Type == 7:
                            Size = 98
                        elif Type == 62:
                            Size = 129
                        else:
                            Size = 58

                        Loss_Start = index + 3
                        Loss_Stop = Loss_Start + Size
                        df_BtoB_1st = df_Data.iloc[Loss_Start:Loss_Stop, :4]
                        BtoB1st_Value = df_BtoB_1st.iloc[:, 1:2].reset_index(drop=True)
                        BtoB1st_Value = BtoB1st_Value.astype(float)
                        BtoB1st_Value.columns = [f"{fname_only}_IP_{lineip}_{Jig}"]
                        # Variance 측정을 위한 Dataframe
                        Var_Table = BtoB1st_Value[f"{fname_only}_IP_{lineip}_{Jig}"][0:18]
                        Variance = np.var(Var_Table)
                        BtoB1st_Item = df_BtoB_1st["0"].str.split(" ", expand=True).reset_index(drop=True)

                        if BtoB1st_Item[0].str.contains("SVC", na=False).any():
                            BtoB1st_Item.columns = ["SVC", "Meas", "BtoB_No", "Path", "Frequency"]
                        elif Current_Type == "BtoB":
                            BtoB1st_Item.columns = ["Meas", "BtoB_No", "Path", "Frequency"]
                        else:
                            BtoB1st_Item.columns = ["Meas", "Path", "Frequency"]

                        BtoB1st_Item = BtoB1st_Item["Frequency"]
                        # BtoB1st_Item.drop(columns=["Meas", "Path"], inplace=True)
                        df_BtoB1 = pd.concat([df_BtoB1, BtoB1st_Value], axis=1)
                        # fill_between 사용할 수 있도록 np로 변경
                        np_BtoB1 = df_BtoB1[f"{fname_only}_IP_{lineip}_{Jig}"].to_numpy(dtype="float")
                        # RF Cable Type이 N/A 인 경우 Plot창 1개

        df_BtoB1 = df_BtoB1.loc[:, ~df_BtoB1.T.duplicated()]
        Count += 1
        for i in range(df_BtoB1.shape[1]):
            # fill_between 사용할 수 있도록 np로 변경
            np_BtoB1 = df_BtoB1.iloc[:, [i]].to_numpy(dtype="float")
            X_index_BtoB1 = np.arange(0, len(np_BtoB1), 1)
            plt.plot(X_index_BtoB1, np_BtoB1, marker=".", label=f"{df_BtoB1.columns[i]}", lw=0.5)

        Spec_L, Spec_H = Type_value(BtoB_Type, Type)
        if Type == 62:
            Over4p2G_X = np.split(X_index_BtoB1, [97])[1]
            Over4p2G_L = np.split(Spec_L, [97])[1]
            Over4p2G_H = np.split(Spec_H, [97])[1]
        else:
            Over4p2G_X = np.split(X_index_BtoB1, [66])[1]
            Over4p2G_L = np.split(Spec_L, [66])[1]
            Over4p2G_H = np.split(Spec_H, [66])[1]

        plt.fill_between(X_index_BtoB1, Spec_L, Spec_H, color="#DDEBF7")  # Matplotlib Specifiying Colors R : 84, G : 97, B : B0
        plt.fill_between(
            Over4p2G_X, Over4p2G_L, Over4p2G_H, color="#F7E8EC"
        )  # Matplotlib Specifiying Colors R : 84, G : 97, B : B0
        # Lighten borders
        plt.gca().spines["top"].set_alpha(1)
        plt.gca().spines["bottom"].set_alpha(1)
        plt.gca().spines["right"].set_alpha(1)
        plt.gca().spines["left"].set_alpha(1)
        plt.xticks(X_index_BtoB1[::6], [str(d) for d in X_index_BtoB1[::6]], fontsize=8)

        if Count <= 20:
            plt.legend(frameon=False, shadow=False, fontsize=7, loc="best")

        df_BtoB1 = pd.merge(BtoB1st_Item, df_BtoB1, left_index=True, right_index=True)
        # df_BtoB1 = df_BtoB1.style.set_properties(**{"font-size": "10pt"})
        df_BtoB1_Mean = round(df_BtoB1.groupby(["Frequency"], sort=False).mean(), 2)
        df_BtoB1_Mean["Average"] = round(df_BtoB1_Mean.mean(axis=1), 2)
        df_BtoB1_Mean["Max"] = round(df_BtoB1_Mean.max(axis=1), 2)
        df_BtoB1_Mean["Min"] = round(df_BtoB1_Mean.min(axis=1), 2)

        f_name = f"{os.path.splitext(filename[0])[0]}.pdf"  # filename을 확장자를 지운 후 pdf 확장자로 지정
        dir, file = os.path.split((f_name))
        Model = file.split("_")[0]
        # Save Data to Excel
        meas_time = datetime.now().strftime("%Y-%m%d_%H_%M_%S")
        Excel_file = f"Export_{Model}_Pathloss_{meas_time}.xlsx"
        with pd.ExcelWriter(Excel_file) as writer:
            df_BtoB1_Mean.to_excel(writer, sheet_name="BtoB1_Mean")
        func.WB_Format(Excel_file, 2, 2, 2)

        # Axis limits
        s, e = plt.gca().get_xlim()
        y, f = plt.gca().get_ylim()
        s = ceil(s)
        e = ceil(e)

        y1_max = ceil(max(f, max(Spec_H)))
        y1_min = floor(min(y, min(Spec_L)))

        plt.xlim(s, e)
        plt.ylim(y1_min, y1_max)
        plt.gca().set_aspect("auto")
        asp1 = func.get_aspect(plt.gca())
        # ax[0].set_aspect(asp1)
        plt.gca().set_aspect(asp1)

        # Draw Horizontal Tick lines
        for k in range(int(y1_min), int(y1_max), 1):
            plt.hlines(k, xmin=s, xmax=e, colors="black", alpha=0.5, linestyles="--", lw=0.5)

        for l in range(0, e, 6):
            plt.vlines(
                l,
                ymin=y1_min,
                ymax=y1_max,
                colors="black",
                alpha=0.5,
                linestyles="--",
                lw=0.5,
            )
        if BtoB_Type == "SVC":
            plt.title(f"{Current_Type} {BtoB_Type} Type {Type} PathLoss Data", fontsize=15)
        else:
            plt.title(f"{Current_Type} Type {Type} PathLoss Data", fontsize=15)
        plt.tight_layout()

        func.save_multi_image(f_name)
        # plt.axes().set_aspect(aspect=0.57)
        plt.show()

    # func.open_file(f_name)
