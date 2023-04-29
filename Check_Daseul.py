import os
from math import ceil, floor
import matplotlib.pyplot as plt

plt.rc("font", family="Malgun Gothic")
plt.rc("axes", unicode_minus=False)

import pandas as pd
import numpy as np

from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment
from openpyxl.styles.numbers import builtin_format_code

from _RF_loss_Spec import *
import _Function as func
from datetime import datetime


def Daseul_plot_figure(filename):
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

    fig, ax = plt.subplots(nrows=1, ncols=2, figsize=(22, 7))
    df_BtoB_1st = pd.DataFrame()
    df_BtoB_2nd = pd.DataFrame()
    df_RFSW1 = pd.DataFrame()

    if filename:
        for FileNumber, file in enumerate(filename):
            fname_only = os.path.basename(file).split(".")[0]
            # Import Data
            my_cols = [str(i) for i in range(12)]  # create some col names
            df_Data = pd.read_csv(file, sep="\t|,", names=my_cols, header=None, engine="python")
            # 캘 도중 에러 발생 시 첫 열에 Nan 데이터 저장되서 count 에러 발생함 -> Drop 처리
            df_null = df_Data[df_Data["0"].isnull()].index
            df_Data.drop(df_null, inplace=True)
            df_Data = df_Data.reset_index(drop=True)
            count = df_Data[df_Data["0"].str.contains("// << Equipment Loss Table - B to B >>")].shape[0]
            # Spec.은 for문에 넣어서 count 마다 업데이트 할 수 있지만, 최종값으로 덮어쓰기 되기 때문에 1번만 실행하는 것으로 수정함.
            Type_Read = df_Data[df_Data["0"].str.contains("RF Cable Type")].iloc[:, :2].reset_index(drop=True)
            Type_Cable = int(Type_Read.iloc[0, 1].strip())
            Type_BtoB = int(Type_Read.iloc[1, 1].strip())
            Jig_list = df_Data[df_Data["0"].str.contains("JIG :")].iloc[:, :1]
            lineip_list = df_Data[df_Data["0"].str.contains("RDM_LOT :")].iloc[:, :1]

            if Type_Cable != "N/A":  # ? BtoB + RF Cable Case
                Spec_BtoB_L, Spec_BtoB_H = Type_value("BtoB", Type_BtoB)
                Spec_Cable_L, Spec_Cable_H = Type_value("RF_Cable", Type_Cable)
                BtoB1 = df_Data.index[(df_Data["0"] == "// << Equipment Loss Table - B to B >>")]
                BtoB2 = df_Data.index[(df_Data["0"] == "// << Equipment Loss Table - B to B 2 >>")]
                RFSW = df_Data.index[(df_Data["0"] == "// << Equipment Loss Table >>")]

                for Number in range(count):
                    Jig = Jig_list.iloc[Number].to_list()[0].strip()
                    lineip = lineip_list.iloc[Number].to_list()[0].split(":")[1]
                    lineip = lineip.split("_")[0].strip()

                    if (Type_BtoB == 18) or (Type_BtoB == 19):
                        BtoB_Size = 98
                    elif Type_BtoB == 62:
                        BtoB_Size = 129
                    else:
                        BtoB_Size = 58

                    if Type_Cable == 7:
                        RFSW_Size = 98
                    elif Type_Cable == 62:
                        RFSW_Size = 129
                    else:
                        RFSW_Size = 58

                    BtoB1_Start = BtoB1[Number] + 3
                    BtoB1_Stop = BtoB1_Start + BtoB_Size
                    df_BtoB1 = df_Data.iloc[BtoB1_Start:BtoB1_Stop, :2]
                    BtoB1_Value = df_BtoB1.iloc[:, 1:].reset_index(drop=True)
                    BtoB1_Value = BtoB1_Value.astype(float)
                    BtoB1_Value.columns = [f"{fname_only}_IP_{lineip}_{Jig}"]
                    BtoB1_Item = df_BtoB1["0"].str.split(" ", expand=True)
                    BtoB1_Item = BtoB1_Item.iloc[:, 3:4].reset_index(drop=True)
                    BtoB1_Item.columns = ["Frequency"]
                    df_BtoB_1st = pd.concat([df_BtoB_1st, BtoB1_Value], axis=1)

                    if BtoB2.size != 0:
                        BtoB2_Start = BtoB2[Number] + 3
                        BtoB2_Stop = BtoB2_Start + BtoB_Size
                        df_BtoB2 = df_Data.iloc[BtoB2_Start:BtoB2_Stop, :2]
                        BtoB2_Value = df_BtoB2.iloc[:, 1:].reset_index(drop=True)
                        BtoB2_Value = BtoB2_Value.astype(float)
                        BtoB2_Value.columns = [f"{fname_only}_IP_{lineip}_{Jig}"]
                        BtoB2_Item = df_BtoB2["0"].str.split(" ", expand=True)
                        BtoB2_Item = BtoB2_Item.iloc[:, 3:4].reset_index(drop=True)
                        BtoB2_Item.columns = ["Frequency"]
                        df_BtoB_2nd = pd.concat([df_BtoB_2nd, BtoB2_Value], axis=1)
                        Check_BtoB2 = True
                    else:
                        Check_BtoB2 = False

                    RFSW_Start = RFSW[Number] + 3
                    RFSW_Stop = RFSW_Start + RFSW_Size
                    df_RFSW = df_Data.iloc[RFSW_Start:RFSW_Stop, :2]
                    RFSW_Value = df_RFSW.iloc[:, 1:].reset_index(drop=True)
                    RFSW_Value = RFSW_Value.astype(float)
                    RFSW_Value.columns = [f"{fname_only}_{Jig}"]
                    RFSW_Item = df_RFSW["0"].str.split(" ", expand=True)
                    RFSW_Item = RFSW_Item.iloc[:, 2:3].reset_index(drop=True)
                    RFSW_Item.columns = ["Frequency"]
                    df_RFSW1 = pd.concat([df_RFSW1, RFSW_Value], axis=1)

                    # X_index 를 주파수로 설정할때 사용
                    # X_index_RFSW1 = df_RFSW1[f"loss_{FileNumber+Number}"].str.split(" ", expand=True)[2]
                    # X_index_RFSW1 = (
                    #     X_index_RFSW1.str.split(".00MHz", expand=True)
                    #     .iloc[:, :1]
                    #     .astype(int)
                    #     .to_numpy()
                    # )
                    # X_index_BtoB1 = df_BtoB_1st[f"loss_{FileNumber+Number}"].str.split(" ", expand=True)[3]
                    # X_index_BtoB1 = (
                    #     X_index_BtoB1.str.split(".00MHz", expand=True)
                    #     .iloc[:, :1]
                    #     .astype(int)
                    #     .to_numpy()
                    # )
                    # X_index_BtoB2 = df_BtoB_2nd[f"loss_{FileNumber+Number}"].str.split(" ", expand=True)[3]
                    # X_index_BtoB2 = (
                    #     X_index_BtoB2.str.split(".00MHz", expand=True)
                    #     .iloc[:, :1]
                    #     .astype(int)
                    #     .to_numpy()
                    # )

            else:  # ? BtoB + BtoB Case
                Spec_BtoB_L, Spec_BtoB_H = Type_value("BtoB", Type_BtoB)
                Spec_Cable_L, Spec_Cable_H = Type_value("RF_Cable", Type_Cable)
                BtoB1 = df_Data.index[(df_Data["0"] == "// << Equipment Loss Table - B to B >>")]
                BtoB2 = df_Data.index[(df_Data["0"] == "// << Equipment Loss Table - B to B 2 >>")]

                for Number in range(count):
                    Jig = Jig_list.iloc[Number].to_list()[0].strip()
                    lineip = lineip_list.iloc[Number].to_list()[0].split(":")[1]
                    lineip = lineip.split("_")[0].strip()

                    if (Type_BtoB == 18) or (Type_BtoB == 19):
                        BtoB_Size = 98
                    elif Type_BtoB == 62:
                        BtoB_Size = 129
                    else:
                        BtoB_Size = 58

                    BtoB1_Start = BtoB1[Number] + 3
                    BtoB1_Stop = BtoB1_Start + BtoB_Size
                    df_BtoB1 = df_Data.iloc[BtoB1_Start:BtoB1_Stop, :2]
                    BtoB1_Value = df_BtoB1.iloc[:, 1:].reset_index(drop=True)
                    BtoB1_Value = BtoB1_Value.astype(float)
                    BtoB1_Value.columns = [f"{fname_only}_IP_{lineip}_{Jig}"]
                    BtoB1_Item = df_BtoB1["0"].str.split(" ", expand=True)
                    BtoB1_Item = BtoB1_Item.iloc[:, 3:4].reset_index(drop=True)
                    BtoB1_Item.columns = ["Frequency"]
                    df_BtoB_1st = pd.concat([df_BtoB_1st, BtoB1_Value], axis=1)

                    if BtoB2.size != 0:
                        BtoB2_Start = BtoB2[Number] + 3
                        BtoB2_Stop = BtoB2_Start + BtoB_Size
                        df_BtoB2 = df_Data.iloc[BtoB2_Start:BtoB2_Stop, :2]
                        BtoB2_Value = df_BtoB2.iloc[:, 1:].reset_index(drop=True)
                        BtoB2_Value = BtoB2_Value.astype(float)
                        BtoB2_Value.columns = [f"{fname_only}_IP_{lineip}_{Jig}"]
                        BtoB2_Item = df_BtoB2["0"].str.split(" ", expand=True)
                        BtoB2_Item = BtoB2_Item.iloc[:, 3:4].reset_index(drop=True)
                        BtoB2_Item.columns = ["Frequency"]
                        df_BtoB_2nd = pd.concat([df_BtoB_2nd, BtoB2_Value], axis=1)
                        Check_BtoB2 = True
                    else:
                        Check_BtoB2 = False

        # plot 의 옵션들은 for문 완료 후에 1번만 수행하기 위해 따로 조건문으로 실행
        if Type_Cable == "N/A":  # ? BtoB Only
            # RF Cable Type이 N/A 인 경우 Plot창 1개
            df_BtoB_1st = df_BtoB_1st.loc[:, ~df_BtoB_1st.T.duplicated()]
            # 중복열 삭제 후 카운트 리셋을 위해 0 으로 세팅, for으로 +1씩 증가
            Plot_count = 0

            for i in range(df_BtoB_1st.shape[1]):
                # fill_between 사용할 수 있도록 np로 변경
                np_BtoB1 = df_BtoB_1st.iloc[:, [i]].to_numpy(dtype="float")
                X_index_BtoB1 = np.arange(0, len(np_BtoB1), 1)

                ax[0].plot(X_index_BtoB1, np_BtoB1, marker=".", label=f"{df_BtoB_1st.columns[i]}", lw=0.5)
                Plot_count += 1

            if int(Type_Cable) == 62:
                Over4p2G_Xindex_BtoB1 = np.split(X_index_BtoB1, [97])[1]
                Over4p2_BtoB_L = np.split(Spec_BtoB_L, [97])[1]
                Over4p2_BtoB_H = np.split(Spec_BtoB_H, [97])[1]
            else:
                Over4p2G_Xindex_BtoB1 = np.split(X_index_BtoB1, [66])[1]
                Over4p2_BtoB_L = np.split(Spec_BtoB_L, [66])[1]
                Over4p2_BtoB_H = np.split(Spec_BtoB_H, [66])[1]

            ax[0].fill_between(
                X_index_BtoB1, Spec_BtoB_L, Spec_BtoB_H, color="#D1E0F9"
            )  # Matplotlib Specifiying Colors R : 84, G : 97, B : B0
            ax[0].fill_between(
                Over4p2G_Xindex_BtoB1, Over4p2_BtoB_L, Over4p2_BtoB_H, color="#F7E8EC"
            )  # Matplotlib Specifiying Colors R : 84, G : 97, B : B0

            # Lighten borders
            ax[0].spines["top"].set_alpha(1)
            ax[0].spines["bottom"].set_alpha(1)
            ax[0].spines["right"].set_alpha(1)
            ax[0].spines["left"].set_alpha(1)
            ax[0].set_xticks(X_index_BtoB1[::6], [str(d) for d in X_index_BtoB1[::6]], fontsize=8)

            # log count 5 이하에서만 legned 추가
            if Plot_count <= 10:
                ax[0].legend(frameon=False, shadow=False, fontsize=7, loc="best")

            # Axis limits
            x1, e1 = ax[0].get_xlim()
            y1, f1 = ax[0].get_ylim()
            x1 = ceil(x1)
            e1 = ceil(e1)
            y1_max = ceil(max(f1, max(Spec_BtoB_H)))
            y1_min = floor(min(y1, min(Spec_BtoB_L)))
            # print(f"y1_mean = {y1_mean}, y1_min = {y1_min}, y1_max = {y1_max}")

            ax[0].set_xlim(x1, e1)
            ax[0].set_ylim(y1_min, y1_max)
            ax[0].set_aspect("auto")
            asp1 = func.get_aspect(ax[0]) - 1
            # ax[0].set_aspect(asp1)
            ax[0].set_aspect(asp1)
            # ratio = 0.50
            # get_asp1 = ((e1-x1)/(f1-y1))*ratio
            # ax[0].set_aspect(get_asp1)

            # Draw Horizontal Tick lines
            for k1 in range(int(y1), int(f1), 1):
                ax[0].hlines(k1, xmin=int(x1), xmax=int(e1), colors="black", alpha=0.5, linestyles="--", lw=0.5)

            for l in range(0, int(e1), 6):
                ax[0].vlines(l, ymin=int(y1) - 1, ymax=int(f1), colors="black", alpha=0.5, linestyles="--", lw=0.5)

            df_BtoB1 = pd.merge(BtoB1_Item, df_BtoB_1st, left_index=True, right_index=True).reset_index(drop=True)
            # df_BtoB2 = pd.merge(BtoB2_Item, df_BtoB_2nd, left_index=True, right_index=True).reset_index(drop=True)

            plt.delaxes(ax[1])
            ax[0].change_geometry(1, 1, 1)
            ax[0].set_title(f"BtoB Type{Type_BtoB} Measured loss", fontsize=12)
            fig.set_figwidth(12)

        else:  # ? BtoB + RF Cable
            # RF Cable Type이 N/A가 아닌 경우는 Plot창 2개
            df_BtoB_1st = df_BtoB_1st.loc[:, ~df_BtoB_1st.T.duplicated()]
            df_BtoB_2nd = df_BtoB_2nd.loc[:, ~df_BtoB_2nd.T.duplicated()]

            df_RFSW1 = df_RFSW1.loc[:, ~df_RFSW1.T.duplicated()]
            # 중복열 삭제 후 카운트 리셋을 위해 0 으로 세팅, for으로 +1씩 증가
            Plot_count = 0
            for i in range(df_BtoB_1st.shape[1]):
                # fill_between 사용할 수 있도록 np로 변경
                np_BtoB1 = df_BtoB_1st.iloc[:, [i]].to_numpy(dtype="float")
                np_RFSW1 = df_RFSW1.iloc[:, [i]].to_numpy(dtype="float")

                X_index_BtoB1 = np.arange(0, len(np_BtoB1), 1)
                X_index_RFSW1 = np.arange(0, len(np_RFSW1), 1)

                ax[0].plot(X_index_BtoB1, np_BtoB1, marker=".", label=f"{df_BtoB_1st.columns[i]}", lw=0.5)
                ax[1].plot(X_index_RFSW1, np_RFSW1, marker=".", label=f"{df_RFSW1.columns[i]}", lw=0.5)
                Plot_count += 1

            if int(Type_Cable) == 62:
                Over4p2G_Xindex_BtoB1 = np.split(X_index_BtoB1, [97])[1]
                Over4p2G_Xindex_RFSW1 = np.split(X_index_RFSW1, [97])[1]
                Over4p2G_BtoB_L = np.split(Spec_BtoB_L, [97])[1]
                Over4p2G_BtoB_H = np.split(Spec_BtoB_H, [97])[1]
                Over4p2G_Cable_L = np.split(Spec_Cable_L, [97])[1]
                Over4p2G_Cable_H = np.split(Spec_Cable_H, [97])[1]
            else:
                Over4p2G_Xindex_BtoB1 = np.split(X_index_BtoB1, [66])[1]
                Over4p2G_Xindex_RFSW1 = np.split(X_index_RFSW1, [66])[1]
                Over4p2G_BtoB_L = np.split(Spec_BtoB_L, [66])[1]
                Over4p2G_BtoB_H = np.split(Spec_BtoB_H, [66])[1]
                Over4p2G_Cable_L = np.split(Spec_Cable_L, [66])[1]
                Over4p2G_Cable_H = np.split(Spec_Cable_H, [66])[1]

            # Matplotlib Specifiying Colors R : 84, G : 97, B : B0
            ax[0].fill_between(X_index_BtoB1, Spec_BtoB_L, Spec_BtoB_H, color="#D1E0F9")
            ax[0].fill_between(Over4p2G_Xindex_BtoB1, Over4p2G_BtoB_L, Over4p2G_BtoB_H, color="#F7E8EC")
            ax[1].fill_between(X_index_RFSW1, Spec_Cable_L, Spec_Cable_H, color="#D1E0F9")
            ax[1].fill_between(Over4p2G_Xindex_RFSW1, Over4p2G_Cable_L, Over4p2G_Cable_H, color="#F7E8EC")

            # Lighten borders
            ax[0].spines["top"].set_alpha(1)
            ax[0].spines["bottom"].set_alpha(1)
            ax[0].spines["right"].set_alpha(1)
            ax[0].spines["left"].set_alpha(1)
            ax[0].set_xticks(X_index_BtoB1[::6], [str(d) for d in X_index_BtoB1[::6]], fontsize=8)
            ax[1].set_xticks(X_index_RFSW1[::6], [str(d) for d in X_index_RFSW1[::6]], fontsize=8)
            # log count 5 이하에서만 legned 추가
            if Plot_count <= 10:
                ax[0].legend(frameon=False, shadow=False, fontsize=7, loc="best")
                ax[1].legend(frameon=False, shadow=False, fontsize=7, loc="best")

            # Axis limits
            x1, e1 = ax[0].get_xlim()
            y1, f1 = ax[0].get_ylim()
            x1 = ceil(x1)
            e1 = ceil(e1)
            y1_max = ceil(max(f1, max(Spec_BtoB_H)))
            y1_min = floor(min(y1, min(Spec_BtoB_L)))

            ax[0].set_xlim(x1, e1)
            ax[0].set_ylim(y1_min, y1_max)
            ax[0].set_aspect("auto")
            asp1 = func.get_aspect(ax[0]) - 1
            ax[0].set_aspect(asp1)

            x2, e2 = ax[1].get_xlim()
            y2, f2 = ax[1].get_ylim()
            x2 = ceil(x2)
            e2 = ceil(e2)
            y2_max = ceil(max(f2, max(Spec_Cable_H)))
            y2_min = floor(min(y2, min(Spec_Cable_L)))

            ax[1].set_xlim(x2, e2)
            ax[1].set_ylim(y2_min, y2_max)
            ax[1].set_aspect("auto")
            asp2 = func.get_aspect(ax[1]) - 1.2
            ax[1].set_aspect(asp2)

            # Draw Horizontal Tick lines
            for k1 in range(int(y1_min), int(y1_max), 1):
                ax[0].hlines(k1, xmin=x1, xmax=e1, colors="black", alpha=0.5, linestyles="--", lw=0.5)
            for l1 in range(0, int(e1), 6):
                ax[0].vlines(l1, ymin=y1_min, ymax=y1_max, colors="black", alpha=0.5, linestyles="--", lw=0.5)

            for k2 in range(int(y2_min), int(y2_max), 1):
                ax[1].hlines(k2, xmin=x2, xmax=e2, colors="black", alpha=0.5, linestyles="--", lw=0.5)
            for l2 in range(0, int(e2), 6):
                ax[1].vlines(l2, ymin=y2_min, ymax=y2_max, colors="black", alpha=0.5, linestyles="--", lw=0.5)

            df_BtoB1 = pd.merge(BtoB1_Item, df_BtoB_1st, left_index=True, right_index=True).reset_index(drop=True)
            # df_BtoB2 = pd.merge(BtoB2_Item, df_BtoB_2nd, left_index=True, right_index=True).reset_index(drop=True)
            df_RFSW1 = pd.merge(RFSW_Item, df_RFSW1, left_index=True, right_index=True).reset_index(drop=True)
            # ax[0].set_ylabel(f"BtoB Type {Type_BtoB} Measured loss", fontsize=12)
            # ax[1].set_ylabel(f"RFSW Type {Type_Cable} Measured loss", fontsize=12)
            ax[0].set_title(f"BtoB Type {Type_BtoB} Measured loss", fontsize=12)
            ax[1].set_title(f"RFSW Type {Type_Cable} Measured loss", fontsize=12)

        f_name = f"{os.path.splitext(filename[0])[0]}.pdf"  # filename을 확장자를 지운 후 pdf 확장자로 지정
        dir, file = os.path.split((f_name))
        Model = file.split("_")[0]
        meas_time = datetime.now().strftime("%Y-%m%d_%H_%M_%S")
        # Save Data to Excel
        Excel_file = f"Export_{Model}_Daseul_{meas_time}.xlsx"
        with pd.ExcelWriter(Excel_file) as writer:
            df_BtoB1.to_excel(writer, sheet_name="BtoB1")
            if Check_BtoB2:
                df_BtoB2.to_excel(writer, sheet_name="BtoB2")
            df_RFSW1.to_excel(writer, sheet_name="RFSW")
        func.WB_Format(Excel_file, 2, 2, 2)

        # fig.suptitle("RF Loss Cal measured Data", fontsize=15)
        plt.tight_layout()
        func.save_multi_image(f_name)
        plt.show()

    # func.open_file(f_name)
