# -*- coding:utf-8 -*-
# %%
import os
import sys
import pandas as pd
from datetime import datetime, timedelta, time
import time as tm
import numpy as np
import tkinter as tk
import tkinter.filedialog
import tqdm
from styleframe import StyleFrame, Styler, utils

if getattr(sys, "frozen", False):
    # 実行ファイルがあるディレクトリ
    dir_path = os.path.dirname(sys.executable)
else:
    # スクリプトがあるディレクトリ
    dir_path = os.path.dirname(os.path.abspath(__file__))
# %%
# ルートウィンドウ作成
root = tk.Tk()
# ルートウィンドウの非表示
root.withdraw()
# %%
fp = tkinter.filedialog.askopenfilename(
    filetypes=[("Excel book", "*.xlsx")], title="施設予約ファイルを選択", initialdir=dir_path
)
# %%
df = pd.read_excel(fp, header=2)
# %%
df["開始時刻"] = df["時間"].str[:5].copy()
df["終了時刻"] = df["時間"].str[-5:].copy()
# %%
startTime = {
    "1時限": "9:00",
    "2時限": "10:40",
    "昼休み": "12:10",
    "3時限": "13:00",
    "4時限": "14:40",
    "5時限": "16:20",
    "6時限": "18:00",
    "7時限": "19:40",
}
# %%
# 施設予約の処理
schedule = df[df["予約区分"] == "施設予約"].copy()
tmp = pd.DataFrame()
start_times = {k: datetime.strptime(v, "%H:%M") for k, v in startTime.items()}
# %%
# 各予約に対して時限を割り当てる
for _, row in tqdm.tqdm(schedule.iterrows()):
    start_time = datetime.strptime(row["開始時刻"], "%H:%M")
    end_time = datetime.strptime(row["終了時刻"], "%H:%M")
    # 終日予約の場合の処理
    if start_time.time() == time(0, 0):
        start_time = datetime.strptime("09:00", "%H:%M")
    if end_time.time() == time(0, 0):
        end_time = datetime.strptime("21:10", "%H:%M")
    # 時限を割り当てる時間帯を取得
    time_range = pd.date_range(
        start=start_time, end=end_time - timedelta(minutes=1), freq="10T"
    )

    # 各時間帯に対して時限を割り当てる
    for dt in time_range:
        for key, value in start_times.items():
            if dt.time() == value.time():
                new_row = row.copy()
                new_row["開講時限"] = key
                tmp = pd.concat([tmp, new_row], axis=1)
schedule = tmp.T
schedule = schedule.dropna(subset=["開講時限"]).drop_duplicates().fillna("-")


# %%
# ここから講義の処理
lectures = df[df["予約区分"] == "講義"].copy()
lectures["開講時限"] = lectures["開講時限"].astype(int).astype(str) + "時限"
# %%
scheduleDf = schedule[["日付", "開講時限", "施設種別", "施設", "タイトル", "使用者氏名"]]
lecturesDf = lectures[["日付", "開講時限", "施設種別", "施設", "科目名称", "主担当教員"]]
lecturesDf.columns = ["日付", "開講時限", "施設種別", "施設", "タイトル", "使用者氏名"]
srcDf = pd.concat([scheduleDf, lecturesDf], axis="index")
srcDf["予約名"] = srcDf["タイトル"] + "\r\n" + srcDf["使用者氏名"]
# %%
# 転記する枠を作成
cols = [
    "index",
    "日付",
    "曜日",
    "施設種別",
    "施設",
    "1時限",
    "2時限",
    "昼休み",
    "3時限",
    "4時限",
    "5時限",
    "6時限",
    "7時限",
]
dstDf = srcDf[["日付", "施設種別", "施設"]].drop_duplicates().reset_index().copy()
dstDf[cols[5:]] = np.nan

# %%
# dstDfにsrcDfの予約名を転記する
# 予約重複があった時の例外処理 改行して追記
for i, row in tqdm.tqdm(srcDf.iterrows()):
    date = row["日付"]
    datetime = row["開講時限"]
    facility = row["施設"]
    reservation = row["予約名"]
    index = dstDf[(dstDf["日付"] == date) & (dstDf["施設"] == facility)].index[0]
    if pd.isnull(dstDf.at[index, datetime]):
        dstDf.at[index, datetime] = reservation
    else:
        dstDf.at[index, datetime] = dstDf.at[index, datetime] + "\r\n" + reservation
# %%
wd = {
    0: "月",
    1: "火",
    2: "水",
    3: "木",
    4: "金",
    5: "土",
    6: "日",
}
dstDf["曜日"] = ""
for i, row in dstDf.iterrows():
    dstDf.loc[i, "曜日"] = wd[row["日付"].weekday()]

dstDf = dstDf[cols]
# %%
dstDf["日付"] = dstDf["日付"].astype(str)

dstDf.sort_values(by=["日付", "施設"], inplace=True)
dstDf.drop("index", axis=1, inplace=True)
# %%

with StyleFrame.ExcelWriter(os.path.dirname(fp) + r"\施設使用一覧.xlsx") as writer:
    # ③ DataFrame/dfをStyleFrameクラスの引数にしてインスタンス生成
    sf = StyleFrame(dstDf)
    sf.set_column_width_dict(
        {
            "日付": 12,
            "曜日": 6,
            "施設種別": 28,
            "施設": 17,
        }
    )
    sf.set_column_width(columns=cols[5:], width=37)
    sf.set_row_height(rows=list(range(2, len(dstDf) + 2)), height=27)

    # StyleFrameのto_excelメソッドで書き込む
    style = Styler(horizontal_alignment=utils.horizontal_alignments.left, font_size=11)
    sf.apply_column_style(cols_to_style=cols[1:], styler_obj=style)
    sf.to_excel(writer, index=False, sheet_name="施設利用一覧")
print("完了しました")
tm.sleep(3)
