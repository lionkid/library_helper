import pandas as pd

# === 1. 載入筆劃對照表 ===
stroke_df = pd.read_csv("unihan_strokes.csv")
stroke_dict = dict(zip(stroke_df["字"], stroke_df["筆劃數"]))

# === 2. 載入作者資料 ===
author_df = pd.read_excel("作者號.xlsx")

# === 3. 常見複姓清單 ===
compound_surnames = [
    "歐陽", "司馬", "上官", "諸葛", "南宮", "夏侯", "司徒", "尉遲", "皇甫", "呼延", "慕容",
    "軒轅", "公孫", "澹臺", "公羊", "長孫", "宇文", "司空", "仲孫", "鍾離", "段干", "東方",
    "西門", "南門", "百里", "北堂", "樂正", "壤駟", "公冶", "宓羲", "夾谷", "微生", "羊舌",
    "宗政", "濮陽", "淳于", "單于", "太叔", "申屠"
]

# === 4. 拆解姓名並查筆劃數 ===
def analyze_name(name):
    name = str(name)
    for compound in compound_surnames:
        if name.startswith(compound):
            surname = compound
            given = name[len(compound):]
            break
    else:
        surname = name[0]
        given = name[1:]

    given1 = given[0] if len(given) > 0 else ""
    given2 = given[1] if len(given) > 1 else ""

    return pd.Series([
        surname,
        stroke_dict.get(surname, 0),
        stroke_dict.get(given1, 0),
        stroke_dict.get(given2, 0)
    ])

# 套用到 DataFrame
author_df[["姓氏", "姓筆劃", "名1筆劃", "名2筆劃"]] = author_df["作者"].apply(analyze_name)

# === 5. 依筆劃排序 ===
sorted_df = author_df.sort_values(by=["姓筆劃", "姓氏", "名1筆劃", "名2筆劃"]).reset_index(drop=True)

# === 6. 輸出為 Excel 檔案 ===
sorted_df.to_excel("作者筆劃排序結果.xlsx", index=False)
print("✅ 排序完成，已輸出為 作者筆劃排序結果.xlsx")
