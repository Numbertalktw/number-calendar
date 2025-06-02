import streamlit as st
import datetime
import pandas as pd
from io import BytesIO
import calendar
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side

# ===== 主日數與幸運物件資料 =====
day_meaning = {
    1: {"名稱": "創造日", "星": "⭐⭐⭐⭐"},
    2: {"名稱": "連結日", "星": "⭐⭐"},
    3: {"名稱": "表達日", "星": "⭐⭐⭐"},
    4: {"名稱": "實作日", "星": "⭐⭐⭐"},
    5: {"名稱": "行動日", "星": "⭐⭐⭐⭐"},
    6: {"名稱": "關係日", "星": "⭐⭐⭐"},
    7: {"名稱": "內省日", "星": "⭐"},
    8: {"名稱": "成果日", "星": "⭐⭐⭐⭐"},
    9: {"名稱": "釋放日", "星": "⭐⭐"},
}

lucky_map = {
    1: {"色": "🔴 紅色", "水晶": "紅瑪瑙", "小物": "原子筆"},
    2: {"色": "🟠 橘色", "水晶": "太陽石", "小物": "月亮吊飾"},
    3: {"色": "🟡 黃色", "水晶": "黃水晶", "小物": "紙膠帶"},
    4: {"色": "🟢 綠色", "水晶": "綠幽靈", "小物": "方形石頭"},
    5: {"色": "🔵 淺藍色", "水晶": "拉利瑪", "小物": "交通票卡"},
    6: {"色": "🔷 靛色", "水晶": "青金石", "小物": "愛心吊飾"},
    7: {"色": "🟣 紫色", "水晶": "紫水晶", "小物": "書籤"},
    8: {"色": "💗 粉色", "水晶": "粉晶", "小物": "鋼筆"},
    9: {"色": "⚪ 白色", "水晶": "白水晶", "小物": "小香包"},
}

combination_guidance = {
    "11/2": "這是靈魂覺醒的日子，勇敢面對真實的自己，感受內心的渴望。",
    "12/3": "今天適合展現你的表達天賦，讓創意與活力充滿周遭。",
    "13/4": "這一天需要穩健的行動和計畫，堅持不懈才能實現目標。",
    "14/5": "轉變與冒險的日子，勇敢跳出舒適圈，迎接挑戰。",
    "15/6": "關注家庭與關係，照顧好親密連結。",
    "16/7": "內在覺察與療癒的日子，沉澱自我，進行心靈探索。",
    "23/5": "展現創意與靈感的日子，讓你的點子閃耀光芒。",
    "32/5": "這一天，創意與行動的平衡將帶來新的計畫，準備好啟動變革。",
    "41/5": "務實的行動將與創意結合，為新機會打下基礎。",
    "50/5": "今天適合拓展視野，放眼未來，勇敢踏出第一步。",
    "59/14/5": "這是轉變與療癒的日子，記得放下過往的重擔。",
    "60/6": "關注愛與支持，今天適合與家人朋友共度時光。",
    "69/15/6": "這是愛與行動結合的日子，分享你的關懷。",
    "70/7": "沉澱思緒，進行深層的學習與反思。",
    "79/16/7": "深度探索與覺察的日子，反思內在與外界的連結，找到內在平靜。",
}

def reduce_to_digit(n):
    while n > 9:
        n = sum(int(x) for x in str(n))
    return n

def format_layers(total):
    mid = sum(int(x) for x in str(total))
    return f"{total}/{mid}/{reduce_to_digit(mid)}" if mid > 9 else f"{total}/{mid}"

def get_flowing_year_ref(query_date, bday):
    return query_date.year - 1 if query_date < datetime.date(query_date.year, bday.month, bday.day) else query_date.year

def get_flowing_month_ref(query_date, bday):
    return query_date.month - 1 if query_date.day < bday.day else query_date.month

def style_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="流年月曆")
        ws = writer.book["流年月曆"]
        for cell in ws[1]: cell.font = Font(bold=True, color="FFFFFF"); cell.fill = PatternFill("solid", fgColor="4F81BD")
        for row in ws.iter_rows(): [setattr(cell, 'alignment', Alignment(horizontal="center", vertical="center")) for cell in row]
        for row in ws.iter_rows(): ws.row_dimensions[row[0].row].height = 30
    return output

def get_guidance_for_day(flowing_day):
    return combination_guidance.get(flowing_day, "這是平凡但充滿潛力的一天，請保持正念與專注。")

st.set_page_config(page_title="樂覺製所生命靈數", layout="centered")
st.title("🧭 樂覺製所生命靈數")
st.markdown("在數字之中，\n我們與自己不期而遇。\n**Be true, be you — 讓靈魂，自在呼吸。**")

birthday = st.date_input("請輸入生日", datetime.date(1990,1,1))
year = st.number_input("年份",1900,2100,datetime.datetime.now().year)
month = st.selectbox("月份",list(range(1,13)),index=datetime.datetime.now().month-1)

if st.button("🎉 產生日曆"):
    last_day = calendar.monthrange(year, month)[1]
    days = pd.date_range(datetime.date(year,month,1),datetime.date(year,month,last_day))
    data = []
    for d in days:
        fd_total = sum(int(x) for x in f"{birthday.year}{birthday.month:02}{d.day:02}")
        flowing_day = format_layers(fd_total)
        main_number = reduce_to_digit(fd_total)
        lucky = lucky_map.get(main_number, {})
        guidance = get_guidance_for_day(flowing_day)
        data.append({
            "日期": d.strftime("%Y-%m-%d"),
            "星期": d.strftime("%A"),
            "流年": format_layers(sum(int(x) for x in f"{get_flowing_year_ref(d,birthday)}{birthday.month:02}{birthday.day:02}")),
            "流月": format_layers(sum(int(x) for x in f"{birthday.year}{get_flowing_month_ref(d,birthday):02}{birthday.day:02}")),
            "流日": flowing_day,
            "運勢指數": day_meaning.get(main_number, {}).get("星", ""),
            "指引": guidance,
            "幸運色": lucky.get("色", ""), "水晶": lucky.get("水晶", ""), "幸運小物": lucky.get("小物", "")
        })
    df = pd.DataFrame(data)
    st.dataframe(df)
    st.download_button("📥 下載日曆", style_excel(df).getvalue(), f"LuckyCalendar_{year}_{month}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
