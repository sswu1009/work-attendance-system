from flask import Flask, render_template, request, send_file
import openpyxl, os, json
from datetime import datetime
from openpyxl.styles import PatternFill

app = Flask(__name__)

# 讀取設定檔
with open('config.json', 'r', encoding='utf-8') as f:
    config = json.load(f)

TEMPLATE_PATH = config["input_file"]
OUTPUT_FOLDER = config["output_folder"]

REASON_OPTIONS = [
    "休假", "曠職", "體檢", "年休返泰",
    "事假返鄉", "工傷", "病假", "待返",
    "遣返", "提前解聘", "逃跑", "調派"
]

# ✅ 固定寫死的管理員名單（無預設值，前端會有「請選擇管理員」）
MANAGER_OPTIONS = ["鄭豐原", "黃國昌", "柯文哲"]

def update_excel(absentees, weather, manager_name=None):
    wb = openpyxl.load_workbook(TEMPLATE_PATH)
    ws_main = wb["出勤表"]
    ws_log = wb["休假調查表(新)"]

    reason_map = {emp_id.strip(): reason for emp_id, reason in absentees}
    emp_columns = [2, 8, 14, 20, 26]
    start_row, end_row = 6, 61

    # 各種原因對應顏色
    fill_colors = {
        "休假": "FFFF00",  # 黃色
        "曠職": "FF6666",  # 紅色
        "體檢": "B7DEE8",  # 天藍
        "年休返泰": "D9EAD3",  # 淺綠
        "事假返鄉": "D0E0E3",  # 灰藍
        "工傷": "FFD966",  # 淺黃
        "病假": "C9DAF8",  # 淡藍
        "待返": "EAD1DC",  # 紫粉
        "遣返": "F6B26B",  # 橘
        "提前解聘": "A4C2F4",  # 藍紫
        "逃跑": "E06666",  # 深紅
        "調派": "76A5AF"   # 墨綠藍
    }

    count_map = {key: 0 for key in REASON_OPTIONS}

    for row in range(start_row, end_row + 1):
        for col in emp_columns:
            emp_cell = ws_main.cell(row=row, column=col)
            upper_cell = ws_main.cell(row=row, column=col + 1)
            emp_id = str(emp_cell.value).strip() if emp_cell.value else ""
            if len(emp_id) < 3:
                continue

            if emp_id in reason_map:
                reason = reason_map[emp_id]
                upper_cell.value = "X"
                fill_color = fill_colors.get(reason, "DDDDDD")  # 預設灰
                emp_cell.fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type="solid")
                if reason in count_map:
                    count_map[reason] += 1
            else:
                upper_cell.value = "V"

    # C4：yyyy年mm月dd日 星期X
    weekdays = ['一', '二', '三', '四', '五', '六', '日']
    today = datetime.now()
    formatted_date = today.strftime(f"%Y年%m月%d日 星期{weekdays[today.weekday()]}")
    ws_main["C4"].value = formatted_date

    # 天氣：P4/S4/V4 打 X
    weather_map = {"晴": "P4", "陰": "S4", "雨": "V4"}
    if weather in weather_map:
        ws_main[weather_map[weather]].value = "X"

    # 休假調查表日期 I2
    ws_log["I2"].value = today.strftime("%Y年%m月%d日")

    # 休假調查表列表
    insert_row = 5
    serial_number = 1
    today_str = today.strftime("%m/%d")
    for emp_id, reason in absentees:
        ws_log.cell(row=insert_row, column=1).value = today_str
        ws_log.cell(row=insert_row, column=2).value = serial_number
        ws_log.cell(row=insert_row, column=3).value = "GC01"
        ws_log.cell(row=insert_row, column=4).value = emp_id
        ws_log.cell(row=insert_row, column=5).value = reason
        ws_log.cell(row=insert_row, column=6).value = "宿舍"
        insert_row += 1
        serial_number += 1

    # 統計欄位
    ws_main["C62"].value = count_map["體檢"]
    ws_main["C63"].value = count_map["工傷"]
    ws_main["C64"].value = count_map["遣返"]
    ws_main["K62"].value = count_map["休假"]
    ws_main["K63"].value = count_map["曠職"]
    ws_main["K64"].value = count_map["提前解聘"]
    ws_main["T62"].value = count_map["年休返泰"]
    ws_main["T63"].value = count_map["病假"]
    ws_main["T64"].value = count_map["逃跑"]
    ws_main["AA62"].value = count_map["事假返鄉"]
    ws_main["AA63"].value = count_map["待返"]
    ws_main["AA64"].value = count_map["調派"]

    # 出工人數 = D66 - 缺勤人數
    ws_main["L66"].value = int(ws_main["D66"].value or 0) - len(absentees)

    # S69：移工管理員：吳廷湘 <姓名>（保持原樣）
    if manager_name:
        ws_main["S69"].value = f"移工管理員：吳廷湘 {manager_name}"

    if not os.path.exists(OUTPUT_FOLDER):
        os.makedirs(OUTPUT_FOLDER, exist_ok=True)

    output_file = f"{OUTPUT_FOLDER}/每天出工統計表_{today.strftime('%Y-%m-%d')}.xlsx"
    wb.save(output_file)
    return output_file

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        emp_ids = request.form.getlist('emp_id')
        reasons = request.form.getlist('reason')
        absentees = [(emp_ids[i], reasons[i]) for i in range(len(emp_ids)) if emp_ids[i].strip()]
        weather = request.form.get('weather')

        # 取得移工管理員（沒選會是 ""）
        manager_name = request.form.get('manager', '').strip()

        result_path = update_excel(absentees, weather, manager_name)
        return send_file(result_path, as_attachment=True)

    # ✅ managers 從後端提供給前端（畫布模板已用 Jinja 渲染）
    return render_template('index.html', reasons=REASON_OPTIONS, managers=MANAGER_OPTIONS)

if __name__ == '__main__':
    app.run(debug=True)
