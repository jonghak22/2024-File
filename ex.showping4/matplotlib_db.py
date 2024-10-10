import os
import sqlite3
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import tkinter as tk

# 데이터베이스 연결 설정
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_FILE = 'showping.db'  # 데이터베이스 파일명
db_path = os.path.join(BASE_DIR, DB_FILE)
connection = sqlite3.connect(db_path)
cursor = connection.cursor()

# 데이터 추출
cursor.execute("SELECT CompanyName FROM bid_list")
data = cursor.fetchall()

# 업체명별 빈도수 계산
company_counts = {}
for row in data:
    company_name = row[0]
    company_counts[company_name] = company_counts.get(company_name, 0) + 1

# Matplotlib 한글 폰트 설정\
plt.rcParams['font.size'] = 10
plt.rcParams['font.family'] = 'Malgun Gothic'  # 폰트 사용
plt.rcParams['axes.unicode_minus'] = False  # 마이너스 기호 깨짐 방지

# 그래프 생성
plt.figure(figsize=(10, 6))
plt.bar(company_counts.keys(), company_counts.values(), color='blue')
plt.xlabel('업체명')
plt.ylabel('빈도수 (5의 배수)')
plt.yticks(range(0, max(company_counts.values()) + 1, 5))
plt.xticks(rotation=45, ha='right')
plt.tight_layout()

# UI 생성
root = tk.Tk()
root.title('Company Bid Frequency')

# Matplotlib 그래프를 Tkinter에 삽입
canvas = FigureCanvasTkAgg(plt.gcf(), master=root)
canvas.draw()
canvas.get_tk_widget().pack()

# UI 실행
tk.mainloop()

# 데이터베이스 연결 종료
connection.close()
