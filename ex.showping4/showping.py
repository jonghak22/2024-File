
import os
import time
import math
import json
import sqlite3      # DBMS 임포트
import pandas
import webbrowser
from openpyxl import Workbook
from urllib import parse, request
from pandas import DataFrame
from bs4 import BeautifulSoup
from PyQt5.QtWidgets import *
from PyQt5 import uic
from PyQt5.QtCore import *
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import tkinter as tk

from WaitingSpinnerWidget import Overlay                # 로딩 스피너
import apikey
import ssl

# TLS 1.2를 사용하는 SSL 컨텍스트 설정
ssl_context = ssl.SSLContext(ssl.PROTOCOL_TLSv1_2)
ssl_context.set_ciphers('DEFAULT@SECLEVEL=1')

BASE_DIR = os.path.dirname(os.path.abspath(__file__))   # python실행 경로

## 고정값 설정  
DB_FILE = "showping.db"            # DB 파일명 지정
API_KEY = apikey.mykey
API_URL = "https://apis.data.go.kr/1230000/ShoppingMallPrdctInfoService06"  
OPT_NAME_BIDC = "/getShoppingMallPrdctInfoList01?"   #입찰공고 공사조회

## 검색중인URL 저장용 전역변수
url_pre = ""
url_sub = ""
## DB파일이 없으면 새로 만들고
if os.path.isfile(BASE_DIR + "//" + DB_FILE):
    con = sqlite3.connect(BASE_DIR + "//" + DB_FILE)
    cursor = con.cursor()
else:                       ####변수 변경 예정
    con = sqlite3.connect(BASE_DIR + "//" + DB_FILE)
    cursor = con.cursor()
    cursor.execute("CREATE TABLE bid_list(bidno text PRIMARY KEY, ContractDate date, Item text, ItemName text, CompanyName text, Amount text, Unit text)")
    cursor.execute("CREATE TABLE bid_saved(bidno text PRIMARY KEY, bidname text)")
    
Ui_MainWindow = uic.loadUiType(BASE_DIR+r'\showping.ui')[0]
#Ui_MainWindow = uic.loadUiType(r'D:\VSCode\MyPython_collection\G2BDataAPI\G2BDataAPI\G2BDataAPI.ui')[0]

global start_time       #데이터 다운로드 시간 계산용 전역변수
start_time = 0.0

## 데이터 크롤링을 담당할 쓰레드
class CrawlRunnable(QRunnable):
    def __init__(self, dialog):
        QRunnable.__init__(self)
        self.w = dialog

    # 크롤링 루틴
    def crawl(self):        
        # debugpy.debug_this_thread()     # 멀티 쓰레드 디버깅을 위해서 추가
        date_start = self.w.dateEdit_start.date().toString("yyyyMMdd")
        date_end = self.w.dateEdit_end.date().toString("yyyyMMdd")
        page = self.w.lineEdit_curPage.text()                            #페이지
        dminsttNm = self.w.lineEdit_cropNm.text()                           #업체명
        bidNtceNm = self.w.lineEdit.text()                                  #품명
        
        start = parse.quote(str(date_start))                                #검색기간 시작일
        end = parse.quote(str(date_end))                                    #검색기간 종료일
        keyword = parse.quote(self.w.lineEdit.text())                       #공고명 검색 키워드
        numOfRows = parse.quote(str(999))                                    #페이지당 표시할 공고 수
        page = parse.quote(str(page))                                    #다운로드할 페이지 번호

        url = ""

        if self.w.radioButton_prod.isChecked():                     # "입찰공고 공사" 선택 시
            url = (API_URL + OPT_NAME_BIDC
                + "ServiceKey="+API_KEY+"&"
                + "pageNo="+ page +"&"
                + "numOfRows="+numOfRows+"&"
                + "inqryDiv=1&"
                + "inqryBgnDate="+start+"&"
                + "inqryEndDate="+end+"&"#"2359"
                + "shopngCntrctNo=1"
                )
            
            if dminsttNm != "":                                         # 업체명 검색조건이 있으면 추가
                url = url+"&prdctIdntNoNm="+parse.quote(dminsttNm)     

            if bidNtceNm != "":
                url = url+"&prdctClsfcNoNm="+parse.quote(bidNtceNm)       #품명
                
        print("Constructed URL:", url)

        req = request.Request(url)
        resp = request.urlopen(req, context=ssl_context)

        rescode = resp.getcode()
        if(rescode==200):
            response_body = resp.read()
            html=response_body.decode('utf-8')
            soup = BeautifulSoup(html, 'lxml')

            errmsg = soup.find('errmsg')
            if(errmsg!=None):
                return 0

            ## BeautifulSoup에서 아이템 검색시 모두 소문자로 검색해야 함
            totalCount = soup.find('totalcount')    #검색조건에 해당하는 전체 공고수
            pageNo = soup.find('pageno')            #요청한 페이지 번호
            totalPageNo = math.ceil(int(totalCount.string) / int(numOfRows))    #전체 페이지 수

            self.w.label_3.setText(str(totalPageNo)+"페이지("+totalCount.string+"건)중 페이지")
            self.w.lineEdit_curPage.setText(pageNo.string)

            ## 크롤링 쓰레드에서 처리할 DB연결
            con = sqlite3.connect(BASE_DIR + "//" + DB_FILE)
            cursor = con.cursor()

            for itemElement in soup.find_all('item'):
                #조회 데이터를 DB에 입력
                if self.w.radioButton_prod.isChecked():
                    price = itemElement.cntrctprceamt.string
                    if price == None:
                        price = "0"
                    price = format(int(price),',')
                    cursor.execute("INSERT or IGNORE INTO bid_list VALUES(?,?,?,?,?,?,?);",
                                  (itemElement.prdctidntno.string,       # 쇼핑계약번호
                                  itemElement.rgstdt.string,                # 공고일시
                                  itemElement.prdctclsfcnonm.string,         # 품명
                                  itemElement.prdctspecnm.string,
                                  itemElement.cntrctcorpnm.string,          # 계약회사
                                  price,                                    # 가격
                                  itemElement.prdctunit.string))        # 단위
                    

                con.commit()    # 작업내용을 테이블에 수행
            con.close()
        else:
            print("Error Code:" + rescode)

    def run(self):
        self.crawl()
        QMetaObject.invokeMethod(self.w, "search_finish",
                                 Qt.QueuedConnection)
        
class MyDialog(QMainWindow, Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.initMainTable()

        self.overlay = Overlay(self.centralWidget())
        self.overlay.hide()
        
        self.pushButton.clicked.connect(self.btn_search)
        self.pushButton_move.clicked.connect(self.btn_move)
        self.pushButton_del.clicked.connect(self.btn_del)
        self.pushButton_xlsx.clicked.connect(self.btn_excel)

        self.pushButton_mat.clicked.connect(self.btn_mat)

        self.pushButton_kogas.clicked.connect(self.btn_kogas)
        self.pushButton_today.clicked.connect(self.btn_today)
        self.pushButton_3days.clicked.connect(self.btn_3days)
        self.pushButton_1week.clicked.connect(self.btn_1week)
        self.pushButton_1month.clicked.connect(self.btn_1month)
        self.pushButton_3months.clicked.connect(self.btn_3months)

        self.radioButton_prod.clicked.connect(self.radioB_prod)

        self.tableWidget.cellClicked.connect(self.cell_clicked)
        # self.tableWidget.cellDoubleClicked.connect(self.cell_DBclicked)

        self.dateEdit_start.setDate(QDate.currentDate().addDays(-1))
        self.dateEdit_end.setDate(QDate.currentDate())


    # Pyqt종료시 호출
    def closeEvent(self, event):
        con.close()     # DB연결 종료
        super(MyDialog, self).closeEvent(event)

    # Resize 이벤트
    def resizeEvent(self, event):
        super(MyDialog, self).resizeEvent(event)
        self.arrangecolumn()
        self.overlay.resize(event.size())

    def showEvent(self, a0):
        self.arrangecolumn()
        return super().showEvent(a0)

    # 메인테이블에 DB 데이터 표시하기
    def refreshMainTable(self):
        con = sqlite3.connect(BASE_DIR + "//" + DB_FILE)
        cursor = con.cursor()
        cursor.execute("SELECT * FROM bid_list")
        table = self.tableWidget
        table.setRowCount(0)
        for row, form in enumerate(cursor):
            table.insertRow(row)
            for column, item in enumerate(form):
                if (column<7):
                    table.setItem(row, column, QTableWidgetItem(str(item)))
        self.arrangecolumn()
        con.close()  
        
    # 메인테이블 초기화
    def initMainTable(self):
        table = self.tableWidget

        table.setColumnCount(7)
        table.setRowCount(0)
        table.setHorizontalHeaderLabels(["번호","계약일"," 품목","품목명","업체명","금액","단위"])

        self.refreshMainTable()

    # 테이블 삭제
    def btn_del(self):
        con = sqlite3.connect(BASE_DIR + "//" + DB_FILE)
        cursor = con.cursor()
        cursor.execute("DELETE FROM bid_list;")
        con.commit()    # 작업내용을 테이블에 수행
        con.close()
        self.refreshMainTable()

    def btn_excel(self):
        # SQLite 데이터베이스 연결 생성
        db_path = os.path.join(BASE_DIR, DB_FILE)
        connection = sqlite3.connect(db_path)
        cursor = connection.cursor()

        # 데이터 추출
        cursor.execute("DELETE FROM bid_list")
        data = cursor.fetchall()

        # Excel 파일 이름
        excel_file_name = "output.xlsx"

        # Excel 파일 생성
        workbook = Workbook()
        sheet = workbook.active

        # 데이터를 Excel 시트에 쓰기
        for row in data:
            sheet.append(row)

        # Excel 파일 저장
        excel_path = os.path.join(BASE_DIR, excel_file_name)
        workbook.save(excel_path)

        # 연결 닫기
        connection.close()

        print("데이터를 Excel 파일로 저장 완료:", excel_path)


    def btn_mat(self):  
        # 그레프 데이터 추출
        cursor.execute("SELECT CompanyName FROM bid_list")
        data = cursor.fetchall()

        # 업체명별 빈도수 계산
        company_counts = {}
        for row in data:
            company_name = row[0]
            company_counts[company_name] = company_counts.get(company_name, 0) + 1

        # Matplotlib 한글 폰트 설정\
        plt.rcParams['font.size'] = 6
        plt.rcParams['font.family'] = 'Malgun Gothic'  # 폰트 사용
        plt.rcParams['axes.unicode_minus'] = False  # 마이너스 기호 깨짐 방지

        # 그래프 생성
        plt.figure(figsize=(14, 8))
        plt.bar(company_counts.keys(), company_counts.values(), color='blue')
        plt.xlabel('업체명')
        plt.ylabel('빈도수 (5의 배수)')
        plt.yticks(range(0, max(company_counts.values()) + 1, 10))
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

    # 입찰공고를 검색하기
    def btn_search(self):
        global start_time
        start_time = time.time()            # 시작시간 리셋
        self.overlay.setVisible(True)       # 스피너 시작

        ## 검색페이지 요청
        table = self.tableWidget
        table.clearContents()

        runnable = CrawlRunnable(self)
        QThreadPool.globalInstance().start(runnable)

    # 페이지 이동
    def btn_move(self):
        self.btn_del()
        self.btn_search()

    # 키워드목록에서 클릭이 발생하면 해당 키워드를 에디트창에 반영
    def cell_clicked(self, row, col):
        # 동작영역을 데이터가 있는 범위내로 한정해야 함
        sel_key = self.tableWidget.item(row,col)
        if (sel_key):
            sel_key = sel_key.text()

    # 검색조건 프리셋
    ## 수요기관
    def btn_kogas(self):
        self.lineEdit_cropNm.setText("(주)케이에스아이")

    ## 검색기간
    def btn_today(self):
        self.dateEdit_start.setDate(QDate.currentDate().addDays(-1))
        self.dateEdit_end.setDate(QDate.currentDate())

    def btn_3days(self):
        self.dateEdit_start.setDate(QDate.currentDate().addDays(-2))
        self.dateEdit_end.setDate(QDate.currentDate())

    def btn_1week(self):
        self.dateEdit_start.setDate(QDate.currentDate().addDays(-7))
        self.dateEdit_end.setDate(QDate.currentDate())

    def btn_1month(self):
        self.dateEdit_start.setDate(QDate.currentDate().addMonths(-1))
        self.dateEdit_end.setDate(QDate.currentDate())

    def btn_3months(self):
        self.dateEdit_start.setDate(QDate.currentDate().addMonths(-2))
        self.dateEdit_end.setDate(QDate.currentDate())

    def radioB_prod(self):
        self.radioButton_prod.setChecked(True)

    def arrangecolumn(self):
        table = self.tableWidget
        header = table.horizontalHeader()
        twidth = header.width()
        width = []
        for column in range(header.count()):
            header.setSectionResizeMode(column, QHeaderView.ResizeToContents)
            width.append(header.sectionSize(column))

        wfactor = twidth / sum(width)
        for column in range(header.count()):
            header.setSectionResizeMode(column, QHeaderView.Interactive)
            
            header.resizeSection(column, int(width[column]*wfactor))

    @pyqtSlot()
    def search_finish(self):
        self.refreshMainTable()
        self.overlay.setVisible(False)
        
        ## 실행시간 표시
        global start_time
        end_time = time.time() ##계산완료시간
        time_consume = end_time - start_time
        time_consume = '%0.2f' % time_consume  ##소수점2째자리이하는 버림
        self.lineEdit_time.setText(str(time_consume) + "초 소요")
             
if __name__ == "__main__":
    import sys

    app = QApplication(sys.argv)
    dial = MyDialog()
    dial.show()           
    sys.exit(app.exec_())