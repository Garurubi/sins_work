import re
import pandas as pd
import itertools
from PyQt5.QtWidgets import (
    QApplication, QWidget, QLabel, QVBoxLayout, QHBoxLayout,
    QPushButton, QTextBrowser, QTextEdit, QLineEdit
)
from openpyxl import load_workbook

class MyApp(QWidget):

    def __init__(self):
      super().__init__()
      self.initUI()

    def initUI(self):
        self.setWindowTitle('멀티크로 면적 오차 계산')
        # self.setWindowIcon(QIcon('resources/faviconV2.png'))
        self.setGeometry(1000, 200, 800, 300)
        #드래그 드롭을 활성화하려면 True로 변경할 것!
        # self.setAcceptDrops(True)

        #레이아웃 설정
        self.h1_layout = QHBoxLayout()
        self.h1_layout.addStretch(3)
        self.h2_layout = QHBoxLayout()
        self.h3_layout = QHBoxLayout()
        self.v_layout = QVBoxLayout()
        self.v_layout.addLayout(self.h1_layout)
        self.v_layout.addLayout(self.h2_layout)

        self.label = QLabel("오차(%)", self)
        self.h1_layout.addWidget(self.label)
        self.line = QLineEdit()
        self.line.setText("5")
        self.h1_layout.addWidget(self.line)

        # 계산 버튼
        self.check_btn = QPushButton("계산", self)
        # self.check_btn.move(350, 250)
        self.check_btn.clicked.connect(self.loadData)
        self.h1_layout.addWidget(self.check_btn)

        self.te = QTextEdit()
        self.te.setAcceptRichText(False)
        # self.te.setReadOnly(False)
        # self.te.setFocus()
        self.te.setFixedSize(500, 300)
        # self.te.setStyleSheet("color: black; background-color: white;")
        self.h2_layout.addWidget(self.te)

        self.tb = QTextBrowser()
        self.tb.setAcceptRichText(True)
        self.tb.setOpenExternalLinks(True)
        # self.tb.setFixedSize(300, 300)
        # self.tb.move(100, 150)
        self.h2_layout.addWidget(self.tb)

        self.setLayout(self.v_layout)
        self.show()

    # def dragEnterEvent(self, event):
    #     if event.mimeData().hasUrls():
    #         event.accept()
    #     else:
    #         event.ignore()

    # def dropEvent(self, event):
    #     labelword = ""
    #     files = [u.toLocalFile() for u in event.mimeData().urls()]
    #     for f in files:
    #         print(f)
    #         labelword += route_first + f

    #     self.label.setText(labelword)

    # 중복 제거된 조합 찾기
    def find_combination(self, standard, unit_num):
        index_list = self.index_list

        remaining_numbers = [num for num in index_list if num != standard]
        combinations = list(itertools.combinations(remaining_numbers, unit_num))
        
        return combinations

    # '2402', 0.05, 3
    def calculate_error(self, standard, error, unit_num):
        combination_set = self.find_combination(standard, unit_num)

        output = {}
        for set in combination_set:
            temp_dict = {}
            for k, df in self.data.items():
                std_val = df.loc[standard]['area']
                avg = df.loc[[x for x in set]]['area'].mean()
                if (avg > std_val * (1+error)) or (avg < std_val * (1-error)) : break
                temp_dict[k] = (avg/std_val - 1) * 100
            if len(self.data) == len(temp_dict) : output[','.join([x for x in set])] = temp_dict
 
        return output

    def calculate(self):
        # 오차 제대로 기입되었는지 검증
        error_regex = re.compile(r'^(100(\.0{1,2})?|[1-9][0-9]?(\.\d{1,2})?|0(\.\d{1,2})?)$')
        error_text = self.line.text()
        if not error_regex.match(error_text):
           self.tb.setText("오차에 1~100의 값을 입력해주세요.")
           return

        error = float(error_text)/100

        result_dict = {}
        for standard in self.index_list:
            output = self.calculate_error(standard, error, 3)
            if output:
                for k, v in output.items():
                    result_dict[f"{standard} <=> ({k})"] = v

        # output 출력 텍스트 형식
        str_list = []
        str_list.append(f"오차범위: ±{error_text}%")
        if result_dict :
            # output 출력 텍스트 형식
            for comb, v in result_dict.items():
                str_list.append(f"\n{comb}")
                for k, error in v.items():
                    str_list.append(f"{k} : {error:.2f}%")
        else :
            str_list.append("없음")
        
        self.tb.clear()
        self.tb.setText("\n".join(str_list))


    def loadData(self):
        text = self.te.toPlainText()
        row_list = text.split("\n")[2:]

        # 정규 표현식
        material_regex = re.compile(r'^성분:\s')
        numbering_regex = re.compile(r'^\d{4}\s\d{4}-\d{2}-\d{2}')

        material_name = None
        data = {}
        index_list = set()
        try:
            for row in row_list:
                if row == "" : continue
                col_list = row.split("\t")
                A_col = col_list[0]
                D_col = col_list[3]
                if material_regex.match(A_col):
                    material_name = A_col[4:].strip()
                    data[material_name] = pd.DataFrame(columns=['area'])
                elif numbering_regex.match(A_col):
                    numbering = A_col[:4].strip()
                    data[material_name].loc[numbering] = float(D_col)
                    index_list.add(numbering)
        except:
            self.tb.setText("데이터 형식이 맞지 않습니다.")
            return

        self.data = data
        self.index_list = list(index_list)
        self.calculate()


if __name__ == '__main__':
  app = QApplication([])
  myapp = MyApp()
  app.exec_()