import re
import pandas as pd
import itertools
from PyQt5.QtWidgets import (
    QApplication, QWidget, QLabel, QVBoxLayout, QHBoxLayout,
    QPushButton, QTextBrowser, QTextEdit, QLineEdit, QScrollArea
)

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
        self.v_layout = QVBoxLayout()
        self.h1_layout = QHBoxLayout()
        self.h1_layout.addStretch(3)
        self.h2_layout = QHBoxLayout()
        self.v_layout.addLayout(self.h1_layout)
        self.v_layout.addLayout(self.h2_layout)

        self.label = QLabel("오차(%)", self)
        self.h1_layout.addWidget(self.label)
        self.line = QLineEdit()
        self.line.setText("5")
        self.h1_layout.addWidget(self.line)

        # 계산 버튼
        self.check_btn = QPushButton("계산", self)
        self.check_btn.clicked.connect(self.loadData)
        self.h1_layout.addWidget(self.check_btn)

        self.te = QTextEdit()
        self.te.setAcceptRichText(False)
        self.te.setFixedSize(500, 300)
        self.h2_layout.addWidget(self.te)

        # 출력 결과 만들 레이어(스크롤 추가)
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)

        self.scroll_content = QWidget()
        self.scroll_layout = QVBoxLayout(self.scroll_content)
        self.scroll_area.setWidget(self.scroll_content)
        self.h2_layout.addWidget(self.scroll_area)

        self.setLayout(self.v_layout)
        self.show()

    # 해당 면적값 복사
    def copy_clipboard(self, standard, comb):
        comb_list = [x for x in comb.split(",")]

        text_list = []
        for k, df in self.data.items():
            for comb in comb_list:
                text_list.append(str(df.loc[comb]['area']))
            text_list.append(str(df.loc[standard]['area']))

        # 클립보드에 텍스트 복사
        clipboard = QApplication.clipboard()
        clipboard.setText("\n".join(text_list))

    # 중복 제거된 조합 찾기
    def find_combination(self, standard, unit_num):
        index_list = self.index_list

        remaining_numbers = [num for num in index_list if num != standard]
        combinations = list(itertools.combinations(remaining_numbers, unit_num))
        
        return [tuple(sorted(comb)) for comb in combinations]

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
                    result_dict[f"{standard}<=>({k})"] = v

        # output 출력 텍스트 형식
        str_list = []
        # str_list.append(f"오차범위: ±{error_text}%")
        if result_dict :
            # output 출력 텍스트 형식
            for comb, v in result_dict.items():
                temp_list = []
                temp_list.append(f"{comb}")
                for k, error in v.items():
                    temp_list.append(f"{k} : {error:.2f}%")
                str_list.append(temp_list)
        
        # 이전에 만든 레이아웃 지우기
        self.scroll_content.deleteLater()

        self.scroll_content = QWidget()
        self.scroll_layout = QVBoxLayout(self.scroll_content)
        
        # 오차내에 들어오는 값이 하나도 없는 경우
        if len(str_list) == 0 :
            qh = QHBoxLayout()
            tb = QTextBrowser()
            tb.setText("없음")
            qh.addWidget(tb)
            self.scroll_layout.addLayout(qh)
        else :
            for i, text_list in enumerate(str_list):
                qh = QHBoxLayout()
                tb = QTextBrowser()
                tb.setAcceptRichText(True)
                tb.setText("\n".join(text_list))
                qh.addWidget(tb)

                standard, comb = text_list[0].split("<=>", 1)
                btn = QPushButton("복사", self)
                btn.setObjectName(f'button{i}')
                btn.clicked.connect(lambda _, s=standard, c=comb: self.copy_clipboard(s, c.replace("(", "").replace(")", "")))
                qh.addWidget(btn)

                self.scroll_layout.addLayout(qh)

        self.scroll_area.setWidget(self.scroll_content)
        self.scroll_area.update()

    def loadData(self):
        text = self.te.toPlainText()
        row_list = text.split("\n")

        column_names = row_list[1]
        area_col_num = [i for i, name in enumerate(column_names.split("\t")) if name.strip() == "면적"]
        if len(area_col_num) != 1:
            self.tb.setText("면적 컬럼은 반드시 한개가 있어야 합니다.")
            return

        # 정규 표현식
        material_regex = re.compile(r'^성분:\s')
        numbering_regex = re.compile(r'^\d{4}\s\d{4}-\d{2}-\d{2}')

        material_name = None
        data = {}
        index_list = set()
        try:
            for row in row_list[2:]:
                if row == "" : continue
                col_list = row.split("\t")
                A_col = col_list[0]
                D_col = col_list[area_col_num[0]]
                if material_regex.match(A_col):
                    material_name = A_col[4:].strip()
                    data[material_name] = pd.DataFrame(columns=['area'])
                elif material_name and numbering_regex.match(A_col):
                    if not D_col : 
                        del data[material_name]
                        material_name = None
                        continue
                    numbering = A_col[:4].strip()
                    data[material_name].loc[numbering] = int(D_col)
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