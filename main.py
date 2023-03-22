import os
import sys
import math
import enum
import webbrowser
import PyQt5
import shutil
import pandas as pd
from os import path, listdir
#from pathlib import Path
from functools import partial
from datetime import datetime

import csv_checker
import erp_lodix_translator

from openpyxl import load_workbook
from PyQt5 import QtCore, QtWidgets
from PyQt5.QtGui import *
from PyQt5.QtWidgets import (QApplication, QWidget, QFormLayout, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QPushButton, QTextBrowser)



class OrderDropButton(QPushButton):

    def __init__(self, title, parent):
        super().__init__(title, parent)

        self.parent = parent
        self.setAcceptDrops(True)

        self.kt_order_file = ""
        self.mns_order_file = ""

    def dragEnterEvent(self, e):
        tmp_file_path = e.mimeData().urls()[0].toString()
        tmp_file_extension = tmp_file_path[-4:]

        tmp_file_segment = tmp_file_path.split("/")
        # print(tmp_file_path, tmp_file_segment)
        tmp_file_name = tmp_file_segment[-1]
        segment_len = len(tmp_file_name.split("_"))

        if segment_len == 4:
            company_name = tmp_file_name[:-4].split('_')[3]
        else:
            company_name = None
        print(tmp_file_name, tmp_file_name.split("_"), segment_len)

        
        if tmp_file_extension == ".csv":
            if segment_len == 4:
                if company_name.upper() == "KT" or company_name.upper() == "MNS":
                    e.accept()
                else:
                    e.ignore()
            elif segment_len == 3:
                e.accept()
            else:
                e.ignore()
        else:
            e.ignore()

    def dropEvent(self, e):
        self.parent.log_browser.clear()
        self.parent.translate_btn.setDisabled(True)

        self.parent.integrated_file_path = None
        self.parent.kt_file_path = None
        self.parent.mns_file_path = None
        self.curr_file_path = None
        
        if os.name == 'nt':
            tmp_file_path = e.mimeData().urls()[0].toString()
            file_prefix = tmp_file_path[:8]
            tmp_file_path = tmp_file_path[8:]
            tmp_file_name = tmp_file_path.split('/')[-1]
            tmp_file_path = '/'.join(tmp_file_path.split('/')[0:-1]) + '/'
            print(tmp_file_path)
            self.parent.curr_file_path = tmp_file_path

        else:
            tmp_file_path = e.mimeData().urls()[0].toString()
            tmp_file_name = tmp_file_path.split('/')[-1]
            tmp_file_path = '/'.join(tmp_file_path.split('/')[0:-1]) + '/'
            self.parent.curr_file_path = tmp_file_path

        is_file_exists = False
        is_ready_to_process = False
        
        file_type = -1
        # -1: None or error
        #  1: integrated file
        #  2: separated file

        #try:
        file_nm_split = tmp_file_name.split('_')

        # Check is separated file or integrated file
        if len(file_nm_split) == 4:
            # is separated
            file_type = 2
        elif len(file_nm_split) == 3:
            file_type = 1

        center_id = file_nm_split[0]
        cluster_id = file_nm_split[1]
        order_gen_date = file_nm_split[2]

        if file_type == 1:
            # Check error
            if os.name == "nt":
                integrated_file_path = (tmp_file_path + tmp_file_name)
            else:
                integrated_file_path = (tmp_file_path + tmp_file_name)[7:]
            is_integrated_file_valid, integrated_error_dict = csv_checker.check_integrity_of_order_file(integrated_file_path)

            if is_integrated_file_valid:
                # Update info (updated date, order date, and cluster name)
                self.parent.file_type_line_edit.setText('KT & MnS 통합된 주문')
                self.parent.input_type = 1
                self.parent.integrated_file_path = integrated_file_path
                self.parent.order_gen_date_line_edit.setText(order_gen_date.split('.')[0])
                self.parent.order_cluster_line_edit.setText(str(cluster_id))
                self.parent.order_center_line_edit.setText(str(center_id))

                # Get unique list of delivery date
                integrated_df = pd.read_csv(integrated_file_path, quotechar="'")

                integrated_delivery_date_set = set(integrated_df['ORDER_DATE'].tolist())

                #print(len(integrated_delivery_date_set), list(integrated_delivery_date_set))

                if len(integrated_delivery_date_set) == 1:
                    # Order date is valid
                    tmp_delivery_date = list(integrated_delivery_date_set)[0]
                    self.parent.order_delivery_date_line_edit.setText(str(tmp_delivery_date))

                    #tmp_str_to_append = "파일 추가 완료: " + tmp_file_name
                    tmp_str_to_append = "파일 추가 완료: " + tmp_file_name
                    self.parent.log_browser.append(tmp_str_to_append)

                    is_ready_to_process = True
                else:
                    # Error!!!
                    self.parent.log_browser.append("[파일 내용 내 오류 발생]")
                    self.parent.log_browser.append(" - 배송 일자가 다른 주문건 존재")

                    # Get unique list of delivery date

                    pass
            else:
                self.parent.log_browser.append("[파일 내용 내 오류 발생]")
                #self.parent.log_browser.append("*** 통합 주문 파일 내 오류 ***")

                for line_idx, tmp_err_dict in integrated_error_dict.items():
                    error_line_idx = line_idx
                    error_contents = ""
                    for error_key, error_description in tmp_err_dict.items():
                        if len(error_contents) > 0:
                            error_contents = error_contents + " / " + error_description[0]
                        else:
                            error_contents = error_description[0]
                    self.parent.log_browser.append(" - {0}번째 줄 - ".format(int(error_line_idx)) + error_contents)


        elif file_type == 2:
            order_company_type = file_nm_split[3].split('.')[0]

            if order_company_type.lower() == "KT".lower():
                # KT order, load 
                kt_file_name = tmp_file_name
                mns_file_name = tmp_file_name[:-6] + "MnS.csv"
                kt_file_path = tmp_file_path + kt_file_name
                mns_file_path = tmp_file_path + mns_file_name

                tmp_name = ""

                if os.name== "nt":
                    for name in listdir(tmp_file_path):
                        if mns_file_name.lower() == name.lower():
                            tmp_name = name
                            mns_file_path = tmp_file_path + tmp_name

                            print(mns_file_path)
                else:
                    for name in listdir(tmp_file_path[7:]):
                        if mns_file_name.lower() == name.lower():
                            tmp_name = name
                            mns_file_path = tmp_file_path + tmp_name
                
                if tmp_name == "":
                    # print("No MnS file exists!!!")
                    self.parent.log_browser.append("오류 - " + order_gen_date + "일 {0}권역 MnS 주문 파일 존재하지 않음!!".format(cluster_id))
                
                # print("MnS: " + mns_file_path)

                if os.name== "nt":
                    if path.isfile(mns_file_path):
                        is_file_exists = True
                    else:
                        # Set error message at label
                        pass
                else:
                    if path.isfile(mns_file_path[7:]):
                        is_file_exists = True
                    else:
                        # Set error message at label
                        pass
            elif order_company_type.lower() == "MnS".lower():
                mns_file_name = tmp_file_name
                kt_file_name = tmp_file_name[:-7] + "KT.csv"
                kt_file_path = tmp_file_path + kt_file_name
                mns_file_path = tmp_file_path + mns_file_name

                tmp_name = ""

                if os.name == "nt":
                    for name in listdir(tmp_file_path):
                        if kt_file_name.lower() == name.lower():
                            tmp_name = name
                            kt_file_path = tmp_file_path + tmp_name
                else:
                    for name in listdir(tmp_file_path[7:]):
                        if kt_file_name.lower() == name.lower():
                            tmp_name = name
                            kt_file_path = tmp_file_path + tmp_name
                
                if tmp_name == "":
                    #print("No KT file exists!!!")
                    self.parent.log_browser.append("오류 - " + order_gen_date + "일 {0}권역 KT 주문 파일 존재하지 않음!!".format(cluster_id))
                
                # print("KT: " + kt_file_path)

                if os.name == "nt":
                    if path.isfile(kt_file_path):
                        is_file_exists = True
                    else:
                        # Set error message at label
                        pass
                else:
                    if path.isfile(kt_file_path[7:]):
                        is_file_exists = True
                    else:
                        # Set error message at label
                        pass

            if os.name == "nt":
                kt_file_path = kt_file_path
                mns_file_path = mns_file_path
            else:
                kt_file_path = kt_file_path[7:]
                mns_file_path = mns_file_path[7:]


            # print(is_file_exists)
            # print("KT:" + kt_file_path)
            # print("MnS:" + mns_file_path)

            # Check file existance
            if is_file_exists:
                is_kt_valid, kt_error_dict = csv_checker.check_integrity_of_order_file(kt_file_path)
                is_mns_valid, mns_error_dict = csv_checker.check_integrity_of_order_file(mns_file_path)

                if is_kt_valid and is_mns_valid:
                    self.kt_order_file = kt_file_path
                    self.mns_order_file = mns_file_path

                    self.parent.file_type_line_edit.setText('KT & MnS별 분리된 주문')
                    self.parent.input_type = 2
                    self.parent.kt_file_path = kt_file_path
                    self.parent.mns_file_path = mns_file_path
                    self.parent.order_gen_date_line_edit.setText(order_gen_date.split('.')[0])
                    self.parent.order_cluster_line_edit.setText(str(cluster_id))
                    self.parent.order_center_line_edit.setText(str(center_id))

                    # Get unique list of delivery date
                    kt_df = pd.read_csv(kt_file_path, quotechar="'")
                    mns_df = pd.read_csv(mns_file_path, quotechar="'")

                    kt_delivery_date_set = set(kt_df['ORDER_DATE'].tolist())
                    mns_delivery_date_set = set(mns_df['ORDER_DATE'].tolist())

                    integrated_delivery_date_set = kt_delivery_date_set.union(mns_delivery_date_set)

                    #print(len(integrated_delivery_date_set), list(integrated_delivery_date_set))

                    if len(integrated_delivery_date_set) == 1:
                        # Order date is valid
                        tmp_delivery_date = list(integrated_delivery_date_set)[0]
                        self.parent.order_delivery_date_line_edit.setText(str(tmp_delivery_date))

                        tmp_str_to_append = "파일 추가 완료: " + tmp_file_name + ", " + mns_file_name
                        self.parent.log_browser.append(tmp_str_to_append)

                        is_ready_to_process = True

                else:
                    self.parent.log_browser.append("[파일 내용 내 오류 발생]")

                    if len(kt_error_dict.keys()) > 0:
                        # KT 주문 파일 내 오류
                        self.parent.log_browser.append("*** KT 주문 파일 내 오류 ***")

                        for line_idx, tmp_err_dict in kt_error_dict.items():
                            error_line_idx = line_idx
                            error_contents = ""
                            for error_key, error_description in tmp_err_dict.items():
                                if len(error_contents) > 0:
                                    error_contents = error_contents + " / " + error_description[0]
                                else:
                                    error_contents = error_description[0]
                            self.parent.log_browser.append(" - {0}번째 줄 - ".format(int(error_line_idx)) + error_contents)

                    if len(mns_error_dict.keys()) > 0:
                        # MnS 주문 파일 내 오류
                        self.parent.log_browser.append("*** MnS 주문 파일 내 오류 ***")

                        for line_idx, tmp_err_dict in mns_error_dict.items():
                            error_line_idx = line_idx
                            error_contents = ""
                            for error_key, error_description in tmp_err_dict.items():
                                if len(error_contents) > 0:
                                    error_contents = error_contents + " / " + error_description[0]
                                else:
                                    error_contents = error_description[0]
                            self.parent.log_browser.append(" - {0}번째 줄 - ".format(int(error_line_idx)) + error_contents)


        if is_ready_to_process:
            self.parent.translate_btn.setDisabled(False)
            self.parent.raw_file_status_label.setText("  > 상태: 파일 등록 완료")
        else:
            self.parent.translate_btn.setDisabled(True)


class Example(QWidget):

    def __init__(self):
        super().__init__()

        # Declare input file type to determine how to translate
        self.input_type = -1

        self.integrated_file_path = None
        self.kt_file_path = None
        self.mns_file_path = None
        self.curr_file_path = None

        self.given_order_info = [0, 0, 0]


        # Set geometry of window
        self.setGeometry(100, 100, 600, 400)

        # Initialize UI
        self.initUI()

    def initUI(self):
        t_bold_font = PyQt5.QtGui.QFont()
        t_bold_font.setBold(True)

        main_v_box = QVBoxLayout()

        # STEP 1 - login page
        step_1_v_box = QVBoxLayout()
        step_1_1_h_box = QHBoxLayout()
        step_1_1_1_v_box = QVBoxLayout()
        step_1_label = QLabel("Step 1: LODIX 페이지 로그인 (https://lodix.co.kr)")
        step_1_label.setFont(t_bold_font)
        step_1_description_label = QLabel("  > 아래 버튼을 누른 후 LODIX 페이지에서 로그인 해주세요")
        step_1_1_1_v_box.addWidget(step_1_label)
        step_1_1_1_v_box.addWidget(step_1_description_label)
        step_1_1_h_box.addLayout(step_1_1_1_v_box)
        step_1_button = QPushButton('사이트 접속', self)
        step_1_button.clicked.connect(partial(self.open_webbrowser, "lodix_button"))
        step_1_button.setMinimumSize(250, 50)
        step_1_1_h_box.addWidget(step_1_button, alignment=QtCore.Qt.AlignRight)
        step_1_v_box.setContentsMargins(0, 0, 0, 20)  # left, top, right, bottom
        step_1_v_box.addLayout(step_1_1_h_box)
        main_v_box.addLayout(step_1_v_box)

        # STEP 2 - raw order file download
        step_2_v_box = QVBoxLayout()
        step_2_1_h_box = QHBoxLayout()
        step_2_1_1_v_box = QVBoxLayout()
        step_2_label = QLabel("Step 2: 주문 원본 파일 다운")
        step_2_label.setFont(t_bold_font)
        step_2_description_label = QLabel("  > 주문의 생성 날짜를 확인 후 다운로드 받아주세요")
        step_2_description_label_2 = QLabel("  > 생성 일자는 일반적으로 운송 일자의 직전 근무일입니다.")
        step_2_1_1_v_box.addWidget(step_2_label, alignment=QtCore.Qt.AlignLeft)
        step_2_1_1_v_box.addWidget(step_2_description_label, alignment=QtCore.Qt.AlignLeft)
        step_2_1_1_v_box.addWidget(step_2_description_label_2, alignment=QtCore.Qt.AlignLeft)
        step_2_2_h_box = QHBoxLayout()
        order_date_label = QLabel("  > 주문 생성 일자:")
        # order_date_label.setText('')
        self.step_2_order_date_text_box = QLineEdit()
        self.step_2_order_date_text_box.setInputMask("00000000")
        date_today_str = datetime.today().strftime('%Y%m%d')
        self.step_2_order_date_text_box.setText(date_today_str)
        step_2_button = QPushButton('주문파일 다운로드', self)
        step_2_button.clicked.connect(partial(self.open_webbrowser, "merged_order_file_down"))
        step_2_button.setMinimumSize(250, 50)
        # step_2_1_button = QPushButton('병합 주문 다운', self)
        # step_2_1_button.clicked.connect(partial(self.open_webbrowser, "merged_order_file_down"))
        # step_2_2_button = QPushButton('분리 주문 다운', self)
        # step_2_2_button.clicked.connect(partial(self.open_webbrowser, "separated_order_file_down"))
        step_2_2_h_box.addWidget(order_date_label, alignment=QtCore.Qt.AlignLeft)
        step_2_2_h_box.addWidget(self.step_2_order_date_text_box, alignment=QtCore.Qt.AlignLeft)
        # step_2_1_h_box.addWidget(step_2_1_button, alignment=QtCore.Qt.AlignRight)
        # step_2_1_h_box.addWidget(step_2_2_button, alignment=QtCore.Qt.AlignRight)
        step_2_v_box.setContentsMargins(0, 0, 0, 20)  # left, top, right, bottom
        step_2_1_1_v_box.addLayout(step_2_2_h_box)
        step_2_1_h_box.addLayout(step_2_1_1_v_box)
        step_2_1_h_box.addWidget(step_2_button, alignment=QtCore.Qt.AlignRight)
        step_2_v_box.addLayout(step_2_1_h_box)
        main_v_box.addLayout(step_2_v_box)

        # STEP 3 - register raw order file
        step_3_v_box = QVBoxLayout()
        step_3_label = QLabel("Step 3: 주문 원본 파일 등록")
        step_3_label.setMaximumHeight(50)
        step_3_label.setFont(t_bold_font)
        step_3_description_label = QLabel("  > 주문 원본 파일(csv)을 하단의 주문버튼에 드래그&드랍해주세요")
        step_3_description_label.setMaximumHeight(50)
        step_3_v_box.addWidget(step_3_label, alignment=QtCore.Qt.AlignLeft)
        step_3_v_box.addWidget(step_3_description_label, alignment=QtCore.Qt.AlignLeft)

        h_box_0 = QHBoxLayout()
        v_box_0_1 = QVBoxLayout()
        # v_box_0_1.setContentsMargins(0, 0, 20, 0)  # left, top, right, bottom
        v_box_0_2 = QVBoxLayout()
        # v_box_0_2.setContentsMargins(20, 30, 0, 0)  # left, top, right, bottom
        raw_file_drop_btn = OrderDropButton('주문', self)
        raw_file_drop_btn.setMinimumSize(250, 150)
        v_box_0_1.addWidget(raw_file_drop_btn, alignment=QtCore.Qt.AlignCenter)
        v_box_0_1.setContentsMargins(0, 0, 0, 0)  # left, top, right, bottom
        h_box_0.setContentsMargins(0, 0, 0, 0)  # left, top, right, bottom
        #v_box_0_2.alignment = Qt.AlignTop
        self.file_type_line_edit = QLineEdit()
        self.order_gen_date_line_edit = QLineEdit()
        self.order_delivery_date_line_edit = QLineEdit()
        self.order_center_line_edit = QLineEdit()
        self.order_cluster_line_edit = QLineEdit()
        self.file_type_line_edit.setReadOnly(True)
        self.order_gen_date_line_edit.setReadOnly(True)
        #self.order_delivery_date_line_edit.setReadOnly(True)
        self.onlyInt = QIntValidator()
        self.order_delivery_date_line_edit.setValidator(self.onlyInt)

        self.order_center_line_edit.setReadOnly(True)
        self.order_cluster_line_edit.setReadOnly(True)

        self.translate_btn = QPushButton("주문 통합/병합 실행")

        form_layout_1 = QFormLayout()
        form_layout_1.addRow(QLabel("파일 유형 :"), self.file_type_line_edit)
        form_layout_1.addRow(QLabel("센터 번호 :"), self.order_center_line_edit)
        form_layout_1.addRow(QLabel("권역 번호 :"), self.order_cluster_line_edit)
        form_layout_1.addRow(QLabel("주문생성일:"), self.order_gen_date_line_edit)
        form_layout_1.addRow(QLabel("실제출고일:"), self.order_delivery_date_line_edit)
        form_layout_1.setContentsMargins(0, 0, 0, 0)
        v_box_0_2.addLayout(form_layout_1)

        h_box_0.addLayout(v_box_0_2)
        h_box_0.addLayout(v_box_0_1)
        step_3_v_box.addLayout(h_box_0)
        step_3_v_box.setContentsMargins(0, 0, 0, 20)  # left, top, right, bottom
        main_v_box.addLayout(step_3_v_box)

        # STEP 4 - translate order file to web template
        step_4_v_box = QVBoxLayout()
        step_4_1_h_box = QHBoxLayout()
        step_4_1_1_v_box = QVBoxLayout()
        step_4_label = QLabel("Step 4: 주문 원본 파일 변환")
        step_4_label.setFont(t_bold_font)
        step_4_description_label = QLabel("  > 파일 등록 후 우측 버튼을 눌러 주문을 변환하세요")
        step_4_1_1_v_box.addWidget(step_4_label, alignment=QtCore.Qt.AlignLeft)
        step_4_1_1_v_box.addWidget(step_4_description_label, alignment=QtCore.Qt.AlignLeft)
        self.raw_file_status_label = QLabel("  > 상태: 파일을 등록해주세요")
        self.raw_file_status_label.setFont(t_bold_font)
        step_4_1_1_v_box.addWidget(self.raw_file_status_label, alignment=QtCore.Qt.AlignLeft)
        step_4_1_h_box.addLayout(step_4_1_1_v_box)
        # self.translate_btn.setContentsMargins(20, 20, 20, 20)
        # self.translate_btn.setMinimumHeight(70)
        self.translate_btn.setDisabled(True)
        self.translate_btn.clicked.connect(self.do_translate_raw_order)
        self.translate_btn.setMinimumSize(250, 80)
        step_4_1_h_box.addWidget(self.translate_btn, alignment=QtCore.Qt.AlignRight)
        step_4_v_box.addLayout(step_4_1_h_box)
        step_4_v_box.setContentsMargins(0, 0, 0, 20)  # left, top, right, bottom
        main_v_box.addLayout(step_4_v_box)

        h_box_1 = QHBoxLayout()
        v_box_1_1 = QVBoxLayout() # For title and log
        log_title_label = QLabel("처리 내역")
        log_title_label.setMaximumHeight(50)
        log_title_label.setFont(t_bold_font)
        self.log_browser = QTextBrowser()
        self.log_browser.append("대기 중")
        self.log_browser.setMinimumHeight(200)
        self.log_browser.setMaximumHeight(400)
        v_box_1_1.addWidget(log_title_label)
        v_box_1_1.addWidget(self.log_browser)
        h_box_1.addLayout(v_box_1_1)
        main_v_box.addLayout(h_box_1)

        self.setLayout(main_v_box)
        self.setWindowTitle('KT & MnS 주문정보 변환기')
        
        self.show()


    def update_button_text(self, text):
        self.raw_file_status_label.setText(text)


    def open_webbrowser(self, _btn_type):
        if _btn_type == "lodix_button":
            webbrowser.open("https://lodix.co.kr")
        elif _btn_type == "separated_order_file_down":
            date_info = self.step_2_order_date_text_box.text()
            order_file_download_path = "https://lodix.co.kr/nasOrigineFileDown.json?date=" + date_info
            webbrowser.open(order_file_download_path)
            pass
        elif _btn_type == "merged_order_file_down":
            date_info = self.step_2_order_date_text_box.text()
            order_file_download_path = "https://lodix.co.kr/nasMergeFileDown.json?date=" + date_info
            webbrowser.open(order_file_download_path)
            pass
        else:
            pass
    

    def do_translate_raw_order(self):
        self.given_order_info = [int(self.order_center_line_edit.text()), int(self.order_cluster_line_edit.text()), int(self.order_delivery_date_line_edit.text())]

        if self.input_type == 1:
            concatenated_order_df = pd.read_csv(self.integrated_file_path)
        elif self.input_type == 2:
            concatenated_order_df, _ = erp_lodix_translator.integrate_kt_and_mns_order(self.kt_file_path, self.mns_file_path, "'", "'")
        else:
            self.log_browser.append("잘못된 입력!")
            return
        
        #correction_rule_applied_df = erp_lodix_translator.apply_rules_to_integrated_order(concatenated_order_df)
        #erp_lodix_translator.compare_two_dfs(correction_rule_applied_df, concatenated_order_df)
        
        excel_df = erp_lodix_translator.convert_etl_format_to_excel_format(concatenated_order_df)
        #excel_df = erp_lodix_translator.convert_etl_format_to_excel_format(correction_rule_applied_df)

        if os.name == "nt":
            destination_path = self.curr_file_path
        else:
            destination_path = self.curr_file_path[7:]
        result_file_name = "{0:04d}_{1:02d}_{2:08d}_web_format.xlsx".format(int(self.given_order_info[0]), int(self.given_order_info[1]), int(self.given_order_info[2]))

        result_file_full_path = destination_path + result_file_name

        erp_lodix_translator.generate_excel_for_web_upload(excel_df, result_file_full_path, "v2")

        self.log_browser.append("웹 업로드용 주문 파일 생성 완료 - " + result_file_name)

        self.translate_btn.setDisabled(True)
        self.raw_file_status_label.setText("  > 상태: 대기 중")



if __name__ == '__main__':
     app = QApplication(sys.argv)
     ex = Example()
     sys.exit(app.exec_())    


