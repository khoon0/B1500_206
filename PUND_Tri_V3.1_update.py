import sys
import csv
import os
from datetime import datetime
from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog, QVBoxLayout, QWidget, QLabel, QInputDialog, QMessageBox, QProgressBar, QHBoxLayout
from PyQt5.QtGui import QFont
from PyQt5.QtCore import Qt, QPropertyAnimation, QRect
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
import requests

class CSVFileProcessor(QMainWindow):
    def __init__(self):
        super().__init__()
        self.swap_current_voltage = False  # 초기화 추가
        self.current_version = '3.0.0'  # 현재 버전 설정
        self.latest_version = None  # 최신 버전 저장
        self.initUI()  # UI 초기화
        self.check_for_update()  # 업데이트 체크 호출

    def initUI(self):
        self.setWindowTitle('PUND auto calculator (Keysight B1500)')
        self.setGeometry(400, 150, 1024, 768)

        # 메인 레이아웃 생성
        main_layout = QVBoxLayout()

        # 드래그 앤 드롭 안내 문구 추가
        label = QLabel(self.create_label_text())
        label.setTextFormat(Qt.RichText)
        label.setTextInteractionFlags(Qt.TextBrowserInteraction)
        label.setOpenExternalLinks(True)
        label.mousePressEvent = self.select_files  # 마우스 클릭 이벤트 연결

        # 폰트 설정
        label.setFont(self.create_font())
        label.setAlignment(Qt.AlignCenter)  # 가운데 정렬
        main_layout.addWidget(label)

        # 버전 정보 레이아웃 추가
        version_layout = QHBoxLayout()
        version_label = QLabel(f"Version: {self.current_version}")
        version_label.setAlignment(Qt.AlignRight)  # 우측 정렬
        version_label.setFont(QFont("Arial", 10))  # 10포인트 폰트 설정
        version_layout.addStretch()  # 여백 추가하여 우측 정렬 유지
        version_layout.addWidget(version_label)

        # 버전 레이아웃의 높이를 줄이기 위해 최소 높이 설정
        version_layout.setContentsMargins(0, 0, 0, 0)  # 여백 제거
        version_label.setFixedHeight(15)  # 고정 높이 설정

        main_layout.addLayout(version_layout)  # 메인 레이아웃에 버전 레이아웃 추가

        self.central_widget = QWidget()
        self.central_widget.setLayout(main_layout)
        self.setCentralWidget(self.central_widget)

        # 드래그 앤 드롭 이벤트 설정
        self.setAcceptDrops(True)

    def create_label_text(self):
        return ("Click <a href='#'>here</a> to select .csv files.<br><br>"
                "(or, Drag and Drop the .csv files)<br><br><br>"
                "@Made by. Kihoon Kim@")

    def create_font(self):
        font = QFont()
        font.setPointSize(45)
        font.setBold(True)
        return font
    
    def check_for_update(self):
        try:
            response = requests.get('https://raw.githubusercontent.com/khoon0/B1500-calculator/master/latest_version.txt')  # 최신 버전 정보 요청
            self.latest_version = response.text.strip()
            if self.is_newer_version(self.latest_version, self.current_version):
                self.prompt_update()
            else:
                self.show_message("현재 버전이 최신입니다.", "업데이트 확인")
        except Exception as e:
            self.show_message(f"업데이트 체크 중 오류 발생: {e}", "오류")

    def is_newer_version(self, latest, current):
        # 버전 비교 로직
        latest_parts = list(map(int, latest.split('.')))
        current_parts = list(map(int, current.split('.')))
        return latest_parts > current_parts

    def prompt_update(self):
        # 업데이트가 필요하다는 메시지 출력
        self.show_message("새로운 업데이트가 있습니다. 다운로드 중...", "업데이트 알림")
        self.download_update()

    def download_update(self):
        try:
            response = requests.get('https://github.com/khoon0/B1500_sdrl/raw/master/PUND_Tri.exe')  # 최신 .exe 파일 다운로드
            with open('PUND.exe', 'wb') as f:  # 기존 파일 덮어쓰기
                f.write(response.content)
            self.current_version = self.latest_version  # 현재 버전 업데이트
            self.show_message("업데이트 완료! 프로그램을 다시 시작하세요.", "업데이트 완료")
            self.restart_program()  # 프로그램 재시작
        except Exception as e:
            self.show_message(f"업데이트 다운로드 중 오류 발생: {e}", "오류")

    def show_message(self, message, title):
        msg_box = QMessageBox()
        msg_box.setText(message)
        msg_box.setWindowTitle(title)
        msg_box.setStandardButtons(QMessageBox.Ok)

        # 팝업창 크기 조정
        msg_box.resize(800, 400)  # 크기를 조정
        msg_box.exec_()

    def restart_program(self):
        """현재 프로그램을 재시작합니다."""
        os.execv(sys.executable, ['python'] + sys.argv)

    def select_files(self, event):
        file_paths, _ = QFileDialog.getOpenFileNames(self, "Select CSV Files", "./", "CSV Files (*.csv)")
        if file_paths:
            file_names = [os.path.basename(path) for path in file_paths]
            self.process_selected_files(file_paths, file_names)

    def dragEnterEvent(self, event):
        event.accept() if event.mimeData().hasUrls() else event.ignore()

    def get_data_order(self):
        order_options = ["'Time', 'Voltage', 'Current'", "'Time', 'Current', 'Voltage'"]
        order, ok = QInputDialog.getItem(self, "Choose Data Order", 
                                          "Choose the right data order in your .csv data files:", 
                                          order_options, 0, False)
        if ok:
            self.swap_current_voltage = (order == "'Time', 'Current', 'Voltage'")
            return True
        QMessageBox.information(self, "Saving Canceled", "Saving file has been canceled.")
        return False
        
    def process_files(self, file_paths, file_names):
        all_data = []
        
        # 사용자에게 측정 방법 선택 창 표시
        measurement_method, ok = QInputDialog.getItem(self, "Select the method to calculate", 
                                                    "Select the method to calculate (1. PUND    2. Triangular Pulse    3. I-V Curve)", 
                                                    ["PUND measurement. P-E curve", "Triangular pulse measurement. Q-E curve", "I-V Curve"], 0, False)
        if not ok:
            # 사용자가 Cancel 버튼을 누르면 전체 프로세스 취소
            return
        
        for file_path, file_name in zip(file_paths, file_names):
            data = None
            if measurement_method == "PUND measurement. P-E curve":
                data = self.process_csv_file_PUND(file_path, file_name)
            elif measurement_method == "Triangular pulse measurement. Q-E curve":
                data = self.process_csv_file_Tri(file_path, file_name)
            elif measurement_method == "I-V Curve":
                data = self.process_csv_file_IV(file_path, file_name)
            else:
                QMessageBox.warning(self, "Error", "Invalid measurement method selected.")
                return
            
            if data is not None:
                all_data.append(data)
        
        if all_data and self.get_device_dimensions(all_data):
            self.save_data(measurement_method, all_data, file_names)

    def save_data(self, measurement_method, all_data, file_names):
        if measurement_method == "PUND measurement. P-E curve":
            self.save_to_single_file_PUND(all_data, file_names)
        elif measurement_method == "Triangular pulse measurement. Q-E curve":
            self.save_to_single_file_Tri(all_data, file_names)
        elif measurement_method == "I-V Curve":
            self.save_to_single_file_IV(all_data, file_names)

    def process_selected_files(self, file_paths, file_names):
        if self.get_data_order():
            self.process_files(file_paths, file_names)

    def dropEvent(self, event):
        if self.get_data_order():
            urls = event.mimeData().urls()
            file_paths = [str(url.toLocalFile()) for url in urls]
            file_names = [os.path.basename(path) for path in file_paths]
            self.process_files(file_paths, file_names)
        else:
            event.ignore()

    def process_csv_file(self, file_path, file_name, measurement_type):
        try:
            with open(file_path, 'r', encoding='utf-8') as file:
                reader = csv.reader(file)
                data = list(reader)

            # 첫 번째 'DataValue' 행 찾기
            start_row = next((i for i, row in enumerate(data) if 'DataValue' in row[0]), None)

            # 'DataValue'가 아닌 첫 번째 행 찾기
            end_row = None
            if start_row is not None:
                end_row = next((i + start_row for i, row in enumerate(data[start_row:]) if 'DataValue' not in row[0]), None)

            # 데이터 행 수 계산
            if start_row is not None and end_row is not None:
                B_data = [float(row[1]) for row in data[start_row:end_row]]
                C_data = [float(row[2]) for row in data[start_row:end_row]]
                D_data = [float(row[3]) * (-1) for row in data[start_row:end_row]]

                # 데이터 순서에 따른 스와핑
                if self.swap_current_voltage:
                    C_data, D_data = D_data, C_data

                # 데이터 가공
                time_interval = B_data[1] - B_data[0]

                if measurement_type == "PUND":
                    num_chunks = 4
                    chunk_size = len(C_data) // num_chunks
                    C_chunks = [C_data[i * chunk_size:(i + 1) * chunk_size] for i in range(num_chunks)]
                    D_chunks = [D_data[i * chunk_size:(i + 1) * chunk_size] for i in range(num_chunks)]

                    D_diff = [d1 - d2 for d1, d2 in zip(D_chunks[0], D_chunks[1])] + [d3 - d4 for d3, d4 in zip(D_chunks[2], D_chunks[3])]

                    # 데이터 프레임 생성
                    df = pd.DataFrame({
                        'Time Interval': [time_interval] * (len(C_chunks[0]) + len(C_chunks[2])),
                        'Voltage_diff': C_chunks[0] + C_chunks[2],
                        'Current_diff': D_diff,
                        ' ': [''] * (len(C_chunks[0]) + len(C_chunks[2])),
                        'Voltage_V': C_chunks[0] + C_chunks[2],
                        'Δ Charge': [(d * time_interval) for d in D_diff],
                        '  ': [''] * (len(C_chunks[0]) + len(C_chunks[2])),
                        'E-field': C_chunks[0] + C_chunks[2],
                        'Polarization': 0
                    })

                elif measurement_type == "Tri":
                    # 데이터 프레임 생성
                    df = pd.DataFrame({
                        'Time Interval': [time_interval] * len(B_data),
                        'Voltage': C_data,
                        'Current': D_data,
                        ' ': [''] * len(B_data),
                        'Voltage_V': C_data,
                        'Δ Charge': [(d * time_interval) for d in D_data],
                        '  ': [''] * len(B_data),
                        'E-field': 0,
                        'Charge': 0
                    })

                elif measurement_type == "IV":
                    # 데이터 프레임 생성
                    df = pd.DataFrame({
                        'Time': B_data,
                        'Voltage': C_data,
                        'Current': D_data
                    })

                return df
            else:
                print(f"CSV 파일 '{file_path}' 에서 'DataValue' 데이터를 찾을 수 없습니다.")
                return None

        except Exception as e:
            print(f"Error processing CSV file: {e}")
            return None

    def process_csv_file_PUND(self, file_path, file_name):
        return self.process_csv_file(file_path, file_name, measurement_type="PUND")

    def process_csv_file_Tri(self, file_path, file_name):
        return self.process_csv_file(file_path, file_name, measurement_type="Tri")

    def process_csv_file_IV(self, file_path, file_name):
        return self.process_csv_file(file_path, file_name, measurement_type="IV")
        
    def get_device_dimensions(self, all_data):
        while True:
            try:
                dimensions, ok = QInputDialog.getText(self, "Device Dimensions", "Enter the device's width(μm), height(μm), and thickness(nm) separated by commas (e.g., 80, 90, 10):")
                if ok:
                    width, height, thickness = map(float, dimensions.split(','))
                    self.device_width = width
                    self.device_height = height
                    self.device_thickness = thickness
                    self.device_area = width * height
                    return True
                else:
                    QMessageBox.information(self, "Saving Canceled", "Saving file has been canceled.")
                    return False
            except ValueError:
                QMessageBox.information(self, "Error input", "Error input. Please enter the device dimensions in correct form (e.g., 80, 90, 10).")
                continue

    def save_to_single_file_PUND(self, all_data, file_names):
        # 현재 날짜와 시간 가져오기
        now = datetime.now()
        timestamp = now.strftime('%y%m%d_%H%M%S')

        # 기본 파일 이름 설정
        default_file_name = f'{timestamp}_PUND_W{self.device_width:.0f}xH{self.device_height:.0f}xt{self.device_thickness:.0f}'

        # 파일 저장 경로 선택 대화 상자 열기
        output_dir = QFileDialog.getExistingDirectory(self, "Select Folder to Save Files", os.path.expanduser("~"))

        # 사용자가 취소 버튼을 누른 경우
        if not output_dir:
            QMessageBox.information(self, "Saving Canceled", "Saving file has been canceled.")
            return

        # 새 폴더 생성
        new_folder_path = os.path.join(output_dir, default_file_name)
        try:
            os.makedirs(new_folder_path)
        except FileExistsError:
            pass

        # Excel 파일 경로 설정
        output_file_path = os.path.join(new_folder_path, f"{default_file_name}.xlsx")

        with pd.ExcelWriter(output_file_path, engine='xlsxwriter') as writer:
            # 그래프 생성을 위한 리스트 초기화
            e_field_list = []
            polarization_list = []

            for i, df in enumerate(all_data):
                if df is not None:
                    # 'Polarization' 데이터 계산
                    charge_data = df['Δ Charge'].cumsum()
                    pol_max = max(charge_data)
                    pol_min = min(charge_data)
                    pol_average = (pol_max + pol_min) / 2
                    df['Polarization'] = [(c - pol_average)/(self.device_area * 1e-14) for c in charge_data]

                    # 'E-field' 계산
                    efield = df['Voltage_V']
                    df['E-field'] = [(v*10/self.device_thickness) for v in efield]
                    
                    # 각 sheet에 'E-field', 'Polarization' 데이터 저장
                    df.to_excel(writer, sheet_name=file_names[i], index=False)

                    # 그래프 데이터 리스트에 추가
                    e_field_list.append(df['E-field'])
                    polarization_list.append(df['Polarization'])
                                
            # 새로운 sheet 생성
            new_df = pd.DataFrame()
            for j, sheet_df in enumerate(all_data):
                if sheet_df is not None:
                    new_df[f'E-field_{j+1}'] = sheet_df['E-field']
                    new_df[f'Polarization_{j+1}'] = sheet_df['Polarization']

            new_df.to_excel(writer, sheet_name='E-field_Polarization', index=False)

            # 'Pr_Ec' 시트 생성
            pr_ec_df = pd.DataFrame()
            for k, sheet_df in enumerate(all_data):
                if sheet_df is not None:
                    #pr_max = max(sheet_df['Polarization'])
                    ec1_values = []
                    ec2_values = []
                    pr1_values = []
                    pr2_values = []
                    for j in range(len(sheet_df['Polarization'])-1):
                        if sheet_df['Polarization'][j] * sheet_df['Polarization'][j+1] < 0:
                            if sheet_df['Polarization'][j] > 0:
                                ec1 = sheet_df['E-field'][j+1]
                                ec1_values.append(ec1)
                                #print(ec1)
                            else:
                                ec2 = sheet_df['E-field'][j+1]
                                ec2_values.append(ec2)
                                #print(ec2)
                        else:
                            ec1_values.append(None)
                            ec2_values.append(None)
                            pr1_values.append(None)
                            pr2_values.append(None)

                        if sheet_df['E-field'][j] * sheet_df['E-field'][j+1] < 0:
                            if sheet_df['E-field'][j] > 0:
                                pr1 = sheet_df['Polarization'][j+1]
                                pr1_values.append(pr1)
                                #print(pr1, "pr1")
                        else:
                            pr1_values.append(None)
                    
                    # 'None' 값을 제거하고 평균 계산
                    ec1_values = [value for value in ec1_values if value is not None]
                    ec2_values = [value for value in ec2_values if value is not None]
                    pr1_values = [value for value in pr1_values if value is not None]
         
                    if pr1_values:
                        pr_ec_df[f'Pr_max_{k+1}'] = [abs(np.mean(pr1_values))]                    
                    else:
                        pr_ec_df[f'Pr_max_{k+1}'] = [None]               
                    if ec1_values and ec2_values:
                        pr_ec_df[f'Ec_{k+1}'] = [(abs(np.mean(ec1_values)) + abs(np.mean(ec2_values))) / 2]
                    else:
                        pr_ec_df[f'Ec_{k+1}'] = [None]

            # 'Pr_Ec' 시트에 데이터 저장
            pr_ec_df.to_excel(writer, sheet_name='Pr_Ec', index=False)

        # 진행 상황 표시를 위한 위젯 생성
        progress_widget = QWidget()
        progress_widget.setWindowTitle("Saving Progress")
        progress_widget.setGeometry(650, 400, 500, 250)

        # 배경색 및 스타일 설정
        progress_widget.setStyleSheet("background-color: #f0f0f0; border: 2px solid #0078d7; border-radius: 10px;")

        # 레이아웃 생성 (전체 레이아웃은 배경 없음)
        progress_layout = QVBoxLayout()
        progress_layout.setContentsMargins(10, 10, 10, 10)  # 여백 설정

        self.progress_bar = QProgressBar()
        self.progress_label = QLabel("Progressing...... 0%")

        # 기본 스타일을 사용하여 진행 바 설정
        self.progress_bar.setTextVisible(True)  # 텍스트 표시
        self.progress_bar.setMinimum(0)          # 최소값
        self.progress_bar.setMaximum(100)        # 최대값

        self.progress_label.setAlignment(Qt.AlignCenter)
        self.progress_bar.setAlignment(Qt.AlignCenter)  

        # 프로그레스 바를 감싸는 레이아웃 생성
        progress_bar_layout = QVBoxLayout()
        progress_bar_layout.addWidget(self.progress_bar)
        progress_layout.addLayout(progress_bar_layout)  # 프로그레스 바 레이아웃 추가

        # 레이아웃에 레이블 추가
        progress_layout.addWidget(self.progress_label)

        # 전체 레이아웃 설정
        progress_widget.setLayout(progress_layout)

        # 애니메이션 효과 추가
        animation = QPropertyAnimation(progress_widget, b"geometry")
        animation.setDuration(500)
        animation.setStartValue(QRect(750, 450, 300, 0))
        animation.setEndValue(QRect(750, 450, 300, 150))
        animation.start()

        progress_widget.show()

        # Matplotlib을 사용하여 그래프 생성 및 저장
        fig = plt.figure(figsize=(10, 16))
        ax = fig.add_subplot(111)

        # 그래프 색상 설정 (그라데이션 느낌)
        colors = plt.cm.Accent(np.linspace(0, 1, len(e_field_list)))

        for i, (e_field, polarization) in enumerate(zip(e_field_list, polarization_list)):
            ax.plot(e_field, polarization, color=colors[i], label=file_names[i], linewidth=5)

            # 진행 상황 업데이트
            progress = int((i + 1) / len(polarization_list) * 50)  # 전체 진행률의 33.3%를 차지
            self.progress_label.setText(f"Progressing...... {int(progress)}%")
            
            # 진행 바 업데이트
            self.progress_bar.setValue(progress)
            QApplication.processEvents()  # UI 업데이트

        # 축 라벨 폰트 크기 및 스타일 변경
        ax.set_xlabel('E-field [MV/cm]', fontsize=28, fontweight='bold')
        ax.set_ylabel('Polarization [μC/cm²]', fontsize=28, fontweight='bold')

        # x축 값 폰트 스타일 변경
        ax.set_xticklabels(ax.get_xticks(), fontsize=14, fontweight='bold')
        ax.set_yticklabels(ax.get_xticks(), fontsize=14, fontweight='bold')

        # x축 눈금 값의 유효숫자 조정
        ax.xaxis.set_major_formatter(ticker.FormatStrFormatter('%.0f'))

        # y축 눈금 값의 유효숫자 조정
        ax.yaxis.set_major_formatter(ticker.FormatStrFormatter('%.0f'))

        # 눈금 라벨 폰트 크기 변경
        ax.tick_params(axis='both', which='major', labelsize=22, direction='in')

        # 제목 폰트 크기 및 스타일 변경
        ax.set_title(f'{default_file_name}', fontsize=16, fontweight='bold')

        # 축 라벨 위치 조정
        ax.xaxis.set_label_coords(0.5, -0.1)
        ax.yaxis.set_label_coords(-0.1, 0.5)

        ax.legend(loc='upper left', prop={'weight':'bold', 'size':12}, borderaxespad=0.1, frameon=False)

        # 그래프 틀의 선 두께 변경
        ax.spines['left'].set_linewidth(2)
        ax.spines['right'].set_linewidth(2)
        ax.spines['top'].set_linewidth(2)
        ax.spines['bottom'].set_linewidth(2)

        # 그리드 추가 (점선 및 minor grid)
        ax.grid(True, which='both', linestyle='--')
        ax.minorticks_on()
        ax.grid(True, which='minor', linestyle=':')

        # Pr_max와 Ec 값 표시 (bold체)
        pr_max_list = []
        ec_list = []
        for i, (e_field, polarization) in enumerate(zip(e_field_list, polarization_list)):
            ec1_values = []
            ec2_values = []
            pr1_values = []
            for j in range(len(polarization)-1):
                if polarization[j] * polarization[j+1] < 0:
                    if polarization[j] > 0:
                        ec1 = e_field[j+1]
                        ec1_values.append(ec1)
                    else:
                        ec2 = e_field[j+1]
                        ec2_values.append(ec2)
                if e_field[j] * e_field[j+1] < 0:
                    if e_field[j] > 0:
                        pr1 = polarization[j+1]
                        pr1_values.append(pr1)

            # 'None' 값을 제거하고 평균 계산
            ec1_values = [value for value in ec1_values if value is not None]
            ec2_values = [value for value in ec2_values if value is not None]
            pr1_values = [value for value in pr1_values if value is not None]

            if pr1_values:
                pr_max = abs(np.mean(pr1_values))
            else:
                pr_max = None
            if ec1_values and ec2_values:
                ec = (abs(np.mean(ec1_values)) + abs(np.mean(ec2_values))) / 2
            else:
                ec = None

            pr_max_list.append(pr_max)
            ec_list.append(ec)

            # 진행 상황 업데이트
            progress = int(50 + (i + 1) / len(polarization_list) * 50)  # 전체 진행률의 33.3%를 차지
            self.progress_label.setText(f"Progressing...... {int(progress)-1}%")
            
            # 진행 바 업데이트
            self.progress_bar.setValue(progress-1)
            QApplication.processEvents()  # UI 업데이트

        # Pr_max와 Ec 값 표 형식으로 표시
        table_data = [[file_name, f'{pr_max:.2f}', f'{ec:.2f}'] for file_name, pr_max, ec in zip(file_names, pr_max_list, ec_list)]
        table = ax.table(cellText=table_data, colLabels=['File Name', 'Pr_max [μC/cm²]', 'Ec [MV/cm]'], loc='bottom', bbox=[0.1, -0.75, 0.8, 0.5])
        table.auto_set_font_size(True)  # 자동 폰트 크기 조정 비활성화
        #table.set_fontsize(12)  # 폰트 크기 조정
        table.auto_set_column_width(col=list(range(len(table_data[0]))))  # 열 너비 자동 조정

        # 텍스트 정렬
        for key, cell in table.get_celld().items():
            cell.set_text_props(ha='center', va='center', weight='bold')

        # 그래프와 표 사이의 간격 추가 조정
        fig.subplots_adjust(bottom=0.55)  # 하단 여백 추가 조정

        plot_file_path = os.path.join(new_folder_path, f'{timestamp}_PUND_W{self.device_width:.0f}xH{self.device_height:.0f}xt{self.device_thickness:.0f}_plot.png')
        fig.savefig(plot_file_path, dpi=300)

        # 진행 상황 업데이트
        self.progress_bar.setValue(100)
        self.progress_label.setText("Progress: 100%")
        QApplication.processEvents()  # UI 업데이트

        # 파일 저장 성공 메시지 표시
        QMessageBox.information(self, "Files Saved", f"New Excel file '{output_file_path}' and associated plot files have been created in '{output_dir}'.")

        # 진행 상황 창 닫기
        progress_widget.close()

        # 드래그앤드롭 창 유지
        self.show()

    def save_to_single_file_Tri(self, all_data, file_names):
        # 현재 날짜와 시간 가져오기
        now = datetime.now()
        timestamp = now.strftime('%y%m%d_%H%M%S')

        # 기본 파일 이름 설정
        default_file_name = f'{timestamp}_TriQV_W{self.device_width:.0f}xH{self.device_height:.0f}xt{self.device_thickness:.0f}'

        # 파일 저장 경로 선택 대화 상자 열기
        output_dir = QFileDialog.getExistingDirectory(self, "Select Folder to Save Files", os.path.expanduser("~"))

        # 사용자가 취소 버튼을 누른 경우
        if not output_dir:
            QMessageBox.information(self, "Saving Canceled", "Saving file has been canceled.")
            return

        # 새 폴더 생성
        new_folder_path = os.path.join(output_dir, default_file_name)
        try:
            os.makedirs(new_folder_path)
        except FileExistsError:
            pass

        # Excel 파일 경로 설정
        output_file_path = os.path.join(new_folder_path, f"{default_file_name}.xlsx")

        with pd.ExcelWriter(output_file_path, engine='xlsxwriter') as writer:
            # 그래프 생성을 위한 리스트 초기화
            e_field_list = []
            charge_list = []

            for i, df in enumerate(all_data):
                if df is not None:
                    # 'charge' 데이터 계산
                    charge_data = df['Δ Charge'].cumsum()
                    pol_max = max(charge_data)
                    pol_min = min(charge_data)
                    pol_average = (pol_max + pol_min) / 2
                    df['Charge'] = [(c - pol_average)/(self.device_area * 1e-14) for c in charge_data]
                    # 'E-field' 계산
                    efield = df['Voltage_V']
                    df['E-field'] = [(v*10/self.device_thickness) for v in efield]
                    df.to_excel(writer, sheet_name=file_names[i], index=False)
                    # 그래프 데이터 리스트에 추가
                    e_field_list.append(df['E-field'])
                    charge_list.append(df['Charge'])
                    
            # 새로운 sheet 생성
            new_df = pd.DataFrame()
            for j, sheet_df in enumerate(all_data):
                if sheet_df is not None:
                    new_df[f'E-field_{j+1}'] = sheet_df['E-field']
                    new_df[f'Charge_{j+1}'] = sheet_df['Charge']
            new_df.to_excel(writer, sheet_name='E-field_Charge', index=False)

            # 'Pr_Ec' 시트 생성
            pr_ec_df = pd.DataFrame()
            for k, sheet_df in enumerate(all_data):
                if sheet_df is not None:
                    #pr_max = max(sheet_df['charge'])
                    ec1_values = []
                    ec2_values = []
                    pr1_values = []
                    pr2_values = []
                    for j in range(len(sheet_df['Charge'])-1):
                        if sheet_df['Charge'][j] * sheet_df['Charge'][j+1] < 0:
                            if sheet_df['Charge'][j] > 0:
                                ec1 = sheet_df['E-field'][j+1]
                                ec1_values.append(ec1)
                                #print(ec1)
                            else:
                                ec2 = sheet_df['E-field'][j+1]
                                ec2_values.append(ec2)
                                #print(ec2)
                        else:
                            ec1_values.append(None)
                            ec2_values.append(None)
                            pr1_values.append(None)
                            pr2_values.append(None)

                        if sheet_df['E-field'][j] * sheet_df['E-field'][j+1] < 0:
                            if sheet_df['E-field'][j] > 0:
                                pr1 = sheet_df['Charge'][j+1]
                                pr1_values.append(pr1)
                                #print(pr1, "pr1")
                        else:
                            pr1_values.append(None)
                    
                    # 'None' 값을 제거하고 평균 계산
                    ec1_values = [value for value in ec1_values if value is not None]
                    ec2_values = [value for value in ec2_values if value is not None]
                    pr1_values = [value for value in pr1_values if value is not None]
         
                    if pr1_values:
                        pr_ec_df[f'Pr_max_{k+1}'] = [abs(np.mean(pr1_values))]                    
                    else:
                        pr_ec_df[f'Pr_max_{k+1}'] = [None]
                    if ec1_values and ec2_values:
                        pr_ec_df[f'Ec_{k+1}'] = [(abs(np.mean(ec1_values)) + abs(np.mean(ec2_values))) / 2]
                    else:
                        pr_ec_df[f'Ec_{k+1}'] = [None]

            # 'Pr_Ec' 시트에 데이터 저장
            pr_ec_df.to_excel(writer, sheet_name='Pr_Ec', index=False)

        # 진행 상황 표시를 위한 위젯 생성
        progress_widget = QWidget()
        progress_widget.setWindowTitle("Saving Progress")
        progress_widget.setGeometry(650, 400, 500, 250)

        # 배경색 및 스타일 설정
        progress_widget.setStyleSheet("background-color: #f0f0f0; border: 2px solid #0078d7; border-radius: 10px;")

        # 레이아웃 생성 (전체 레이아웃은 배경 없음)
        progress_layout = QVBoxLayout()
        progress_layout.setContentsMargins(10, 10, 10, 10)  # 여백 설정

        self.progress_bar = QProgressBar()
        self.progress_label = QLabel("Progressing...... 0%")

        # 기본 스타일을 사용하여 진행 바 설정
        self.progress_bar.setTextVisible(True)  # 텍스트 표시
        self.progress_bar.setMinimum(0)          # 최소값
        self.progress_bar.setMaximum(100)        # 최대값

        self.progress_label.setAlignment(Qt.AlignCenter)
        self.progress_bar.setAlignment(Qt.AlignCenter)  

        # 프로그레스 바를 감싸는 레이아웃 생성
        progress_bar_layout = QVBoxLayout()
        progress_bar_layout.addWidget(self.progress_bar)
        progress_layout.addLayout(progress_bar_layout)  # 프로그레스 바 레이아웃 추가

        # 레이아웃에 레이블 추가
        progress_layout.addWidget(self.progress_label)

        # 전체 레이아웃 설정
        progress_widget.setLayout(progress_layout)

        # 애니메이션 효과 추가
        animation = QPropertyAnimation(progress_widget, b"geometry")
        animation.setDuration(500)
        animation.setStartValue(QRect(750, 450, 300, 0))
        animation.setEndValue(QRect(750, 450, 300, 150))
        animation.start()

        progress_widget.show()

        # Matplotlib을 사용하여 그래프 생성 및 저장
        fig = plt.figure(figsize=(10, 16))
        ax = fig.add_subplot(111)

        # 그래프 색상 설정 (그라데이션 느낌)
        colors = plt.cm.Accent(np.linspace(0, 1, len(e_field_list)))

        for i, (e_field, charge) in enumerate(zip(e_field_list, charge_list)):
            ax.plot(e_field, charge, color=colors[i], label=file_names[i], linewidth=5)

            # 진행 상황 업데이트
            progress = int((i + 1) / len(charge_list) * 50)  # 전체 진행률의 33.3%를 차지
            self.progress_label.setText(f"Progressing...... {int(progress)}%")
            
            # 진행 바 업데이트
            self.progress_bar.setValue(progress)
            QApplication.processEvents()  # UI 업데이트

        # 축 라벨 폰트 크기 및 스타일 변경
        ax.set_xlabel('E-field [MV/cm]', fontsize=28, fontweight='bold')
        ax.set_ylabel('Charge [μC/cm²]', fontsize=28, fontweight='bold')

        # x축 값 폰트 스타일 변경
        ax.set_xticklabels(ax.get_xticks(), fontsize=14, fontweight='bold')
        ax.set_yticklabels(ax.get_xticks(), fontsize=14, fontweight='bold')

        # x축 눈금 값의 유효숫자 조정
        ax.xaxis.set_major_formatter(ticker.FormatStrFormatter('%.0f'))

        # y축 눈금 값의 유효숫자 조정
        ax.yaxis.set_major_formatter(ticker.FormatStrFormatter('%.0f'))

        # 눈금 라벨 폰트 크기 변경
        ax.tick_params(axis='both', which='major', labelsize=22, direction='in')

        # 제목 폰트 크기 및 스타일 변경
        ax.set_title(f'{default_file_name}', fontsize=16, fontweight='bold')

        # 축 라벨 위치 조정
        ax.xaxis.set_label_coords(0.5, -0.1)
        ax.yaxis.set_label_coords(-0.1, 0.5)

        ax.legend(loc='upper left', prop={'weight':'bold', 'size':12}, borderaxespad=0.1, frameon=False)

        # 그래프 틀의 선 두께 변경
        ax.spines['left'].set_linewidth(2)
        ax.spines['right'].set_linewidth(2)
        ax.spines['top'].set_linewidth(2)
        ax.spines['bottom'].set_linewidth(2)

        # 그리드 추가 (점선 및 minor grid)
        ax.grid(True, which='both', linestyle='--')
        ax.minorticks_on()
        ax.grid(True, which='minor', linestyle=':')

        # Pr_max와 Ec 값 표시 (bold체)
        pr_max_list = []
        ec_list = []
        for i, (e_field, charge) in enumerate(zip(e_field_list, charge_list)):
            ec1_values = []
            ec2_values = []
            pr1_values = []
            for j in range(len(charge)-1):
                if charge[j] * charge[j+1] < 0:
                    if charge[j] > 0:
                        ec1 = e_field[j+1]
                        ec1_values.append(ec1)
                    else:
                        ec2 = e_field[j+1]
                        ec2_values.append(ec2)
                if e_field[j] * e_field[j+1] < 0:
                    if e_field[j] > 0:
                        pr1 = charge[j+1]
                        pr1_values.append(pr1)

            # 'None' 값을 제거하고 평균 계산
            ec1_values = [value for value in ec1_values if value is not None]
            ec2_values = [value for value in ec2_values if value is not None]
            pr1_values = [value for value in pr1_values if value is not None]

            if pr1_values:
                pr_max = abs(np.mean(pr1_values))
            else:
                pr_max = None
            if ec1_values and ec2_values:
                ec = (abs(np.mean(ec1_values)) + abs(np.mean(ec2_values))) / 2
            else:
                ec = None

            pr_max_list.append(pr_max)
            ec_list.append(ec)

            # 진행 상황 업데이트
            progress = int(50 + (i + 1) / len(charge_list) * 50)  # 전체 진행률의 33.3%를 차지
            self.progress_label.setText(f"Progressing...... {int(progress)-1}%")
            
            # 진행 바 업데이트
            self.progress_bar.setValue(progress-1)
            QApplication.processEvents()  # UI 업데이트

        # Pr_max와 Ec 값 표 형식으로 표시
        table_data = [[file_name, f'{pr_max:.2f}', f'{ec:.2f}'] for file_name, pr_max, ec in zip(file_names, pr_max_list, ec_list)]
        table = ax.table(cellText=table_data, colLabels=['File Name', 'Pr_max [μC/cm²]', 'Ec [MV/cm]'], loc='bottom', bbox=[0.1, -0.75, 0.8, 0.5])
        table.auto_set_font_size(True)  # 자동 폰트 크기 조정 비활성화
        #table.set_fontsize(12)  # 폰트 크기 조정
        table.auto_set_column_width(col=list(range(len(table_data[0]))))  # 열 너비 자동 조정

        # 텍스트 정렬
        for key, cell in table.get_celld().items():
            cell.set_text_props(ha='center', va='center', weight='bold')

        # 그래프와 표 사이의 간격 추가 조정
        fig.subplots_adjust(bottom=0.55)  # 하단 여백 추가 조정

        plot_file_path = os.path.join(new_folder_path, f'{timestamp}_Tri_W{self.device_width:.0f}xH{self.device_height:.0f}xt{self.device_thickness:.0f}_plot.png')
        fig.savefig(plot_file_path, dpi=300)

        # 진행 상황 업데이트
        self.progress_bar.setValue(100)
        self.progress_label.setText("Progress: 100%")
        QApplication.processEvents()  # UI 업데이트

        # 파일 저장 성공 메시지 표시
        QMessageBox.information(self, "Files Saved", f"New Excel file '{output_file_path}' and associated plot files have been created in '{output_dir}'.")

        # 진행 상황 창 닫기
        progress_widget.close()

        # 드래그앤드롭 창 유지
        self.show()

    def save_to_single_file_IV(self, all_data, file_names):
        # 현재 날짜와 시간 가져오기
        now = datetime.now()
        timestamp = now.strftime('%y%m%d_%H%M%S')

        # 기본 파일 이름 설정
        default_file_name = f'{timestamp}_IV_W{self.device_width:.0f}xH{self.device_height:.0f}xt{self.device_thickness:.0f}'

        # 파일 저장 경로 선택 대화 상자 열기
        output_dir = QFileDialog.getExistingDirectory(self, "Select Folder to Save Files", os.path.expanduser("~"))

        # 사용자가 취소 버튼을 누른 경우
        if not output_dir:
            QMessageBox.information(self, "Saving Canceled", "Saving file has been canceled.")
            return

        # 새 폴더 생성
        new_folder_path = os.path.join(output_dir, default_file_name)
        try:
            os.makedirs(new_folder_path)
        except FileExistsError:
            pass

        # Excel 파일 경로 설정
        output_file_path = os.path.join(new_folder_path, f"{default_file_name}.xlsx")

        with pd.ExcelWriter(output_file_path, engine='xlsxwriter') as writer:
            # 그래프 생성을 위한 리스트 초기화
            time_list = []
            voltage_list = []
            current_list = []

            for i, df in enumerate(all_data):
                if df is not None:
                    df.to_excel(writer, sheet_name=file_names[i], index=False)
                    time_list.append(df['Time'])
                    voltage_list.append(df['Voltage'])
                    current_list.append(df['Current'])

            # 새로운 sheet 생성
            new_df = pd.DataFrame()
            for j, sheet_df in enumerate(all_data):
                if sheet_df is not None:
                    new_df[f'Time_{j+1}'] = sheet_df['Time']
                    new_df[f'Voltage_{j+1}'] = sheet_df['Voltage']
            new_df.to_excel(writer, sheet_name='Time-Voltage', index=False)

            # 새로운 sheet 생성
            new_df = pd.DataFrame()
            for j, sheet_df in enumerate(all_data):
                if sheet_df is not None:
                    new_df[f'Time_{j+1}'] = sheet_df['Time']
                    new_df[f'Current_{j+1}'] = sheet_df['Current']
            new_df.to_excel(writer, sheet_name='Time-Current', index=False)

            # 새로운 sheet 생성
            new_df = pd.DataFrame()
            for j, sheet_df in enumerate(all_data):
                if sheet_df is not None:
                    new_df[f'Time_{j+1}'] = sheet_df['Time']
                    new_df[f'Voltage_{j+1}'] = sheet_df['Voltage']
                    new_df[f'Current_{j+1}'] = sheet_df['Current']
            new_df.to_excel(writer, sheet_name='Time-Voltage-Current', index=False)

        # 진행 상황 표시를 위한 위젯 생성
        progress_widget = QWidget()
        progress_widget.setWindowTitle("Saving Progress")
        progress_widget.setGeometry(650, 400, 500, 250)

        # 배경색 및 스타일 설정
        progress_widget.setStyleSheet("background-color: #f0f0f0; border: 2px solid #0078d7; border-radius: 10px;")

        # 레이아웃 생성 (전체 레이아웃은 배경 없음)
        progress_layout = QVBoxLayout()
        progress_layout.setContentsMargins(10, 10, 10, 10)  # 여백 설정

        self.progress_bar = QProgressBar()
        self.progress_label = QLabel("Progressing...... 0%")

        # 기본 스타일을 사용하여 진행 바 설정
        self.progress_bar.setTextVisible(True)  # 텍스트 표시
        self.progress_bar.setMinimum(0)          # 최소값
        self.progress_bar.setMaximum(100)        # 최대값

        self.progress_label.setAlignment(Qt.AlignCenter)
        self.progress_bar.setAlignment(Qt.AlignCenter)  

        # 프로그레스 바를 감싸는 레이아웃 생성
        progress_bar_layout = QVBoxLayout()
        progress_bar_layout.addWidget(self.progress_bar)
        progress_layout.addLayout(progress_bar_layout)  # 프로그레스 바 레이아웃 추가

        # 레이아웃에 레이블 추가
        progress_layout.addWidget(self.progress_label)

        # 전체 레이아웃 설정
        progress_widget.setLayout(progress_layout)

        # 애니메이션 효과 추가
        animation = QPropertyAnimation(progress_widget, b"geometry")
        animation.setDuration(500)
        animation.setStartValue(QRect(750, 450, 300, 0))
        animation.setEndValue(QRect(750, 450, 300, 150))
        animation.start()

        progress_widget.show()

        # Matplotlib을 사용하여 그래프 생성 및 저장
        fig = plt.figure(figsize=(24, 16))
        ax1 = fig.add_subplot(111)

        # 그래프 색상 설정 (그라데이션 느낌)
        colors = plt.cm.Accent(np.linspace(0, 1, len(time_list)))

        # 왼쪽 y축에 Voltage 데이터 추가
        for i, (time, voltage) in enumerate(zip(time_list, voltage_list)):
            ax1.plot(time*1e3, voltage, color='gray', linestyle='--', label=file_names[i], linewidth=3)

            # 진행 상황 업데이트
            progress = int((i + 1) / len(time_list) * 25)  # 전체 진행률의 33.3%를 차지
            self.progress_label.setText(f"Progressing...... {int(progress)}%")
            
            # 진행 바 업데이트
            self.progress_bar.setValue(progress)
            QApplication.processEvents()  # UI 업데이트

        # 축 라벨 폰트 크기 및 스타일 변경
        ax1.set_xlabel('Time [msec]', fontsize=52, fontweight='bold')
        ax1.set_ylabel('Polarization [μC/cm²]', fontsize=52, fontweight='bold')

        # x축 값 폰트 스타일 변경
        ax1.set_xticklabels(ax1.get_xticks(), fontsize=14, fontweight='bold')
        ax1.set_yticklabels(ax1.get_yticks(), fontsize=14, fontweight='bold')

        # x축 눈금 값의 유효숫자 조정
        ax1.xaxis.set_major_formatter(ticker.FormatStrFormatter('%.1f'))

        # y축 눈금 값의 유효숫자 조정
        ax1.yaxis.set_major_formatter(ticker.FormatStrFormatter('%.0f'))

        # 눈금 라벨 폰트 크기 변경
        ax1.tick_params(axis='both', which='major', labelsize=40, direction='in')

        # 제목 폰트 크기 및 스타일 변경
        ax1.set_title(f'{default_file_name}', fontsize=20, fontweight='bold')

        # 축 라벨 위치 조정
        ax1.xaxis.set_label_coords(0.5, -0.075)
        ax1.yaxis.set_label_coords(-0.06, 0.5)

        # 오른쪽 y축 추가
        ax2 = ax1.twinx()  # 오른쪽 y축 생성

        # 오른쪽 y축에 Current 데이터 추가
        for i, (time, current) in enumerate(zip(time_list, current_list)):
            ax2.plot(time*1e3, current*1e6, color=colors[i], label=f'{file_names[i]}', linewidth=5)

            # 진행 상황 업데이트
            progress = int(25 + (i + 1) / len(current_list) * 25)  # 전체 진행률의 33.3%를 차지
            self.progress_label.setText(f"Progressing...... {int(progress)}%")
            
            # 진행 바 업데이트
            self.progress_bar.setValue(progress)
            QApplication.processEvents()  # UI 업데이트

        # 오른쪽 y축 설정
        ax2.set_ylabel('Current [μA]', fontsize=52, fontweight='bold')
        ax2.set_yticklabels(ax2.get_yticks(), fontsize=14, fontweight='bold')

        # y축 눈금 값의 유효숫자 조정
        ax2.yaxis.set_major_formatter(ticker.FormatStrFormatter('%.0f'))

        ax2.yaxis.set_label_coords(1.075, 0.5)

        ax2.tick_params(axis='both', which='major', labelsize=40, direction='in')

        # 그래프 틀의 선 두께 변경
        ax1.spines['left'].set_linewidth(2)
        ax1.spines['right'].set_linewidth(2)
        ax1.spines['top'].set_linewidth(2)
        ax1.spines['bottom'].set_linewidth(2)

        # 오른쪽 y축에 대한 범례 추가
        ax2.legend(loc='upper right', prop={'weight':'bold', 'size':21}, borderaxespad=0.1, frameon=False)

        # 그리드 추가 (점선 및 minor grid)
        ax1.grid(True, which='both', linestyle='--')
        ax1.minorticks_on()
        ax1.grid(True, which='minor', linestyle=':')

        # 그래프 저장
        plot_file_path = os.path.join(new_folder_path, f'{timestamp}_t-IV_W{self.device_width:.0f}xH{self.device_height:.0f}xt{self.device_thickness:.0f}_plot.png')
        fig.savefig(plot_file_path, dpi=300)

        # Matplotlib을 사용하여 그래프 생성 및 저장
        fig2 = plt.figure(figsize=(24, 16))
        ax21 = fig2.add_subplot(111)

        # 그래프 색상 설정 (그라데이션 느낌)
        colors2 = plt.cm.Accent(np.linspace(0, 1, len(voltage_list)))

        # 왼쪽 y축에 Voltage 데이터 추가
        for i, (voltage, current) in enumerate(zip(voltage_list, current_list)):
            ax21.plot(voltage, current*1e6, color=colors2[i], label=file_names[i], linewidth=5)

            # 진행 상황 업데이트
            progress = int(50 + (i + 1) / len(current_list) * 25)  # 전체 진행률의 33.3%를 차지
            self.progress_label.setText(f"Progressing...... {int(progress)}%")
            
            # 진행 바 업데이트
            self.progress_bar.setValue(progress)
            QApplication.processEvents()  # UI 업데이트

        # 축 라벨 폰트 크기 및 스타일 변경
        ax21.set_xlabel('Voltage [V]', fontsize=52, fontweight='bold')
        ax21.set_ylabel('Current [μA]', fontsize=52, fontweight='bold')

        # x축 값 폰트 스타일 변경
        ax21.set_xticklabels(ax21.get_xticks(), fontsize=14, fontweight='bold')
        ax21.set_yticklabels(ax21.get_yticks(), fontsize=14, fontweight='bold')

        # x축 눈금 값의 유효숫자 조정
        ax21.xaxis.set_major_formatter(ticker.FormatStrFormatter('%.0f'))

        # y축 눈금 값의 유효숫자 조정
        ax21.yaxis.set_major_formatter(ticker.FormatStrFormatter('%.0f'))

        # 눈금 라벨 폰트 크기 변경
        ax21.tick_params(axis='both', which='major', labelsize=40, direction='in')

        # 제목 폰트 크기 및 스타일 변경
        ax21.set_title(f'{default_file_name}', fontsize=20, fontweight='bold')

        # 축 라벨 위치 조정
        ax21.xaxis.set_label_coords(0.5, -0.075)
        ax21.yaxis.set_label_coords(-0.075, 0.5)

        # 그래프 틀의 선 두께 변경
        ax21.spines['left'].set_linewidth(2)
        ax21.spines['right'].set_linewidth(2)
        ax21.spines['top'].set_linewidth(2)
        ax21.spines['bottom'].set_linewidth(2)

        # 오른쪽 y축에 대한 범례 추가
        ax21.legend(loc='upper left', prop={'weight':'bold', 'size':21}, borderaxespad=0.1, frameon=False)

        # 그리드 추가 (점선 및 minor grid)
        ax21.grid(True, which='both', linestyle='--')
        ax21.minorticks_on()
        ax21.grid(True, which='minor', linestyle=':')

        # 그래프 저장
        plot_file_path2 = os.path.join(new_folder_path, f'{timestamp}_VI_W{self.device_width:.0f}xH{self.device_height:.0f}xt{self.device_thickness:.0f}_plot.png')
        fig2.savefig(plot_file_path2, dpi=300)

        # 진행 상황 업데이트
        self.progress_bar.setValue(100)
        self.progress_label.setText("Progress: 100%")
        QApplication.processEvents()  # UI 업데이트

        # 파일 저장 성공 메시지 표시
        QMessageBox.information(self, "File Saved", f"New Excel file '{output_file_path}' has been created.")

        # 진행 상황 창 닫기
        progress_widget.close()

        # 드래그앤드롭 창 유지
        self.show()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    csv_file_processor = CSVFileProcessor()
    csv_file_processor.show()
    sys.exit(app.exec_())