import sys
import csv
import re
import os
import json
import requests
import resources_rc
from datetime import datetime
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QFileDialog, QDialog,
    QVBoxLayout, QWidget, QLabel, QMessageBox,
    QProgressBar, QHBoxLayout, QComboBox, QLineEdit,
    QDialogButtonBox, QFormLayout, QCompleter, QInputDialog
)
from PyQt5.QtGui import QFont, QIcon
from PyQt5.QtCore import Qt, QPropertyAnimation, QRect, pyqtSignal, QThread, QStandardPaths
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
from matplotlib.ticker import FuncFormatter

CURRENT_VERSION = '1.2.7'

# 릴리스 노트 HTML을 이 아래에 정의해 두면 나중에 수정하기 편합니다.
RELEASE_NOTES_HTML = (
    "<html><body>"
    f"<b>새 버전 V{CURRENT_VERSION} 릴리스 노트</b><br><br>"
    "- 정제혁이 자꾸 찡찡거려서 마이너 업데이트 해줌"
    "- 25년 12월 31일 섭종 예정임ㅅㄱ"
    "</body></html>"
)

# --- TEST FLAG: True 로 두면 항상 인터넷 연결 없음처럼 동작합니다 ---
SIMULATE_OFFLINE = False #배포 시 False로 변경 필수

MAX_HISTORY = 20

# ─────────────────────────────────────────────────────────────────────────────
# 1) 앱 전용 데이터 폴더 경로
app_data_dir = QStandardPaths.writableLocation(QStandardPaths.AppDataLocation)
os.makedirs(app_data_dir, exist_ok=True)

# 2) JSON 파일은 여기에 저장
HISTORY_FILE = os.path.join(app_data_dir, 'history.json')
CONFIG_FILE  = os.path.join(app_data_dir, 'config.json')
# ─────────────────────────────────────────────────────────────────────────────

# 전역 예외 처리기: 처리되지 않은 예외 발생 시 history.json 을 삭제
def handle_uncaught_exception(exc_type, exc_value, exc_traceback):
    try:
        if os.path.exists(HISTORY_FILE):
            os.remove(HISTORY_FILE)
    except Exception:
        pass
    sys.__excepthook__(exc_type, exc_value, exc_traceback)

sys.excepthook = handle_uncaught_exception

class UpdateThread(QThread):
    progress = pyqtSignal(int)
    finished = pyqtSignal(str)

    def __init__(self, version_checker):
        super().__init__()
        self.version_checker = version_checker

    def run(self):
        latest_version = self.version_checker.get_latest_version()
        if latest_version is None:
            self.progress.emit(-1)
            return
        download_url = f"https://github.com/khoon0/B1500_206/raw/master/B1500_calculator_V{latest_version}.exe"
        self.download_file(download_url, latest_version)

    def download_file(self, download_url, latest_version):
        try:
            response = requests.get(download_url, stream=True)
            total_size = int(response.headers.get('content-length', 0))
            downloaded = 0
            output_file = f"B1500_calculator_V{latest_version}.exe"
            with open(output_file, 'wb') as f:
                for chunk in response.iter_content(chunk_size=1024):
                    f.write(chunk)
                    downloaded += len(chunk)
                    self.progress.emit(int(downloaded / total_size * 100))
            self.finished.emit(latest_version)
        except Exception:
            self.progress.emit(-1)

class CSVFileProcessor(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowIcon(QIcon(':/KKH.ico'))
        # dimension history
        self.history = self.load_history()
        self.swap_current_voltage = False
        self.last_load_dir = os.getcwd()
        self.last_save_dir = os.getcwd()

        # ─── 패치 노트 한 번만 표시 ───
        #self.maybe_show_patch_notes()
        # ────────────────────────────

        self.initUI()
        # (moved) 업데이트 체크: 이제 __init__에서는 호출하지 않습니다.

    # ─────────────────────────────────────────────────────────────────────────
    # 설정 파일 읽기/쓰기 (패치 노트 한 번만 표시 관리)
    def load_config(self):
        try:
            with open(CONFIG_FILE, 'r') as f:
                return json.load(f)
        except Exception:
            return {}

    def save_config(self, cfg):
        with open(CONFIG_FILE, 'w') as f:
            json.dump(cfg, f, indent=2)

    def maybe_show_patch_notes(self):
        cfg = self.load_config()
        last_seen = cfg.get('last_seen_version')
        if last_seen != CURRENT_VERSION:
            # 상단에 정의한 RELEASE_NOTES_HTML 을 사용
            msg = QMessageBox(self)
            msg.setWindowTitle(f"What's New in V{CURRENT_VERSION}")
            msg.setTextFormat(Qt.RichText)
            msg.setText(RELEASE_NOTES_HTML)
            msg.setStandardButtons(QMessageBox.Ok)
            msg.exec_()

            cfg['last_seen_version'] = CURRENT_VERSION
            self.save_config(cfg)

    def load_history(self):
        try:
            with open(HISTORY_FILE, 'r') as f:
                return json.load(f)
        except Exception:
            return []

    def save_history(self, dims_str):
        if dims_str in self.history:
            self.history.remove(dims_str)
        self.history.insert(0, dims_str)
        self.history = self.history[:MAX_HISTORY]
        with open(HISTORY_FILE, 'w') as f:
            json.dump(self.history, f, indent=2)

    def closeEvent(self, event):
        try:
            if os.path.exists(HISTORY_FILE):
                os.remove(HISTORY_FILE)
        except Exception:
            pass
        event.accept()

    def initUI(self):
        self.setWindowTitle('B1500 Auto Calculator for Ferroelectric (Keysight B1500)')
        self.setGeometry(400, 150, 1024, 768)
        self.setWindowIcon(QIcon(':/KKH.ico'))
        main_layout = QVBoxLayout()
        label = QLabel(self.create_label_text())
        label.setTextFormat(Qt.RichText)
        label.setTextInteractionFlags(Qt.TextBrowserInteraction)
        label.setOpenExternalLinks(True)
        label.mousePressEvent = self.select_files
        label.setFont(self.create_font())
        label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(label)

        version_layout = QHBoxLayout()
        version_label = QLabel(f"Version: {CURRENT_VERSION}")
        version_label.setAlignment(Qt.AlignRight)
        version_label.setFont(QFont("Arial", 10))
        version_layout.addStretch()
        version_layout.addWidget(version_label)
        version_layout.setContentsMargins(0, 0, 0, 0)
        version_label.setFixedHeight(15)
        main_layout.addLayout(version_layout)

        self.central_widget = QWidget()
        self.central_widget.setLayout(main_layout)
        self.setCentralWidget(self.central_widget)
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

    def select_files(self, event):
        paths, _ = QFileDialog.getOpenFileNames(
            self, "Select CSV Files", self.last_load_dir,
            "CSV Files (*.csv)"
        )
        if paths:
            self.last_load_dir = os.path.dirname(paths[0])
            names = [os.path.basename(p) for p in paths]
            self.process_selected_files(paths, names)

    # ── 업데이트 체크, 다운로드, 삭제 로직 ──
    def check_for_updates(self):
        latest_version = self.get_latest_version()
        if latest_version is None:
            self.show_message("인터넷 연결 오류", "인터넷 연결이 없습니다. 기존 버전으로 실행합니다.")
            return
        if self.is_newer_version(latest_version, CURRENT_VERSION):
            self.prompt_update(latest_version)
        else:
            self.prompt_delete_old_versions()

    def get_latest_version(self):
        if SIMULATE_OFFLINE:
            # 테스트 모드: 항상 오프라인
            return None
        try:
            response = requests.get("https://raw.githubusercontent.com/khoon0/B1500_206/master/latest_version.txt")
            return response.text.strip()
        except requests.ConnectionError:
            return None

    def prompt_update(self, latest_version):
        message = f"새로운 업데이트가 존재합니다.\nV{CURRENT_VERSION} -> V{latest_version}"
        QMessageBox.information(self, 'Update Available', message, QMessageBox.Ok)
        self.start_update_process()

    def prompt_delete_old_versions(self):
        reply = QMessageBox.question(self, 'Newest Version',
                                     "현재 최신 버전입니다. 기존 버전은 삭제할까요?",
                                     QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes:
            self.delete_old_versions()

    def delete_old_versions(self):
        current_executable = os.path.abspath(sys.executable)
        current_version_number = CURRENT_VERSION
        current_version_tuple = tuple(map(int, current_version_number.split(".")))
        directory = os.path.dirname(current_executable)
        deleted_files = []
        for filename in os.listdir(directory):
            if filename.endswith(".exe"):
                match = re.search(r'V(\d+\.\d+\.\d+)', filename)
                if match:
                    version_str = match.group(1)
                    version_tuple = tuple(map(int, version_str.split(".")))
                    if version_tuple < current_version_tuple:
                        file_path = os.path.join(directory, filename)
                        try:
                            os.remove(file_path)
                            deleted_files.append(file_path)
                        except Exception as e:
                            self.show_message("파일 삭제 오류", f"파일 삭제 중 오류 발생: {e}")
        if deleted_files:
            deleted_files_message = "\n".join(deleted_files)
            self.show_message("삭제된 파일", f"다음 파일이 삭제되었습니다:\n{deleted_files_message}")
        else:
            self.show_message("삭제된 파일", "삭제할 파일이 없습니다.")

    def is_newer_version(self, latest, current):
        return tuple(map(int, latest.split("."))) > tuple(map(int, current.split(".")))

    def start_update_process(self):
        self.download_update()

    def download_update(self):
        self.setEnabled(False)
        self.update_dialog = QDialog(self)
        self.update_dialog.setWindowTitle("Updating Process")
        update_layout = QVBoxLayout()
        message_label = QLabel("업데이트 다운로드 중입니다. 잠시만 기다려주세요...")
        update_layout.addWidget(message_label)
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        update_layout.addWidget(self.progress_bar)
        self.update_dialog.setLayout(update_layout)
        self.update_dialog.setFixedSize(400, 150)
        self.update_dialog.show()
        self.update_thread = UpdateThread(self)
        self.update_thread.progress.connect(self.update_progress)
        self.update_thread.finished.connect(self.update_finished)
        self.update_thread.start()

    def update_progress(self, value):
        if value == -1:
            self.show_message(f"Error downloading update", "다운로드 중 오류가 발생했습니다.")
            self.update_dialog.close()
        else:
            self.progress_bar.setValue(value)

    def update_finished(self, latest_version):
        self.update_dialog.close()
        self.show_message("Download complete", f"업데이트가 완료되었습니다. 새로운 버전 V{latest_version} 으로 다시 실행해주세요.")
        QApplication.quit()

    def show_message(self, title, message):
        QMessageBox.information(self, title, message)

    def dragEnterEvent(self, event):
        event.accept() if event.mimeData().hasUrls() else event.ignore()

    def get_user_inputs(self):
        while True:
            dialog = QDialog(self)
            dialog.setWindowTitle("Input Measurement Method")
            layout = QFormLayout(dialog)
            self.data_order_combo = QComboBox()
            self.data_order_combo.addItems([
                "'Time', 'Voltage', 'Current'",
                "'Time', 'Current', 'Voltage'"
            ])
            layout.addRow("Choose Data Order:", self.data_order_combo)
            self.measurement_method_combo = QComboBox()
            self.measurement_method_combo.addItems([
                "PUND measurement. P-E Curve + I-V Curve",
                "Triangular pulse measurement. Q-E Curve + I-V Curve",
                "(보류) I-V Curve"
            ])
            layout.addRow("Select Measurement Method:", self.measurement_method_combo)
            self.device_dimensions_input = QLineEdit()
            completer = QCompleter(self.history)
            completer.setCaseSensitivity(Qt.CaseInsensitive)
            self.device_dimensions_input.setCompleter(completer)
            if self.history:
                self.device_dimensions_input.setText(self.history[0])
            layout.addRow("Device Dimensions (w,h,t):", self.device_dimensions_input)
            self.device_dimensions_input.setFocus()
            buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel, dialog)
            layout.addWidget(buttons)
            buttons.accepted.connect(dialog.accept)
            buttons.rejected.connect(dialog.reject)
            if dialog.exec_() == QDialog.Accepted:
                order = self.data_order_combo.currentText()
                method = self.measurement_method_combo.currentText()
                dims_str = self.device_dimensions_input.text().strip()
                self.swap_current_voltage = (order == "'Time', 'Current', 'Voltage'")
                try:
                    w, h, t = map(float, dims_str.split(','))
                    self.save_history(dims_str)
                    return method, [w, h, t]
                except ValueError:
                    QMessageBox.information(
                        self, "Error input",
                        "Please enter device dimensions as: width,height,thickness"
                    )
            else:
                return None, None

    def process_selected_files(self, file_paths, file_names):
        method, dims = self.get_user_inputs()
        if method is None:
            return
        self.process_files(file_paths, file_names, method, dims)

    def process_files(self, file_paths, file_names, measurement_method=None, dimensions=None):
        all_data = []
        if measurement_method is None or dimensions is None:
            measurement_method, dimensions = self.get_user_inputs()
            if measurement_method is None or dimensions is None:
                return
        width, height, thickness = map(float, dimensions)
        self.device_width = width
        self.device_height = height
        self.device_thickness = thickness
        self.device_area = width * height

        for file_path, file_name in zip(file_paths, file_names):
            data = None
            if measurement_method == "PUND measurement. P-E Curve + I-V Curve":
                data = self.process_csv_file_PUND(file_path, file_name)
            elif measurement_method == "Triangular pulse measurement. Q-E Curve + I-V Curve":
                data = self.process_csv_file_Tri(file_path, file_name)
            elif measurement_method == "(보류) I-V Curve":
                data = self.process_csv_file_IV(file_path, file_name)
            else:
                QMessageBox.warning(self, "Error", "Invalid measurement method selected.")
                return
            if data is not None:
                all_data.append(data)

        if all_data:
            self.save_data(measurement_method, all_data, file_names)

    def save_data(self, measurement_method, all_data, file_names):
        if measurement_method == "PUND measurement. P-E Curve + I-V Curve":
            self.save_to_single_file_PUND(all_data, file_names)
        elif measurement_method == "Triangular pulse measurement. Q-E Curve + I-V Curve":
            self.save_to_single_file_Tri(all_data, file_names)
        elif measurement_method == "(보류) I-V Curve":
            self.save_to_single_file_IV(all_data, file_names)

    def process_selected_files(self, file_paths, file_names):
        # 사용자 입력을 한 번만 요청하고 그 결과를 process_files에 전달
        measurement_method, dimensions = self.get_user_inputs()
        if measurement_method is None or dimensions is None:
            return  # 사용자가 입력을 취소한 경우

        self.process_files(file_paths, file_names, measurement_method, dimensions)

    def dropEvent(self, event):
        urls = event.mimeData().urls()
        file_paths = [str(url.toLocalFile()) for url in urls]
        file_names = [os.path.basename(path) for path in file_paths]

        # 사용자 입력을 한 번만 요청하고 그 결과를 process_files에 전달
        measurement_method, dimensions = self.get_user_inputs()
        if measurement_method is None or dimensions is None:
            return  # 사용자가 입력을 취소한 경우, 이후 처리 중단

        self.process_files(file_paths, file_names, measurement_method, dimensions)

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
                # 시간축
                B_data = [float(row[1]) for row in data[start_row:end_row]]

                # 원본 voltage/current 컬럼 그대로 읽어두고
                raw_C = [float(row[2]) for row in data[start_row:end_row]]
                raw_D = [float(row[3]) for row in data[start_row:end_row]]

                # 스와핑 먼저: swap_current_voltage=True 면 C↔D 교체
                if self.swap_current_voltage:
                    C_data, D_data = raw_D[:], raw_C[:]
                else:
                    C_data, D_data = raw_C[:], raw_D[:]

                # 그제야 current 데이터에만 부호 반전 적용
                D_data = [-d for d in D_data]

                # 데이터 가공
                time_interval = B_data[1] - B_data[0]

                if measurement_type == "PUND":
                    num_chunks = 4
                    chunk_size = len(C_data) // num_chunks
                    C_chunks = [C_data[i * chunk_size:(i + 1) * chunk_size] for i in range(num_chunks)]
                    D_chunks = [D_data[i * chunk_size:(i + 1) * chunk_size] for i in range(num_chunks)]

                    D_diff = [d1 - d2 for d1, d2 in zip(D_chunks[0], D_chunks[1])] + [d3 - d4 for d3, d4 in zip(D_chunks[2], D_chunks[3])]

                    # 데이터 프레임 생성
                    voltage_diff = C_chunks[0] + C_chunks[2]
                    current_diff = D_diff

                    # NaN으로 채우기
                    max_length = len(B_data)
                    voltage_diff += [np.nan] * (max_length - len(voltage_diff))
                    current_diff += [np.nan] * (max_length - len(current_diff))

                    # 데이터프레임 생성
                    df = pd.DataFrame({
                        'Time': B_data,
                        'Voltage': C_data,
                        'Current': D_data,
                        'Time Interval': [time_interval] * max_length,
                        'Voltage_diff': voltage_diff,
                        'Current_diff': current_diff,
                        ' ': [''] * max_length,
                        'Voltage_V': voltage_diff,
                        'Δ Charge': [(d * time_interval) for d in current_diff],
                        '  ': [''] * max_length,
                        'E-field': voltage_diff,
                        'Polarization': [0] * max_length
                    })

                elif measurement_type == "Tri":
                    # 데이터 프레임 생성
                    df = pd.DataFrame({
                        'Time':B_data,
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
    
    def msec_formatter(self, x, pos):
        s = f"{x:.2f}"
        if s.endswith("00"):
            return str(int(round(x)))
        if s.endswith("0"):
            return s[:-1]
        return s

    def save_to_single_file_PUND(self, all_data, file_names):
        # 현재 날짜와 시간 가져오기
        now = datetime.now()
        timestamp = now.strftime('%y%m%d_%H%M%S')

        # 기본 파일 이름 설정
        default_file_name_PUND = f'{timestamp}_PUND_W{self.device_width:.0f}xH{self.device_height:.0f}xt{self.device_thickness:.0f}'
        default_file_name_IV = f'{timestamp}_IV_W{self.device_width:.0f}xH{self.device_height:.0f}xt{self.device_thickness:.0f}'

        # 파일 저장 경로 선택 대화 상자 열기
        # —————————————————————————————————————————————————————————————
        # 1) "폴더+이름"을 한 번에 선택/입력할 다이얼로그 생성
        dlg = QFileDialog(self, "Select folder and enter name", self.last_save_dir)
        dlg.setAcceptMode(QFileDialog.AcceptSave)           # 저장 모드
        dlg.setFileMode(QFileDialog.Directory)              # 디렉토리 선택
        dlg.setOption(QFileDialog.ShowDirsOnly, True)       # 파일은 숨기고 폴더만
        dlg.selectFile(default_file_name_PUND)               # 기본으로 채워둘 새 폴더명

        if dlg.exec_() != QDialog.Accepted:
            QMessageBox.information(self, "Saving Canceled", "Saving file has been canceled.")
            return

        # 2) 선택된 전체 경로 → 부모 디렉토리와 폴더명 분리
        selected = dlg.selectedFiles()[0]                   # e.g. "C:/Users/.../20250521_PUND_..."
        parent_dir = os.path.dirname(selected)
        folder_name = os.path.basename(selected)

        # 3) 결과를 저장 디렉토리로 사용
        self.last_save_dir = parent_dir
        new_folder_path = selected
        os.makedirs(new_folder_path, exist_ok=True)
        # —————————————————————————————————————————————————————————————

        # Excel 파일 경로 설정
        output_file_path_PUND = os.path.join(new_folder_path, f"{default_file_name_PUND}.xlsx")
        output_file_path_IV = os.path.join(new_folder_path, f"{default_file_name_IV}.xlsx")

        with pd.ExcelWriter(output_file_path_PUND, engine='xlsxwriter') as writer:
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
                    df['Polarization'] = [(c - pol_average) / (self.device_area * 1e-14) for c in charge_data]

                    # 'E-field' 계산
                    efield = df['Voltage_V']
                    df['E-field'] = [(v * 10 / self.device_thickness) for v in efield]
                    
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
            ps_list = []
            pr_max_list = []
            ec_list = []

            for k, sheet_df in enumerate(all_data):
                if sheet_df is not None:
                    ec1_values, ec2_values, pr1_values = [], [], []

                    for j in range(len(sheet_df['Polarization']) - 1):
                        # E-field의 연속한 두 데이터의 곱이 음수인 경우
                        if sheet_df['E-field'][j] * sheet_df['E-field'][j + 1] < 0:
                            e1, e2 = sheet_df['E-field'][j], sheet_df['E-field'][j + 1]
                            p1, p2 = sheet_df['Polarization'][j], sheet_df['Polarization'][j + 1]
                            if e1 != e2:
                                m = (p2 - p1) / (e2 - e1)
                                b = abs(p1 - m * e1)
                                pr1_values.append(b)

                        # Polarization의 연속한 두 데이터의 곱이 음수인 경우
                        if sheet_df['Polarization'][j] * sheet_df['Polarization'][j + 1] < 0:
                            e3, e4 = sheet_df['E-field'][j], sheet_df['E-field'][j + 1]
                            p3, p4 = sheet_df['Polarization'][j], sheet_df['Polarization'][j + 1]
                            if p3 != p4:
                                m = (e4 - e3) / (p4 - p3)
                                b = e3 - m * p3
                                if p3 > 0:
                                    ec1_values.append(b)
                                else:
                                    ec2_values.append(b)

                    # ───── 여기서부터 수정 ─────
                    # first_valid_index / last_valid_index 사용하여
                    # 실제 마지막 유효 데이터로 절편 계산
                    idx_e_first = sheet_df['E-field'].first_valid_index()
                    idx_e_last  = sheet_df['E-field'].last_valid_index()
                    idx_p_first = sheet_df['Polarization'].first_valid_index()
                    idx_p_last  = sheet_df['Polarization'].last_valid_index()

                    if idx_e_first is not None and idx_e_last is not None \
                       and idx_p_first is not None and idx_p_last is not None:
                        e_start = sheet_df.at[idx_e_first, 'E-field']
                        e_end   = sheet_df.at[idx_e_last,   'E-field']
                        p_start = sheet_df.at[idx_p_first, 'Polarization']
                        p_end   = sheet_df.at[idx_p_last,   'Polarization']
                        if e_start != e_end:
                            m_end = (p_end - p_start) / (e_end - e_start)
                            b_end = abs(p_start - m_end * e_start)
                            pr1_values.append(b_end)
                    # ───────────────────────────

                    # 'None' 값을 제거하고 평균 계산
                    pr1_values = [value for value in pr1_values if value is not None]
                    ec1_values = [value for value in ec1_values if value is not None]
                    ec2_values = [value for value in ec2_values if value is not None]

                    # Ps, Pr_max, Ec 계산
                    pol_max = abs(max(sheet_df['Polarization']))
                    pol_min = abs(min(sheet_df['Polarization']))
                    ps = (pol_max + pol_min) / 2
                    pr_max = abs(np.mean(pr1_values)) if pr1_values else None
                    ec = (abs(np.mean(ec1_values)) + abs(np.mean(ec2_values))) / 2 if ec1_values and ec2_values else None

                    ps_list.append(ps)
                    pr_max_list.append(pr_max)
                    ec_list.append(ec)

                    # DataFrame에 추가
                    pr_ec_df[f'Ps_{k + 1}']     = [ps]
                    pr_ec_df[f'Pr_max_{k + 1}'] = [pr_max]
                    pr_ec_df[f'Ec_{k + 1}']     = [ec]

            # 'Pr_Ec' 시트에 데이터 저장
            pr_ec_df.to_excel(writer, sheet_name='Ps_Pr_Ec', index=False)

        with pd.ExcelWriter(output_file_path_IV, engine='xlsxwriter') as writer:
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

        # 레이아웃 생성
        progress_layout = QVBoxLayout()
        progress_layout.setContentsMargins(10, 10, 10, 10)

        self.progress_bar = QProgressBar()
        self.progress_label = QLabel("Progressing...... 0%")

        # 진행 바 설정
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setRange(0, 100)

        self.progress_label.setAlignment(Qt.AlignCenter)
        self.progress_bar.setAlignment(Qt.AlignCenter)

        # 프로그레스 바를 감싸는 레이아웃 생성
        progress_layout.addWidget(self.progress_bar)
        progress_layout.addWidget(self.progress_label)

        # 전체 레이아웃 설정
        progress_widget.setLayout(progress_layout)

        # 애니메이션 효과 추가
        animation = QPropertyAnimation(progress_widget, b"geometry")
        animation.setDuration(500)
        animation.setStartValue(QRect(750, 450, 300, 0))
        animation.setEndValue(QRect(750, 450, 300, 150))
        animation.start()

        # 진행 상황 창을 즉시 표시
        progress_widget.show()
        QApplication.processEvents()  # UI 업데이트

        # 파일 이름과 관련된 데이터 리스트를 함께 묶기
        combined_data_PUND = list(zip(file_names, e_field_list, polarization_list, ps_list, pr_max_list, ec_list))
        combined_data_IV = list(zip(file_names, time_list, voltage_list, current_list))

        # 파일 이름을 기준으로 정렬
        combined_data_PUND.sort(key=lambda x: x[0])  # x[0]은 file_name
        combined_data_IV.sort(key=lambda x: x[0])  # x[0]은 file_name

        # 정렬된 데이터 분리
        sorted_file_names_PUND, sorted_e_field_list_PUND, sorted_polarization_list_PUND, sorted_ps_list_PUND, sorted_pr_max_list_PUND, sorted_ec_list_PUND = zip(*combined_data_PUND)
        sorted_file_names_IV, sorted_time_list_IV, sorted_voltage_list_IV, sorted_current_list_IV = zip(*combined_data_IV)

        # Matplotlib을 사용하여 그래프 생성 및 저장
        fig, ax = plt.subplots(figsize=(10, 16))

        # 그래프 색상 설정
        colors = plt.cm.viridis(np.linspace(0, 1, len(sorted_e_field_list_PUND)))

        # 그래프 그리기 및 진행 상황 업데이트
        for i, (e_field, polarization) in enumerate(zip(sorted_e_field_list_PUND, sorted_polarization_list_PUND)):
            ax.plot(e_field, polarization, color=colors[i], label=sorted_file_names_PUND[i], linewidth=5)

        # 축 라벨 및 스타일 설정
        ax.set_xlabel('E-field [MV/cm]', fontsize=28, fontweight='bold')
        ax.set_ylabel('Polarization [μC/cm²]', fontsize=28, fontweight='bold')
        ax.tick_params(axis='both', which='major', labelsize=22, direction='in')

        # x축 값 폰트 스타일 변경
        ax.set_xticklabels(ax.get_xticks(), fontsize=22, fontweight='bold')
        ax.set_yticklabels(ax.get_yticks(), fontsize=22, fontweight='bold')

        # 눈금 값의 유효숫자 조정
        ax.xaxis.set_major_formatter(FuncFormatter(self.msec_formatter))
        ax.yaxis.set_major_formatter(FuncFormatter(self.msec_formatter))

        # 제목 및 레전드 설정
        ax.set_title(f'{default_file_name_PUND}', fontsize=16, fontweight='bold')
        ax.legend(loc='upper left', prop={'weight': 'bold', 'size': 12}, borderaxespad=0.1, frameon=False)

        # 그래프 틀의 선 두께 변경
        for spine in ax.spines.values():
            spine.set_linewidth(2)

        # 그리드 추가
        ax.grid(True, which='both', linestyle='--')
        ax.minorticks_on()
        ax.grid(True, which='minor', linestyle=':')

        # Pr_max와 Ec, Ps 값 표시 (bold체)
        table_data = []
        for fn, ps, pr, ec in zip(sorted_file_names_PUND, sorted_ps_list_PUND, sorted_pr_max_list_PUND, sorted_ec_list_PUND):
            pr_str = f"{pr:.2f}" if pr is not None else "N/A"
            ec_str = f"{ec:.2f}" if ec is not None else "N/A"
            table_data.append([fn, f"{ps:.2f}", pr_str, ec_str])

        # 표 생성 및 크기 조정
        table_width = 1.0  # 표의 전체 너비
        table = ax.table(cellText=table_data, colLabels=['File Name', 'Ps [μC/cm²]', 'Pr [μC/cm²]', 'Ec [MV/cm]'], 
                        loc='bottom', bbox=[(1 - table_width) / 2, -0.75, table_width, 0.5])  # left 값을 중앙으로 조정
        table.auto_set_font_size(True)
        table.auto_set_column_width(col=list(range(len(table_data[0]))))

        # 텍스트 정렬
        for key, cell in table.get_celld().items():
            cell.set_text_props(ha='center', va='center', weight='bold', fontsize=15)

        # 그래프와 표 사이의 간격 추가 조정
        fig.subplots_adjust(bottom=0.55)

        # 그래프 저장
        plot_file_path = os.path.join(new_folder_path, f'{timestamp}_PUND_W{self.device_width:.0f}xH{self.device_height:.0f}xt{self.device_thickness:.0f}_plot.png')
        fig.savefig(plot_file_path, dpi=300)

        # Matplotlib을 사용하여 그래프 생성 및 저장
        def plot_graphs():
            fig1, ax1 = plt.subplots(figsize=(24, 16))
            colors = plt.cm.viridis(np.linspace(0, 1, len(sorted_time_list_IV)))

            # 왼쪽 y축에 Voltage 데이터 추가
            for i, (time, voltage) in enumerate(zip(sorted_time_list_IV, sorted_voltage_list_IV)):
                ax1.plot(time * 1e3, voltage, color='gray', linestyle='--', label=sorted_file_names_IV[i], linewidth=3)
                update_progress(i, len(sorted_time_list_IV), 25)

            ax1.set_xlabel('Time [msec]', fontsize=52, fontweight='bold')
            ax1.set_ylabel('Voltage [V]', fontsize=52, fontweight='bold', labelpad=10)
            ax1.set_xticklabels(ax1.get_xticks(), fontsize=22, fontweight='bold')
            ax1.set_yticklabels(ax1.get_yticks(), fontsize=22, fontweight='bold')

            ax1.xaxis.set_major_formatter(FuncFormatter(self.msec_formatter))
            ax1.yaxis.set_major_formatter(FuncFormatter(self.msec_formatter))

            #ax1.yaxis.set_label_coords(-0.05, 0.5)

            format_axes(ax1)

            # 오른쪽 y축 추가
            ax2 = ax1.twinx()
            for i, (time, current) in enumerate(zip(sorted_time_list_IV, sorted_current_list_IV)):
                ax2.plot(time * 1e3, current * 1e6, color=colors[i], label=f'{sorted_file_names_IV[i]}', linewidth=5)
                update_progress(i, len(sorted_current_list_IV), 50)

            ax2.set_ylabel('Current [μA]', fontsize=52, fontweight='bold')
            ax2.set_yticklabels(ax2.get_xticks(), fontsize=14, fontweight='bold')
            ax2.yaxis.set_major_formatter(FuncFormatter(self.msec_formatter))
            ax2.legend(loc='upper right', prop={'weight':'bold', 'size':21}, borderaxespad=0.1, frameon=False)

            format_axes(ax2)

            # 그래프 저장
            plot_file_path = os.path.join(new_folder_path, f'{timestamp}_t-IV_W{self.device_width:.0f}xH{self.device_height:.0f}xt{self.device_thickness:.0f}_plot.png')
            fig1.savefig(plot_file_path, dpi=300)

            # 두 번째 그래프 생성
            fig2, ax21 = plt.subplots(figsize=(24, 16))
            for i, (voltage, current) in enumerate(zip(sorted_voltage_list_IV, sorted_current_list_IV)):
                ax21.plot(voltage, current * 1e6, color=colors[i], label=sorted_file_names_IV[i], linewidth=5)
                update_progress(i, len(sorted_current_list_IV), 75)

            ax21.set_xlabel('Voltage [V]', fontsize=52, fontweight='bold')
            ax21.set_ylabel('Current [μA]', fontsize=52, fontweight='bold')
            ax21.set_xticklabels(ax21.get_xticks(), fontsize=14, fontweight='bold')
            ax21.set_yticklabels(ax21.get_xticks(), fontsize=14, fontweight='bold')
            # x축 눈금 값의 유효숫자 조정
            ax21.xaxis.set_major_formatter(FuncFormatter(self.msec_formatter))
            # y축 눈금 값의 유효숫자 조정
            ax21.yaxis.set_major_formatter(FuncFormatter(self.msec_formatter))
            ax21.legend(loc='upper left', prop={'weight':'bold', 'size':21}, borderaxespad=0.1, frameon=False)

            format_axes(ax21)

            # 그래프 저장
            plot_file_path2 = os.path.join(new_folder_path, f'{timestamp}_VI_W{self.device_width:.0f}xH{self.device_height:.0f}xt{self.device_thickness:.0f}_plot.png')
            fig2.savefig(plot_file_path2, dpi=300)

        def update_progress(i, total, base_progress):
            progress = base_progress + int((i + 1) / total * 25)
            self.progress_label.setText(f"Progressing...... {int(progress)}%")
            self.progress_bar.setValue(progress)
            QApplication.processEvents()  # UI 업데이트

        def format_axes(ax):
            ax.tick_params(axis='both', which='major', labelsize=40, direction='in')
            ax.spines['left'].set_linewidth(2)
            ax.spines['right'].set_linewidth(2)
            ax.spines['top'].set_linewidth(2)
            ax.spines['bottom'].set_linewidth(2)
            ax.grid(True, which='both', linestyle='--')
            ax.minorticks_on()
            ax.grid(True, which='minor', linestyle=':')

        plot_graphs()

        # 진행 상황 업데이트
        self.progress_bar.setValue(100)
        self.progress_label.setText("Progress: 100%")
        QApplication.processEvents()  # UI 업데이트

        # 파일 저장 성공 메시지 표시
        QMessageBox.information(self, "Files Saved", f"New Excel file '{output_file_path_PUND}' and associated plot files have been created in '{parent_dir}'.")

        # 진행 상황 창 닫기
        progress_widget.close()

        # 그래프 파일 열기
        if os.name == 'nt':  # Windows
            os.startfile(plot_file_path)
        elif os.name == 'posix':  # macOS or Linux
            os.system(f'open "{plot_file_path}"')  # macOS
            # os.system(f'xdg-open "{plot_file_path}"')  # Linux

        # 드래그앤드롭 창 유지
        self.show()

    def save_to_single_file_Tri(self, all_data, file_names):
        # 현재 날짜와 시간 가져오기
        # 현재 날짜와 시간 가져오기
        now = datetime.now()
        timestamp = now.strftime('%y%m%d_%H%M%S')

        # 기본 파일 이름 설정
        default_file_name_Tri = f'{timestamp}_Tri_W{self.device_width:.0f}xH{self.device_height:.0f}xt{self.device_thickness:.0f}'
        default_file_name_IV = f'{timestamp}_IV_W{self.device_width:.0f}xH{self.device_height:.0f}xt{self.device_thickness:.0f}'

        # 파일 저장 경로 선택 대화 상자 열기
        # 파일 저장 경로 선택 대화 상자 열기
        # —————————————————————————————————————————————————————————————
        # 1) "폴더+이름"을 한 번에 선택/입력할 다이얼로그 생성
        dlg = QFileDialog(self, "Select folder and enter name", self.last_save_dir)
        dlg.setAcceptMode(QFileDialog.AcceptSave)           # 저장 모드
        dlg.setFileMode(QFileDialog.Directory)              # 디렉토리 선택
        dlg.setOption(QFileDialog.ShowDirsOnly, True)       # 파일은 숨기고 폴더만
        dlg.selectFile(default_file_name_Tri)               # 기본으로 채워둘 새 폴더명

        if dlg.exec_() != QDialog.Accepted:
            QMessageBox.information(self, "Saving Canceled", "Saving file has been canceled.")
            return

        # 2) 선택된 전체 경로 → 부모 디렉토리와 폴더명 분리
        selected = dlg.selectedFiles()[0]                   # e.g. "C:/Users/.../20250521_PUND_..."
        parent_dir = os.path.dirname(selected)
        folder_name = os.path.basename(selected)

        # 3) 결과를 저장 디렉토리로 사용
        self.last_save_dir = parent_dir
        new_folder_path = selected
        os.makedirs(new_folder_path, exist_ok=True)
        # —————————————————————————————————————————————————————————————

        # Excel 파일 경로 설정
        output_file_path_Tri = os.path.join(new_folder_path, f"{default_file_name_Tri}.xlsx")
        output_file_path_IV = os.path.join(new_folder_path, f"{default_file_name_IV}.xlsx")

        with pd.ExcelWriter(output_file_path_Tri, engine='xlsxwriter') as writer:
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
            ps_list = []
            pr_max_list = []
            ec_list = []

            for k, sheet_df in enumerate(all_data):
                if sheet_df is not None:
                    ec1_values, ec2_values, pr1_values = [], [], []

                    for j in range(len(sheet_df['Charge']) - 1):
                        # E-field의 연속한 두 데이터의 곱이 음수인 경우
                        if sheet_df['E-field'][j] * sheet_df['E-field'][j + 1] < 0:
                            e1, e2 = sheet_df['E-field'][j], sheet_df['E-field'][j + 1]
                            p1, p2 = sheet_df['Charge'][j], sheet_df['Charge'][j + 1]
                            if e1 != e2:
                                m = (p2 - p1) / (e2 - e1)
                                b = abs(p1 - m * e1)
                                pr1_values.append(b)

                        # Polarization의 연속한 두 데이터의 곱이 음수인 경우
                        if sheet_df['Charge'][j] * sheet_df['Charge'][j + 1] < 0:
                            e3, e4 = sheet_df['E-field'][j], sheet_df['E-field'][j + 1]
                            p3, p4 = sheet_df['Charge'][j], sheet_df['Charge'][j + 1]
                            if p3 != p4:
                                m = (e4 - e3) / (p4 - p3)
                                b = e3 - m * p3
                                if p3 > 0:
                                    ec1_values.append(b)
                                else:
                                    ec2_values.append(b)
                    # ───── 여기서부터 수정 ─────
                    # first_valid_index / last_valid_index 사용하여
                    # 실제 마지막 유효 데이터로 절편 계산
                    idx_e_first = sheet_df['E-field'].first_valid_index()
                    idx_e_last  = sheet_df['E-field'].last_valid_index()
                    idx_p_first = sheet_df['Charge'].first_valid_index()
                    idx_p_last  = sheet_df['Charge'].last_valid_index()

                    if idx_e_first is not None and idx_e_last is not None \
                       and idx_p_first is not None and idx_p_last is not None:
                        e_start = sheet_df.at[idx_e_first, 'E-field']
                        e_end   = sheet_df.at[idx_e_last,   'E-field']
                        p_start = sheet_df.at[idx_p_first, 'Charge']
                        p_end   = sheet_df.at[idx_p_last,   'Charge']
                        if e_start != e_end:
                            m_end = (p_end - p_start) / (e_end - e_start)
                            b_end = abs(p_start - m_end * e_start)
                            pr1_values.append(b_end)
                    # ───────────────────────────

                    # 'None' 값을 제거하고 평균 계산
                    pr1_values = [value for value in pr1_values if value is not None]
                    ec1_values = [value for value in ec1_values if value is not None]
                    ec2_values = [value for value in ec2_values if value is not None]

                    # Ps, Pr_max, Ec 계산
                    pol_max = abs(max(sheet_df['Charge']))
                    pol_min = abs(min(sheet_df['Charge']))
                    ps = (pol_max + pol_min) / 2
                    pr_max = abs(np.mean(pr1_values)) if pr1_values else None
                    ec = (abs(np.mean(ec1_values)) + abs(np.mean(ec2_values))) / 2 if ec1_values and ec2_values else None

                    ps_list.append(ps)
                    pr_max_list.append(pr_max)
                    ec_list.append(ec)

                    # DataFrame에 추가
                    pr_ec_df[f'Ps_{k + 1}'] = [ps]
                    pr_ec_df[f'Pr_max_{k + 1}'] = [pr_max]
                    pr_ec_df[f'Ec_{k + 1}'] = [ec]

            # 'Pr_Ec' 시트에 데이터 저장
            pr_ec_df.to_excel(writer, sheet_name='Ps_Pr_Ec', index=False)

        with pd.ExcelWriter(output_file_path_IV, engine='xlsxwriter') as writer:
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

        # 레이아웃 생성
        progress_layout = QVBoxLayout()
        progress_layout.setContentsMargins(10, 10, 10, 10)

        self.progress_bar = QProgressBar()
        self.progress_label = QLabel("Progressing...... 0%")

        # 진행 바 설정
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setRange(0, 100)

        self.progress_label.setAlignment(Qt.AlignCenter)
        self.progress_bar.setAlignment(Qt.AlignCenter)

        # 프로그레스 바를 감싸는 레이아웃 생성
        progress_layout.addWidget(self.progress_bar)
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
        QApplication.processEvents()  # UI 업데이트

        # 파일 이름과 관련된 데이터 리스트를 함께 묶기
        combined_data_Tri = list(zip(file_names, e_field_list, charge_list, ps_list, pr_max_list, ec_list))
        combined_data_IV = list(zip(file_names, time_list, voltage_list, current_list))

        # 파일 이름을 기준으로 정렬
        combined_data_Tri.sort(key=lambda x: x[0])  # x[0]은 file_name
        combined_data_IV.sort(key=lambda x: x[0])  # x[0]은 file_name

        # 정렬된 데이터 분리
        sorted_file_names_Tri, sorted_e_field_list_Tri, sorted_charge_list_Tri, sorted_ps_list_Tri, sorted_pr_max_list_Tri, sorted_ec_list_Tri = zip(*combined_data_Tri)
        sorted_file_names_IV, sorted_time_list_IV, sorted_voltage_list_IV, sorted_current_list_IV = zip(*combined_data_IV)

        # Matplotlib을 사용하여 그래프 생성 및 저장
        fig, ax = plt.subplots(figsize=(10, 16))

        # 그래프 색상 설정
        colors = plt.cm.viridis(np.linspace(0, 1, len(sorted_e_field_list_Tri)))

        # 그래프 그리기 및 진행 상황 업데이트
        for i, (e_field, charge) in enumerate(zip(sorted_e_field_list_Tri, sorted_charge_list_Tri)):
            ax.plot(e_field, charge, color=colors[i], label=sorted_file_names_Tri[i], linewidth=5)

        # 축 라벨 및 스타일 설정
        ax.set_xlabel('E-field [MV/cm]', fontsize=28, fontweight='bold')
        ax.set_ylabel('Charge [μC/cm²]', fontsize=28, fontweight='bold')
        ax.tick_params(axis='both', which='major', labelsize=22, direction='in')

        # x축 값 폰트 스타일 변경
        ax.set_xticklabels(ax.get_xticks(), fontsize=22, fontweight='bold')
        ax.set_yticklabels(ax.get_yticks(), fontsize=22, fontweight='bold')

        # 눈금 값의 유효숫자 조정
        ax.xaxis.set_major_formatter(FuncFormatter(self.msec_formatter))
        ax.yaxis.set_major_formatter(FuncFormatter(self.msec_formatter))

        # 제목 및 레전드 설정
        ax.set_title(f'{default_file_name_Tri}', fontsize=16, fontweight='bold')
        ax.legend(loc='upper left', prop={'weight': 'bold', 'size': 12}, borderaxespad=0.1, frameon=False)

        # 그래프 틀의 선 두께 변경
        for spine in ax.spines.values():
            spine.set_linewidth(2)

        # 그리드 추가
        ax.grid(True, which='both', linestyle='--')
        ax.minorticks_on()
        ax.grid(True, which='minor', linestyle=':')

        # Pr_max와 Ec, Ps 값 표시 (bold체)
        table_data = []
        for fn, ps, pr, ec in zip(sorted_file_names_Tri, sorted_ps_list_Tri, sorted_pr_max_list_Tri, sorted_ec_list_Tri):
            pr_str = f"{pr:.2f}" if pr is not None else "N/A"
            ec_str = f"{ec:.2f}" if ec is not None else "N/A"
            table_data.append([fn, f"{ps:.2f}", pr_str, ec_str])

        # 표 생성 및 크기 조정
        table_width = 1.0  # 표의 전체 너비
        table = ax.table(cellText=table_data, colLabels=['File Name', 'Ps [μC/cm²]', 'Pr [μC/cm²]', 'Ec [MV/cm]'], 
                        loc='bottom', bbox=[(1 - table_width) / 2, -0.75, table_width, 0.5])  # left 값을 중앙으로 조정
        table.auto_set_font_size(True)
        table.auto_set_column_width(col=list(range(len(table_data[0]))))

        # 텍스트 정렬
        for key, cell in table.get_celld().items():
            cell.set_text_props(ha='center', va='center', weight='bold', fontsize=15)

        # 그래프와 표 사이의 간격 추가 조정
        fig.subplots_adjust(bottom=0.55)

        # 그래프 저장
        plot_file_path = os.path.join(new_folder_path, f'{timestamp}_Tri_W{self.device_width:.0f}xH{self.device_height:.0f}xt{self.device_thickness:.0f}_plot.png')
        fig.savefig(plot_file_path, dpi=300)

        # Matplotlib을 사용하여 그래프 생성 및 저장
        def plot_graphs():
            fig1, ax1 = plt.subplots(figsize=(24, 16))
            colors = plt.cm.viridis(np.linspace(0, 1, len(sorted_time_list_IV)))

            # 왼쪽 y축에 Voltage 데이터 추가
            for i, (time, voltage) in enumerate(zip(sorted_time_list_IV, sorted_voltage_list_IV)):
                ax1.plot(time * 1e3, voltage, color='gray', linestyle='--', label=sorted_file_names_IV[i], linewidth=3)
                update_progress(i, len(sorted_time_list_IV), 25)

            ax1.set_xlabel('Time [msec]', fontsize=52, fontweight='bold')
            ax1.set_ylabel('Voltage [V]', fontsize=52, fontweight='bold', labelpad=10)
            ax1.set_xticklabels(ax1.get_xticks(), fontsize=22, fontweight='bold')
            ax1.set_yticklabels(ax1.get_yticks(), fontsize=22, fontweight='bold')

            ax1.xaxis.set_major_formatter(FuncFormatter(self.msec_formatter))
            ax1.yaxis.set_major_formatter(FuncFormatter(self.msec_formatter))

            #ax1.yaxis.set_label_coords(-0.05, 0.5)

            format_axes(ax1)

            # 오른쪽 y축 추가
            ax2 = ax1.twinx()
            for i, (time, current) in enumerate(zip(sorted_time_list_IV, sorted_current_list_IV)):
                ax2.plot(time * 1e3, current * 1e6, color=colors[i], label=f'{sorted_file_names_IV[i]}', linewidth=5)
                update_progress(i, len(sorted_current_list_IV), 50)

            ax2.set_ylabel('Current [μA]', fontsize=52, fontweight='bold')
            ax2.set_yticklabels(ax2.get_xticks(), fontsize=14, fontweight='bold')
            ax2.yaxis.set_major_formatter(FuncFormatter(self.msec_formatter))
            ax2.legend(loc='upper right', prop={'weight':'bold', 'size':21}, borderaxespad=0.1, frameon=False)

            format_axes(ax2)

            # 그래프 저장
            plot_file_path = os.path.join(new_folder_path, f'{timestamp}_t-IV_W{self.device_width:.0f}xH{self.device_height:.0f}xt{self.device_thickness:.0f}_plot.png')
            fig1.savefig(plot_file_path, dpi=300)

            # 두 번째 그래프 생성
            fig2, ax21 = plt.subplots(figsize=(24, 16))
            for i, (voltage, current) in enumerate(zip(sorted_voltage_list_IV, sorted_current_list_IV)):
                ax21.plot(voltage, current * 1e6, color=colors[i], label=sorted_file_names_IV[i], linewidth=5)
                update_progress(i, len(sorted_current_list_IV), 75)

            ax21.set_xlabel('Voltage [V]', fontsize=52, fontweight='bold')
            ax21.set_ylabel('Current [μA]', fontsize=52, fontweight='bold')
            ax21.set_xticklabels(ax21.get_xticks(), fontsize=14, fontweight='bold')
            ax21.set_yticklabels(ax21.get_xticks(), fontsize=14, fontweight='bold')
            # x축 눈금 값의 유효숫자 조정
            ax21.xaxis.set_major_formatter(FuncFormatter(self.msec_formatter))
            # y축 눈금 값의 유효숫자 조정
            ax21.yaxis.set_major_formatter(FuncFormatter(self.msec_formatter))
            ax21.legend(loc='upper left', prop={'weight':'bold', 'size':21}, borderaxespad=0.1, frameon=False)

            format_axes(ax21)

            # 그래프 저장
            plot_file_path2 = os.path.join(new_folder_path, f'{timestamp}_VI_W{self.device_width:.0f}xH{self.device_height:.0f}xt{self.device_thickness:.0f}_plot.png')
            fig2.savefig(plot_file_path2, dpi=300)

        def update_progress(i, total, base_progress):
            progress = base_progress + int((i + 1) / total * 25)
            self.progress_label.setText(f"Progressing...... {int(progress)}%")
            self.progress_bar.setValue(progress)
            QApplication.processEvents()

        def format_axes(ax):
            ax.tick_params(axis='both', which='major', labelsize=40, direction='in')
            ax.spines['left'].set_linewidth(2)
            ax.spines['right'].set_linewidth(2)
            ax.spines['top'].set_linewidth(2)
            ax.spines['bottom'].set_linewidth(2)
            ax.grid(True, which='both', linestyle='--')
            ax.minorticks_on()
            ax.grid(True, which='minor', linestyle=':')

        plot_graphs()
        
        # 진행 상황 업데이트
        self.progress_bar.setValue(100)
        self.progress_label.setText("Progress: 100%")
        QApplication.processEvents()  # UI 업데이트

        # 파일 저장 성공 메시지 표시
        QMessageBox.information(self, "Files Saved", f"New Excel file '{output_file_path_Tri}' and associated plot files have been created in '{parent_dir}'.")

        # 진행 상황 창 닫기
        progress_widget.close()

        # 그래프 파일 열기
        if os.name == 'nt':  # Windows
            os.startfile(plot_file_path)
        elif os.name == 'posix':  # macOS or Linux
            os.system(f'open "{plot_file_path}"')  # macOS
            # os.system(f'xdg-open "{plot_file_path}"')  # Linux
            
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
        if not output_dir:
            QMessageBox.information(self, "Saving Canceled", "Saving file has been canceled.")
            return

        # 새 폴더 생성
        new_folder_path = os.path.join(output_dir, default_file_name)
        os.makedirs(new_folder_path, exist_ok=True)

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
        progress_widget.setStyleSheet("background-color: #f0f0f0; border: 2px solid #0078d7; border-radius: 10px;")

        progress_layout = QVBoxLayout()
        progress_layout.setContentsMargins(10, 10, 10, 10)

        self.progress_bar = QProgressBar()
        self.progress_label = QLabel("Progressing...... 0%")
        self.progress_bar.setMinimum(0)
        self.progress_bar.setMaximum(100)
        self.progress_bar.setAlignment(Qt.AlignCenter)
        self.progress_label.setAlignment(Qt.AlignCenter)

        progress_layout.addWidget(self.progress_bar)
        progress_layout.addWidget(self.progress_label)
        progress_widget.setLayout(progress_layout)

        # 애니메이션 효과 추가
        animation = QPropertyAnimation(progress_widget, b"geometry")
        animation.setDuration(500)
        animation.setStartValue(QRect(750, 450, 300, 0))
        animation.setEndValue(QRect(750, 450, 300, 150))
        animation.start()

        progress_widget.show()
        QApplication.processEvents()  # UI 업데이트

        # 파일 이름과 관련된 데이터 리스트를 함께 묶기
        combined_data = list(zip(file_names, time_list, voltage_list, current_list))

        # 파일 이름을 기준으로 정렬
        combined_data.sort(key=lambda x: x[0])  # x[0]은 file_name

        # 정렬된 데이터 분리
        sorted_file_names, sorted_time_list, sorted_voltage_list, sorted_current_list = zip(*combined_data)

        # Matplotlib을 사용하여 그래프 생성 및 저장
        def plot_graphs():
            fig, ax1 = plt.subplots(figsize=(24, 16))
            colors = plt.cm.viridis(np.linspace(0, 1, len(sorted_time_list)))

            # 왼쪽 y축에 Voltage 데이터 추가
            for i, (time, voltage) in enumerate(zip(sorted_time_list, sorted_voltage_list)):
                ax1.plot(time * 1e3, voltage, color='gray', linestyle='--', label=sorted_file_names[i], linewidth=3)
                update_progress(i, len(sorted_time_list), 25)

            ax1.set_xlabel('Time [msec]', fontsize=52, fontweight='bold')
            ax1.set_ylabel('Voltage [V]', fontsize=52, fontweight='bold')
            ax1.set_xticklabels(ax1.get_xticks(), fontsize=22, fontweight='bold')
            ax1.set_yticklabels(ax1.get_yticks(), fontsize=22, fontweight='bold')

            ax1.xaxis.set_major_formatter(FuncFormatter(self.msec_formatter))
            ax1.yaxis.set_major_formatter(FuncFormatter(self.msec_formatter))

            #ax1.yaxis.set_label_coords(-0.05, 0.5)

            format_axes(ax1)

            # 오른쪽 y축 추가
            ax2 = ax1.twinx()
            for i, (time, current) in enumerate(zip(sorted_time_list, sorted_current_list)):
                ax2.plot(time * 1e3, current * 1e6, color=colors[i], label=f'{sorted_file_names[i]}', linewidth=5)
                update_progress(i, len(sorted_current_list), 50)

            ax2.set_ylabel('Current [μA]', fontsize=52, fontweight='bold')
            ax2.set_yticklabels(ax2.get_xticks(), fontsize=14, fontweight='bold')
            ax2.yaxis.set_major_formatter(FuncFormatter(self.msec_formatter))
            ax2.legend(loc='upper right', prop={'weight':'bold', 'size':21}, borderaxespad=0.1, frameon=False)

            format_axes(ax2)

            # 그래프 저장
            plot_file_path = os.path.join(new_folder_path, f'{timestamp}_t-IV_W{self.device_width:.0f}xH{self.device_height:.0f}xt{self.device_thickness:.0f}_plot.png')
            fig.savefig(plot_file_path, dpi=300)

            # 두 번째 그래프 생성
            fig2, ax21 = plt.subplots(figsize=(24, 16))
            for i, (voltage, current) in enumerate(zip(sorted_voltage_list, sorted_current_list)):
                ax21.plot(voltage, current * 1e6, color=colors[i], label=sorted_file_names[i], linewidth=5)
                update_progress(i, len(sorted_current_list), 75)

            ax21.set_xlabel('Voltage [V]', fontsize=52, fontweight='bold')
            ax21.set_ylabel('Current [μA]', fontsize=52, fontweight='bold')
            ax21.set_xticklabels(ax21.get_xticks(), fontsize=14, fontweight='bold')
            ax21.set_yticklabels(ax21.get_xticks(), fontsize=14, fontweight='bold')
            # x축 눈금 값의 유효숫자 조정
            ax21.xaxis.set_major_formatter(FuncFormatter(self.msec_formatter))
            # y축 눈금 값의 유효숫자 조정
            ax21.yaxis.set_major_formatter(FuncFormatter(self.msec_formatter))
            ax21.legend(loc='upper left', prop={'weight':'bold', 'size':21}, borderaxespad=0.1, frameon=False)

            format_axes(ax21)

            # 그래프 저장
            plot_file_path2 = os.path.join(new_folder_path, f'{timestamp}_VI_W{self.device_width:.0f}xH{self.device_height:.0f}xt{self.device_thickness:.0f}_plot.png')
            fig2.savefig(plot_file_path2, dpi=300)

        def update_progress(i, total, base_progress):
            progress = base_progress + int((i + 1) / total * 25)
            self.progress_label.setText(f"Progressing...... {int(progress)}%")
            self.progress_bar.setValue(progress)
            QApplication.processEvents()

        def format_axes(ax):
            ax.tick_params(axis='both', which='major', labelsize=40, direction='in')
            ax.spines['left'].set_linewidth(2)
            ax.spines['right'].set_linewidth(2)
            ax.spines['top'].set_linewidth(2)
            ax.spines['bottom'].set_linewidth(2)
            ax.grid(True, which='both', linestyle='--')
            ax.minorticks_on()
            ax.grid(True, which='minor', linestyle=':')

        plot_graphs()

        # 진행 상황 업데이트
        self.progress_bar.setValue(100)
        self.progress_label.setText("Progress: 100%")
        QApplication.processEvents()

        # 파일 저장 성공 메시지 표시
        QMessageBox.information(self, "File Saved", f"New Excel file '{output_file_path}' has been created.")
        progress_widget.close()
        self.show()

if __name__ == '__main__':
    import sys, time
    from PyQt5.QtCore import Qt, QTimer
    from PyQt5.QtGui import QPixmap, QColor, QFont, QIcon
    from PyQt5.QtWidgets import QApplication, QSplashScreen

    app = QApplication(sys.argv)
    app.setWindowIcon(QIcon(':/KKH.ico'))

    # 1) SplashScreen 세팅
    splash_pix = QPixmap(400, 300)
    splash_pix.fill(QColor('#313335'))
    splash = QSplashScreen(
        splash_pix,
        Qt.WindowStaysOnTopHint | Qt.SplashScreen
    )
    splash.setFont(QFont('Segoe UI', 12))
    splash.showMessage('Loading... 0%', Qt.AlignCenter|Qt.AlignBottom, QColor('#78c2ad'))
    splash.show()
    app.processEvents()

    # 2) 로딩 퍼센티지 업데이트
    for i in range(1, 101):
        time.sleep(0.01)
        splash.showMessage(f'Loading... {i}%', Qt.AlignCenter, QColor('#78c2ad'))
        app.processEvents()

    # 3) 메인 윈도우 생성 (단, __init__ 에서는 check_for_updates, maybe_show_patch_notes 호출을 제거하세요)
    window = CSVFileProcessor()

    # 4) Splash 닫고 메인 윈도우 띄우기
    splash.finish(window)
    window.show()

    # 5) 이벤트 루프가 한 번 돌고 나서
    def startup_sequence():
        # 5-1) 첫 실행이라면 패치 노트
        window.maybe_show_patch_notes()
        # 5-2) 그리고 나서 업데이트 체크
        window.check_for_updates()

    QTimer.singleShot(0, startup_sequence)

    sys.exit(app.exec_())