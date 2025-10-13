import os
import sys
from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtWidgets import (
    QApplication,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QPushButton,
    QFileDialog,
    QMessageBox,
    QComboBox,
    QTextEdit,
    QProgressBar,
)
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright
import pandas as pd
import time
from pathlib import Path
from typing import Callable, Optional


def read_profiling_excel(filepath):
    df = pd.read_excel(filepath, index_col=0)
    return df

def login(page,context,username,password):
    page.goto("https://matchapro.web.bps.go.id/")

    try:
        page.get_by_text("Sign in with SSO BPS").click(timeout=5000)
        page.get_by_label("username").fill(username)
        page.get_by_label("password").fill(password)
        page.click("#kc-login")
        context.storage_state(path="state.json")
        print("Login Successful")
    except Exception as e:
        print("already logged in")
        print(e)

def update_profiling(page, idsbr, row, emit: Optional[Callable[[str], None]] = None):
    log = emit or print
    edit_button = page.locator(".btn-edit-perusahaan").first
    if edit_button.count() == 1:
        page.locator(".btn-edit-perusahaan").first.click()
        with page.expect_popup() as popup_info:
            page.get_by_text("Ya, edit!").click()
        new_page = popup_info.value
        time.sleep(5)
        new_page.get_by_label("Sumber Profiling").fill(str(row["Sumber profiling"]))
        new_page.get_by_placeholder("Catatan").fill(str(row["Catatan"]))
        value = str(row["Keberadaan usaha"]).strip().lower()
        locator = f'input[name="kondisi_usaha"][value="{value}"]'
        new_page.locator(locator).check()
        if value == "9":
            new_page.get_by_placeholder("IDSBR Master").fill(
                str(row["Idsbr duplikat"])
            )
        email_field = new_page.get_by_placeholder("Email")
        checkbox = new_page.locator("#check-email")
        email_value = email_field.input_value().strip()

        if not email_value :
            if checkbox.is_checked():
                checkbox.uncheck()
                print(f"unchecked email checkbox for {idsbr}")

        log(
            f"{idsbr} {str(row['Nama usaha'])} Sumber Profiling : {str(row['Sumber profiling'])}, Catatan : {str(row['Catatan'])},  status perusahaan {value}"
        )
        new_page.wait_for_timeout(1000)
        new_page.get_by_text("Submit Final").click(force=True)
        konsistensi = new_page.locator("#confirm-consistency")
        if konsistensi.count() == 1:
            konsistensi.click()    
        new_page.locator("button.swal2-confirm", has_text="Ya, Submit!").click()               
        new_page.wait_for_timeout(1000)
        new_page.close()
        time.sleep(5)

def load_sso():
    load_dotenv()

    username = os.getenv("USERNAME_SSO")
    password = os.getenv("PASSWORD")

    if not username or not password :
        raise ValueError("Masukkan Username dan Password di .env file!")
    
    return username,password

class App(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Matchapro Automate Entry")
        self.setMinimumWidth(760)

        # load file excel
        self.path_edit = QLineEdit(self)
        self.path_edit.setPlaceholderText("Pilih file excel: ")
        self.path_edit.setReadOnly(True)
        self.path_edit.setClearButtonEnabled(True)

        self.browse_button = QPushButton("Browse", self)
        self.browse_button.clicked.connect(self.browse_file)
        self.start_button = QPushButton("Start", self)
        self.stop_button = QPushButton("Stop", self)
        # self.load_button = QPushButton("Load", self)

        self.start_button.clicked.connect(self.start_worker)
        self.stop_button.clicked.connect(self.stop_worker)

        self.pilih_kab = QComboBox(self)
        self.pilih_kab.addItems([
            "[01] SUMBA BARAT",
            "[02] SUMBA TIMUR",
            "[03] KUPANG" ,
            "[04] TIMOR TENGAH SELATAN" ,
            "[05] TIMOR TENGAH UTARA" ,
            "[06] BELU",
            "[07] ALOR",
            "[08] LEMBATA",
            "[09] FLORES TIMUR",
            "[10] SIKKA",
            "[11] ENDE",
            "[12] NGADA",
            "[13] MANGGARAI",
            "[14] ROTE NDAO",
            "[15] MANGGARAI BARAT",
            "[16] SUMBA TENGAH",
            "[17] SUMBA BARAT DAYA",
            "[18] NAGEKEO",
            "[19] MANGGARAI TIMUR",
            "[20] SABU RAIJUA",
            "[21] MALAKA",
            "[71] KUPANG"  
        ])

        self.progress = QProgressBar(self)
        self.progress.setRange(0, 100)
        self.progress.setValue(0)

        self.log = QTextEdit(self)
        self.log.setReadOnly(True)

        top = QHBoxLayout()
        top.addWidget(QLabel("File Excel:", self))
        top.addWidget(self.path_edit,1)
        top.addWidget(self.browse_button)

        reg = QHBoxLayout()
        reg.addWidget(QLabel("Satker:", self))
        reg.addWidget(self.pilih_kab, 1)

        btns = QHBoxLayout()
        btns.addWidget(self.start_button)
        btns.addWidget(self.stop_button)

        root = QVBoxLayout(self)
        root.addLayout(top)
        root.addLayout(reg)
        root.addLayout(btns)
        root.addWidget(self.progress)
        root.addWidget(QLabel("Logs:", self))
        root.addWidget(self.log, 1)


    def browse_file(self):
        file_dialog = QFileDialog(self)
        file_dialog.setNameFilter("Excel Files (*.xlsx *.xls)")
        if file_dialog.exec():
            selected_file = file_dialog.selectedFiles()[0]
            self.path_edit.setText(selected_file)
            self.start_button.setEnabled(True)
        else:
            QMessageBox.warning(self, "Warning", "Upload file excel terlebih dahulu!")

    def start_worker(self):
        excel = self.path_edit.text().strip()
        if not excel:
            QMessageBox.warning(
                self, "Missing file", "Upload file excel terlebih dahulu!"
            )
            return
        if not Path(excel).exists():
            QMessageBox.warning(self, "Invalid file", "File Corrupted.")
            return
        
        kabupaten = self.pilih_kab.currentText()
        self.progress.setValue(0)
        self.log.clear()

        self.worker = Worker(excel, kabupaten)
        self.worker.log.connect(self.append_log)
        self.worker.progress.connect(self.progress.setValue)

        self.start_button.setEnabled(False)
        self.stop_button.setEnabled(True)
        self.worker.start()

    def stop_worker(self):
        if self.worker and self.worker.isRunning():
            self.worker.request_stop()
            self.append_log("Stop")
        self.stop_button.setEnabled(False)
        self.start_button.setEnabled(True)
    
    def append_log(self, msg: str):
        self.log.append(msg.rstrip("\n"))

    def worker_finished_ok(self):
        self.append_log("Finished.\n")
        self.start_button.setEnabled(True)
        self.stop_button.setEnabled(False)

    def worker_finished_err(self, err: str):
        self.append_log(f"Error: {err}\n")
        self.start_button.setEnabled(True)
        self.stop_button.setEnabled(False)
        
class Worker(QThread):
    log = Signal(str)
    progress = Signal(int)
    finished_ok = Signal()
    finished_err = Signal(str)

    def __init__(self, excel_path: str, kabupaten_text: str):
        super().__init__()
        self.excel_path = excel_path
        self.kabupaten_text = kabupaten_text
        self._stop_requested = False

    def request_stop(self):
        self._stop_requested = True

    def _emit(self, msg: str):
        print(msg, end="" if msg.endswith("\n") else "\n")
        self.log.emit(msg if msg.endswith("\n") else msg + "\n")

    def run(self):
        try:
            from playwright.sync_api import sync_playwright

            load_dotenv()
            username,password = load_sso()
            self._emit(f"using SSO: {username}\n")

            df = read_profiling_excel(self.excel_path)
            total = len(df)
            if total == 0:
                self._emit("File kosong!\n")
                self.finished_ok.emit()
                return

            self._emit(f"Loaded Excel: {self.excel_path} ({total} baris direktori) untuk Satker {self.kabupaten_text}\n")

            with sync_playwright() as p:
                browser = p.chromium.launch(headless=False)
                context = browser.new_context(
                    storage_state="state.json" if os.path.exists("state.json") else None
                )
                page = context.new_page()

                self._emit("Login...\n")
                login(page, context, username, password)

                page.goto("https://matchapro.web.bps.go.id/direktori-usaha")
                page.click("text=Skip")
                page.locator("#select2-f_provinsi-container").click()
                page.locator(
                    ".select2-results__option", has_text="[53] NUSA TENGGARA TIMUR"
                ).click()
                page.locator("#select2-f_kabupaten-container").click()
                page.locator(".select2-results__option", has_text=f"{self.kabupaten_text}").click()

                # tail_df = df.tail(50)
                for idx, (idsbr, row) in enumerate(df.iterrows(), start=1):
                    if self._stop_requested:
                        self._emit("Stop. Exiting loop...\n")
                        break

                    page.locator('[name="idsbr"]').fill(str(idsbr))
                    self._emit(f"Mengisi {idsbr} - {row['Nama usaha']}\n")
                    time.sleep(5)

                    if page.get_by_label("Lihat History Profiling").count() == 1:
                        history_profiling = page.get_by_label(
                            "Lihat History Profiling"
                        ).first
                        history_profiling.click()

                        page.locator(
                            ".modal-body > .blockUI.blockMsg.blockElement"
                        ).wait_for(
                            state="detached",
                        )

                        status = page.locator(
                            "#table-history-profiling span.badge.rounded-pill"
                        ).first
                        status_text = status.inner_text().lower()
                        self._emit(f"Status: {status_text}\n")

                        if status_text == "submitted" or status_text == "approved":
                            self._emit(f"{idsbr} - {row['Nama usaha']} sudah submit\n")
                            page.wait_for_timeout(1000)
                            page.locator(
                                "#modal-view-history-profiling button", has_text="Close"
                            ).click(force=True)
                        else:
                            self._emit(f"{idsbr} - {row['Nama usaha']} belum submit\n")
                            page.wait_for_timeout(1000)
                            page.locator(
                                "#modal-view-history-profiling button", has_text="Close"
                            ).click(force=True)
                            page.wait_for_timeout(1000)
                            update_profiling(page, idsbr, row, emit=self._emit)
                    else:
                        self._emit("Open\n")
                        page.wait_for_timeout(1000)
                        update_profiling(page, idsbr, row)

                    # update progress
                    self.progress.emit(int(idx / total * 100))

                self._emit("Done. Leaving browser open a moment.\n")
                self.finished_ok.emit()
                # browser.close()

            if self._stop_requested:
                self.finished_ok.emit()
            else:
                self.finished_ok.emit()

        except Exception as e:
            self.finished_err.emit(str(e))


def main():
    print("Hello from matchapro!")
    
    p =  sync_playwright().start()
    browser = p.chromium.launch(headless=False)
    context = browser.new_context(storage_state="state.json" if os.path.exists("state.json") else None)
    page = context.new_page()

    username,password = load_sso()
    login(page,context,username,password)

    page.goto("https://matchapro.web.bps.go.id/direktori-usaha")
    page.click("text=Skip")

    page.locator("#select2-f_provinsi-container").click()
    page.locator(".select2-results__option", has_text="[53] NUSA TENGGARA TIMUR").click()
    page.locator("#select2-f_kabupaten-container").click()
    page.locator(".select2-results__option", has_text="[71] KUPANG").click()

    df = read_profiling_excel("Direktori\\export-550-directories.xlsx")
    for idsbr, row in df.iterrows():
        page.locator('[name="idsbr"]').fill(str(idsbr))
        print("Mengisi", idsbr, "-", row["Nama usaha"])
        time.sleep(5)
        if page.get_by_label("Lihat History Profiling").count() == 1:
            history_profiling = page.get_by_label("Lihat History Profiling").first
            history_profiling.click()
            page.locator(".modal-body > .blockUI.blockMsg.blockElement").wait_for(
                state="detached", 
            )
            status = page.locator("#table-history-profiling span.badge.rounded-pill").first
            status_text = status.inner_text().lower()
            print(status_text)
            if status_text == "submitted" or status_text == "approved":
                print(idsbr, "-", row["Nama usaha"], " sudah submit")
                page.wait_for_timeout(1000)
                page.locator('#modal-view-history-profiling button', has_text="Close").click(force=True)
            else:
                print(idsbr, "-", row["Nama usaha"], " belum submit")
                page.wait_for_timeout(1000)
                page.locator('#modal-view-history-profiling button', has_text="Close").click(force=True)
                page.wait_for_timeout(1000)
                update_profiling(page,idsbr,row)
        else:
            print("Open")
            page.wait_for_timeout(1000)
            update_profiling(page,idsbr,row)
    input("Browser is open. Press Enter to exit...")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = App()
    window.show()
    sys.exit(app.exec())
    # main()