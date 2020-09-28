from PyQt5 import QtCore, QtWidgets
from PyQt5.QtGui import QPixmap
from templates.Ui_template import Ui_Consult2
from Consult_Main import MainExecute
from utils.automail import AutoEmail
from utils.open_zip import open_zipfile
from Closing_plots import MakeClosingPlots
from Append_DB import MainDB
from Int_BR import *
from utils.export import export


class MyMainWindow(QtWidgets.QMainWindow, Ui_Consult2):
    def __init__(self, parent=None):
        super(MyMainWindow, self).__init__(parent)
        self.message = None
        self.start = None
        self.db_start = None
        self.end = None
        self.setupUi(self)
        self.obj_group()
        self.cst_group()
        self.db_group()
        self.intbr_group()
        self.intbr_raw_group()
        self.mail_group()
        self.attachment = None
        self.plot = None
        self.obj = None
        self.plant = None
        self.optimal = 0
        zip_ = r'Cookies\Pocket.zip'
        penguins = open_zipfile(zip_, 'Penguins.jpg')
        self.image.setPixmap(QPixmap(penguins))
        self.image_2.setPixmap(QPixmap(penguins))

    # def _decorator(func):
    #     def wrapper(self):
    #         os.chdir(os.path.join(Path(os.getcwd()).parent))
    #         func(self)
    #         os.chdir(os.path.join(os.getcwd(), 'GUI'))
    #     return wrapper

    def intbr_raw_group(self):
        self.start_m_3.setDate(QtCore.QDate.currentDate())
        self.end_m_3.setDate(QtCore.QDate.currentDate())
        self.br_rawbtn.clicked.connect(self.br_rawbtn_Clicked)
        self.br_rawfoldbtn.clicked.connect(self.br_rawfoldbtn_Clicked)

    # @_decorator
    def br_rawbtn_Clicked(self):
        self.plant = str(self.comboBox_2.currentText())
        self.start = int(str(self.start_m_3.date().year()) + '0' * (
                2 - len(str(self.start_m_3.date().month()))) + str(self.start_m_3.date().month()))
        self.end = int(str(self.end_m_3.date().year()) + '0' * (
                2 - len(str(self.end_m_3.date().month()))) + str(self.end_m_3.date().month()))
        try :
            data = IntBR(self.plant, self.start, self.end)
            df = data.df
            if str(self.comboBox_gen_exc.currentText()) == '일반':
                df = df[df['일반/교류'] == 'Regular']
            elif str(self.comboBox_gen_exc.currentText()) == '교류':
                df = df[df['일반/교류'] == 'Exchange']
            if str(self.comboBox_cc.currentText()) == '켐페인제외':
                filter_1 = df['통보서'].apply(lambda x: x[-2]) == 'C'
                df['캠페인'] = np.where(filter_1, 'C', 'W')
                df = df[df['캠페인'] != 'C']
            if str(self.comboBox_int_spe.currentText()) == '통분':
                df = df[df['분담율검증'] == 'Integrated BR']
            elif str(self.comboBox_int_spe.currentText()) == '개별':
                df = df[df['분담율검증'] == 'Special BR']
            if len(df) > 100000:
                self.statusbar.showMessage(f"100,000건을 초과할 수 없습니다. ({len(df)}건)")
            else:
                export(df, f'RAW_{self.plant}_{self.start}-{self.end}', interval=3000)
                self.statusbar.showMessage(f"로우데이터생성 완료. ({len(df)}건)")
                zip_ = r'Cookies\Pocket.zip'
                penguins = open_zipfile(zip_, 'Penguins.png')
                self.image_2.setPixmap(QPixmap(penguins))
        except Exception as e:
            print(e)
            self.statusbar.showMessage(f"파일을 생성할 수 없습니다. ({e})")
            return None

    # @_decorator
    def br_rawfoldbtn_Clicked(self):
        try:
            os.startfile('Spawn')
        except Exception as e:
            print(e)
            self.statusbar.showMessage(f"폴더를 열 수 없습니다. ({e})")
            return None
        self.statusbar.showMessage("로우데이터생성 폴더오픈")

    def intbr_group(self):
        self.start_m_2.setDate(QtCore.QDate.currentDate())
        self.end_m_2.setDate(QtCore.QDate.currentDate())
        self.calculate_btn.clicked.connect(self.calculate_btn_Clicked)
        self.test_btn.clicked.connect(self.test_btn_Clicked)
        self.br_foldbtn.clicked.connect(self.br_foldbtn_Clicked)
        self.br_filebtn.clicked.connect(self.br_filebtn_Clicked)
        self.br_emailbtn.clicked.connect(self.br_emailbtn_Clicked)

    # @_decorator
    def br_emailbtn_Clicked(self):
        if self.mail_disable.isChecked():
            self.statusbar.showMessage('메일 송부 기능이 비활성화 되었습니다.')
            return None
        if not self.attachment:
            self.statusbar.showMessage('생성된 자료가 없습니다.')
        else:
            print(os.getcwd())
            zip_ = r'Cookies\Pocket.zip'
            footer = open_zipfile(zip_, 'mail_footer.png')
            if self.plot:
                img = [self.plot, footer]
            else:
                img = [footer]
            subject = f'{self.attachment}'
            content = f'안녕하세요.<br /><br />{self.attachment} 통합분담율 자료 송부드리오니 확인부탁드립니다.<br /><br />감사합니다.<br /><br />'
            AutoEmail(receiver=self.lineEdit_to.text(), cc=self.lineEdit_cc.text(), subject=subject,
                      content=content, attachments=[os.path.join('Spawn', self.attachment)], img=img).send()
            try :
                self.statusbar.showMessage('메일 송부완료')
            except Exception as e:
                print(e)
                self.statusbar.showMessage(f"메일을 송부할 수 없습니다. ({e})")
                return None

    # @_decorator
    def br_filebtn_Clicked(self):
        if self.attachment:
            try:
                os.startfile(os.path.join('Spawn', self.attachment))
            except Exception as e:
                print(e)
                self.statusbar.showMessage(f"파일을 열 수 없습니다. ({e})")
                return None
            self.statusbar.showMessage("이의제기 생성자료 오픈")
        else :
            self.statusbar.showMessage('생성된 자료가 없습니다.')
            return None

    # @_decorator
    def br_foldbtn_Clicked(self):
        try:
            os.startfile('Spawn')
        except Exception as e:
            print(e)
            self.statusbar.showMessage(f"폴더를 열 수 없습니다. ({e})")
            return None
        self.statusbar.showMessage("통분자료생성 폴더오픈")

    # @_decorator
    def calculate_btn_Clicked(self):
        self.plant = str(self.comboBox.currentText())
        self.start = int(str(self.start_m_2.date().year()) + '0' * (
                2 - len(str(self.start_m_2.date().month()))) + str(self.start_m_2.date().month()))
        self.end = int(str(self.end_m_2.date().year()) + '0' * (
                2 - len(str(self.end_m_2.date().month()))) + str(self.end_m_2.date().month()))
        try :
            self.obj = IntBR(self.plant, self.start, self.end)
            self.optimal = SimpleIntBRSolver(self.obj.df).answer
            self.attachment = self.obj.calculate()
            print(self.attachment)
            self.statusbar.showMessage(f'통합분담율 계산결과 ({self.obj.new_cbr}), 최적값 ({self.optimal})')
        except Exception as e :
            print(e)
            self.statusbar.showMessage(f'{self.start} ~ {self.end} 데이터에 문제가 있습니다. ({e})')
            return None
        try :
            self.lineEdit_exc.setEnabled(True)
            self.lineEdit_gen.setText(str(self.obj.new_cbr['Regular']))
            self.lineEdit_exc.setText(str(self.obj.new_cbr['Exchange']))
        except Exception as e:
            self.lineEdit_exc.setText("")
            self.lineEdit_exc.setEnabled(False)

    # @_decorator
    def test_btn_Clicked(self):
        try :
            self.obj.new_cbr['Regular'] = float(self.lineEdit_gen.text())
            if self.lineEdit_exc.text() != "":
                self.obj.new_cbr['Exchange'] = float(self.lineEdit_exc.text())
        except Exception as e:
            self.statusbar.showMessage(f'테스트 설정값에 문제가 있습니다. ({e})')
            return None
        self.statusbar.showMessage(f'다음 값으로 테스트합니다 ({self.obj.new_cbr}), 최적값 ({self.optimal})')
        try :
            self.obj.plot()
        except Exception as e :
            self.statusbar.showMessage(f'테스트 수행에 실패하였습니다. ({e})')
            return None
        try :
            self.plot = os.path.join('Images', f"{self.plant}_{self.start}-{self.end}_plot.png")
            self.image_2.setPixmap(QPixmap(self.plot))
        except Exception as e :
            self.statusbar.showMessage(f'이미지 출력에 문제가 있습니다. ({e})')
            return None

    def mail_group(self):
        self.mail_disable.stateChanged.connect(self.mail_checkbox_Checked)
        self.lineEdit_to.setText('bj.kim958@glovis.net;kjy5454@glovis.net')
        self.lineEdit_cc.setText('lih@glovis.net')

    def mail_checkbox_Checked(self):
        if self.mail_disable.isChecked():
            self.lineEdit_to.setEnabled(False)
            self.lineEdit_cc.setEnabled(False)
        else:
            self.lineEdit_to.setEnabled(True)
            self.lineEdit_cc.setEnabled(True)

    def cst_group(self):
        self.start_m.setDate(QtCore.QDate.currentDate())
        self.end_m.setDate(QtCore.QDate.currentDate())
        self.end_m.setEnabled(False)
        self.year_checkBX_1.stateChanged.connect(self.cst_checkbox_Checked)
        self.consult_confirmBtn.clicked.connect(self.cst_confirm_btn_Clicked)
        self.consult_folderBtn.clicked.connect(self.cst_consult_folder_btn_Clicked)
        self.email_Btn.clicked.connect(self.cst_email_btn_Clicked)
        self.file_Btn.clicked.connect(self.file_btn_Clicked)
        self.sourceBtn.clicked.connect(self.source_btn_Clicked)

    def cst_checkbox_Checked(self):
        if self.year_checkBX_1.isChecked():
            self.end_m.setEnabled(True)
        else:
            self.end_m.setEnabled(False)

    # @_decorator
    def cst_confirm_btn_Clicked(self):
        self.start = int(str(self.start_m.date().year()) + '0' * (
                2 - len(str(self.start_m.date().month()))) + str(self.start_m.date().month()))
        self.end = int(str(self.end_m.date().year()) + '0' * (
                2 - len(str(self.end_m.date().month()))) + str(self.end_m.date().month()))
        if self.year_checkBX_1.isChecked():
            try :
                MainExecute(start_m=self.start, end_m=self.end).consult_mul()
            except Exception as e:
                print(e)
                self.statusbar.showMessage(f'{self.start} ~ {self.end} 데이터에 문제가 있습니다. ({e})')
                return None
            self.statusbar.showMessage(f'{self.start} ~ {self.end} 품의자료 생성완료')
            self.attachment = f'{self.start}-{self.end}_Consult.xlsx'
        else:
            try :
                MainExecute(start_m=self.start, end_m=None).consult_sin()
            except Exception as e:
                print(e)
                self.statusbar.showMessage(f'{self.start} 데이터에 문제가 있습니다. ({e})')
                return None
            self.attachment = f'{self.start}_Consult.xlsx'
            self.statusbar.showMessage(f'{self.start} 품의자료 생성완료')
        MakeClosingPlots(self.attachment).plot_data()
        self.plot = os.path.join('Images', self.attachment.split('.')[0] + '.png')
        self.image.setPixmap(QPixmap(self.plot))

    # @_decorator
    def cst_consult_folder_btn_Clicked(self):
        try:
            os.startfile('Spawn')
        except Exception as e:
            print(e)
            self.statusbar.showMessage(f"폴더를 열 수 없습니다. ({e})")
            return None
        self.statusbar.showMessage("품의자료생성 폴더오픈")

    # @_decorator
    def cst_email_btn_Clicked(self):
        if self.mail_disable.isChecked():
            self.statusbar.showMessage('메일 송부 기능이 비활성화 되었습니다.')
            return None
        print(self.lineEdit_to.text())
        print(self.lineEdit_cc.text())
        if not self.attachment:
            self.statusbar.showMessage('생성된 자료가 없습니다.')
        else:
            print(os.getcwd())
            zip_ = r'Cookies\Pocket.zip'
            footer = open_zipfile(zip_, 'mail_footer.png')
            if self.plot:
                img = [self.plot, footer]
            else:
                img = [footer]
            subject = f'{self.attachment}'
            content = f'안녕하세요.<br /><br />{self.attachment} 품의/변제율 자료 송부드리오니 확인부탁드립니다.<br /><br />감사합니다.<br /><br />'
            try :
                AutoEmail(receiver=self.lineEdit_to.text(), cc=self.lineEdit_cc.text(), subject=subject,
                          content=content, attachments=[os.path.join('Spawn', self.attachment)], img=img).send()
                self.statusbar.showMessage('메일 송부완료')
            except Exception as e:
                print(e)
                self.statusbar.showMessage(f"메일을 송부할 수 없습니다. ({e})")
                return None

    # @_decorator
    def file_btn_Clicked(self):
        if not self.attachment:
            self.statusbar.showMessage('생성된 자료가 없습니다.')
        else:
            os.startfile(os.path.join('Spawn' ,self.attachment))
            self.statusbar.showMessage('생성자료 오픈')

    # @_decorator
    def source_btn_Clicked(self):
        os.startfile(r'Cookies')
        try :
            self.statusbar.showMessage("업로드폴더 오픈")
        except Exception as e:
            print(e)
            self.statusbar.showMessage(f"폴더를 열 수 없습니다. ({e})")
            return None

    def obj_group(self):
        self.GSCM_BTN.clicked.connect(self.GSCM_btn_Clicked)
        self.obj_folderBtn.clicked.connect(self.execute_folder_btn_Clicked)
        self.excute_BTN_1.clicked.connect(self.execute_btn_Clicked)
        self.obj_fileBtn.clicked.connect(self.obj_file_btn_Clicked)

    # @_decorator
    def execute_btn_Clicked(self):
        # sys.path.insert(1, os.path.join(os.getcwd(), 'Objections'))
        try :
            os.startfile(r'Objections\Entry.exe')
        except Exception as e:
            print(e)
            self.statusbar.showMessage(f"이의제기 실행시 오류가 발생했습니다. ({e})")
            return None
        self.statusbar.showMessage("이의제기 실행")

    # @_decorator
    def execute_folder_btn_Clicked(self):
        try :
            os.startfile(r'Cookies_objection')
        except Exception as e:
            print(e)
            self.statusbar.showMessage(f"폴더를 열 수 없습니다. ({e})")
            return None
        self.statusbar.showMessage("이의제기자료 폴더오픈")

    # @_decorator
    def obj_file_btn_Clicked(self):
        try :
            os.startfile(r'Cookies_objection\objection.xls')
        except Exception as e:
            print(e)
            self.statusbar.showMessage(f"자료를 열 수 없습니다. ({e})")
            return None
        self.statusbar.showMessage("이의제기자료오픈")

    def GSCM_btn_Clicked(self):
        import webbrowser
        self.statusbar.showMessage("GSCM 페이지 오픈")
        webbrowser.open('https://partner.hyundai.com/gscm/', new=2, autoraise=False)

    def db_group(self):
        self.year_checkBX_2.stateChanged.connect(self.db_checkbox_Checked)
        self.DB_start_m.setDate(QtCore.QDate.currentDate())
        self.DB_CRT_BTN.clicked.connect(self.db_crt_btn_Clicked)
        self.DB_OPN_BTN.clicked.connect(self.db_opn_btn_Clicked)
        self.upload_BTN_2.clicked.connect(self.db_upload_btn_2_Clicked)

    def db_checkbox_Checked(self):
        if self.year_checkBX_2.isChecked():
            self.db_start = int(str(self.DB_start_m.date().year()))
        else:
            self.db_start = int(str(self.DB_start_m.date().year()) + '0' * (
                    2 - len(str(self.DB_start_m.date().month()))) + str(self.DB_start_m.date().month()))

    # @_decorator
    def db_crt_btn_Clicked(self):
        self.db_checkbox_Checked()
        if self.year_checkBX_2.isChecked():
            try :
                MainDB.save_year(self.db_start)
            except Exception as e:
                print(e)
                self.statusbar.showMessage(f'데이터베이스를 생성하지 못했습니다. ({e})')
                return None
        else:
            try :
                MainDB(self.db_start).save_db()
            except Exception as e:
                print(e)
                self.statusbar.showMessage(f'데이터베이스를 생성하지 못했습니다. ({e})')
                return None
        self.statusbar.showMessage("데이터베이스 생성완료")
        zip_ = r'Cookies\Pocket.zip'
        penguins = open_zipfile(zip_, 'Penguins.png')
        self.image.setPixmap(QPixmap(penguins))

    # @_decorator
    def db_opn_btn_Clicked(self):
        try :
            os.startfile(r'Main_DB\Main_DB.db')
        except Exception as e:
            print(e)
            self.statusbar.showMessage(f"데이터베이스를 열 수 없습니다. ({e})")
            return None
        self.statusbar.showMessage("데이터베이스 오픈")

    # @_decorator
    def db_upload_btn_2_Clicked(self):
        try :
            os.startfile(r'Cookies')
        except Exception as e:
            print(e)
            self.statusbar.showMessage(f"폴더를 열 수 없습니다. ({e})")
            return None
        self.statusbar.showMessage("업로드폴더 오픈")

    @staticmethod
    def main():
        app = QtWidgets.QApplication(sys.argv)
        MainWindow_ = MyMainWindow()
        MainWindow_.show()
        sys.exit(app.exec_())


if __name__ == '__main__':
    MyMainWindow.main()