from ast import Return
from functools import partial
import re
import subprocess
import sys
import os
from tabnanny import check
import threading
import uuid
import docx2pdf
import shutil
import glob
from PIL import Image
from PyQt5.uic import loadUi
from PyQt5 import QtWidgets, QtCore
from PyQt5.QtWidgets import *
from PyQt5.QtGui import QPixmap, QIcon
from PyQt5.QtCore import *
from docxtpl import DocxTemplate, InlineImage, RichText
from PyPDF2 import PdfFileReader, PdfMerger, PdfReader
from PyQt5.QtMultimedia import QMediaContent, QMediaPlayer
from PyQt5.QtMultimediaWidgets import QVideoWidget
import time
import platform
from cryptography.fernet import Fernet
import mysql.connector


# regex
emailreg = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'
phonereg = r'(?:\+\d{2})?\d{3,4}\D?\d{3}\D?\d{3}'

# Lists of static pdf
PDFs = [
    ["VRD TERRASSEMENT", False, "FICHE VRD TERRASSEMENT.pdf"],
    ["FONDATION SPÉCIALE", False, "FICHE FONDATION SPECIALE.pdf"],
    ["DÉMOLITION CURAGE", False, "FICHE DEMOLITION CURAGE.pdf"],
    ["GROS ŒUVRE", False, "FICHE-GROS-OEUVRE.pdf"],
    ["COUVERTURE ETANCHÉITÉ", False, "FICHE-COUVERTURE-ETANCHEITE.pdf"],
    ["ZINGUERIE DESCENTE EP", False, "FICHE ZINGUERIE DESCENTE EP.pdf"],
    ["CHARPENTE", False, "FICHE-ChARPENTE.pdf"],
    ["REVÊTEMENT FAÇADE & BARDAGE", False, "FICHE-REVETEMENT-FACADE-BARDAGE.pdf"],
    ["MENUISERIES EXTÉRIEURES", False, "FICHE MENUISERIES EXTERIEURES.pdf"],
    ["MENUISERIES INTÉRIEURES", False, "FICHE-MENUISERIES-INTERIEURES.pdf"],
    ["PLÂTRERIE FAUX PLAFONDS ISOLATION", False,
        "FICHE PLATRERIE FAUX PLAFONDS ISOLATION.pdf"],
    ["REVÊTEMENT SOL CARRELAGE", False, "FICHE REVETEMENT SOL LOT CARRELAGE.pdf"],
    ["REVÊTEMENT SOL SOUPLE", False, "FICHE REVETEMENT SOL LOT SOL SOUPLE.pdf"],
    ["ELECTRICITÉ", False, "FICHE ELECTRICITE.pdf"],
    ["CVC", False, "FICHE-13-CVC.pdf"],
    ["PLOMBERIE", False, "FICHE-PLOMBERIE.pdf"],
    ["SERRURERIE", False, "FICHE SERRURERIE.pdf"],
    ["ALARME CAMÉRA DÉTECTION INCENDIE SSI", False,
        "FICHE ALARME CAMERA DETECTION INCENDIE SSI.pdf"],
    ["MIROITERIE", False, "FICHE-MIROITERIE.pdf"],
    ["NETTOYAGE DE CHANTIER", False, "FICHE NETTOYAGE.pdf"],
    ["FIBRE OPTIQUE", False, "FICHE-OPTIQUE.pdf"]
]

# Number of companies
companies = 0
cur_comp = 0
comps = {}

name = ""

appData = {}

cur_plat = platform.system()

# Function to control shown widget


def set_wid(i):
    # i + 1 Becuse first window is not calculated
    holder.setCurrentIndex(i + 1)


# Function to hide the error labels on app launch
def hide_lblerror(w):
    for widget in w.children():
        # error lablels name starts with lbl always
        if (isinstance(widget, QLabel) & (widget.objectName()[:3] == "lbl")):
            widget.setVisible(False)


# Hide some error label
def remove_error(txt, label, border=True):
    if border:
        txt.setStyleSheet(":focus{border: 1px Solid #ff9801;}")
    else:
        txt.setStyleSheet("")
    label.setHidden(True)


def remove_error_m(border=True, *txts):
    i = 0
    while (i < txts.__len__() - 1):
        remove_error(txts[i], txts[i + 1], border)
        i += 2

# Show some error label


def show_error(txt, label):
    txt.setStyleSheet("border: 1px solid red;")
    label.setHidden(False)


# textboxes Validation
def validate_txt(txt, label, email=False, phone=False, image=False, num=False):
    # Check Empty
    if(len(txt.toPlainText()) == 0):
        show_error(txt, label)
        txt.textChanged.connect(
            partial(remove_error, txt, label))
        return False
    else:
        # check email or phone valid
        if(email):
            if not(re.fullmatch(emailreg, txt.toPlainText())):
                show_error(txt, label)
                txt.textChanged.connect(
                    partial(remove_error, txt, label))
                return False
        elif(phone):
            if not(re.fullmatch(phonereg, txt.toPlainText())):
                show_error(txt, label)
                txt.textChanged.connect(
                    partial(remove_error, txt, label))
                return False
        elif(image):
            if not (os.path.exists(txt.toPlainText())):
                show_error(txt, label)
                txt.textChanged.connect(
                    partial(remove_error, txt, label))
                return False
        elif(num):
            try:
                int(txt.toPlainText())
            except:
                show_error(txt, label)
                txt.textChanged.connect(
                    partial(remove_error, txt, label))
                return False

        return True


def min_max(txt_max, txt_min, lbl_max, lbl_min):
    if(isinstance(txt_max, QDateEdit) & isinstance(txt_min, QDateEdit)):
        if(txt_min.date() >= txt_max.date()):
            show_error(txt_min, lbl_min)
            show_error(txt_max, lbl_max)
            txt_min.dateChanged.connect(
                partial(remove_error_m, False, txt_min, lbl_min, txt_max, lbl_max))
            txt_max.dateChanged.connect(
                partial(remove_error_m, False, txt_min, lbl_min, txt_max, lbl_max))
            return False
    elif (isinstance(txt_max, QTimeEdit) & isinstance(txt_min, QTimeEdit)):
        if(txt_min.time() >= txt_max.time()):
            show_error(txt_min, lbl_min)
            show_error(txt_max, lbl_max)
            txt_min.timeChanged.connect(
                partial(remove_error_m, False, txt_min, lbl_min, txt_max, lbl_max))
            txt_max.timeChanged.connect(
                partial(remove_error_m, False, txt_min, lbl_min, txt_max, lbl_max))
            return False
    elif (isinstance(txt_max, QTextEdit) & isinstance(txt_min, QTextEdit)):
        try:
            if(int(txt_min.toPlainText()) >= int(txt_max.toPlainText())):
                show_error(txt_min, lbl_min)
                show_error(txt_max, lbl_max)
                txt_min.textChanged.connect(
                    partial(remove_error_m, True, txt_min, lbl_min, txt_max, lbl_max))
                txt_max.textChanged.connect(
                    partial(remove_error_m, True, txt_min, lbl_min, txt_max, lbl_max))
                return False
        except:
            return True

    return True


def restart():
    for i in range(0, 13):
        holder.removeWidget(holder.widget(0))

    frms()
    set_wid(0)


class Frm_li(QDialog):
    # Load UI
    def __init__(self):
        super(Frm_li, self).__init__()
        loadUi("UI/frm_li.ui", self)
        hide_lblerror(self.pnl_main)
        self.btn_next.clicked.connect(self.check_l)

    def check_l(self):
        if not (validate_txt(self.txt_li, self.lbl_li)):
            return
        try:
            mydb = mysql.connector.connect(
                host="bm1cehzgazdmczql5toa-mysql.services.clever-cloud.com",
                user="ubucfxpk856ntdmg",
                password="s0U5zDcArkKxKVl1xGBh",
                database="bm1cehzgazdmczql5toa")
            mycursor = mydb.cursor()
            mycursor.execute("SELECT * FROM Users WHERE serial=%s AND DATE(en_date) >= CURDATE() ", [
                             str(self.txt_li.toPlainText()), ])
            myresult = mycursor.fetchall()
            if(myresult.__len__() <= 0):
                mycursor.close()
                mydb.close()
                msg = QMessageBox()
                msg.setIcon(QMessageBox.Critical)
                msg.setText("لا يوجد سريال بهذه المواصفات")
                msg.setWindowTitle("فشل")
                retval = msg.exec_()
            else:
                mb = str(uuid.UUID(int=uuid.getnode()))
                mycursor.execute("UPDATE Users SET board=%s WHERE serial=%s", [
                                 mb, str(self.txt_li.toPlainText())])
                mydb.commit()
                mycursor.close()
                mydb.close()
                check_license()
                msg = QMessageBox()
                msg.setIcon(QMessageBox.Icon.Information)
                msg.setText("قم بإعادة تشغيل البرنامج الأن")
                msg.setWindowTitle("نجح")
                retval = msg.exec_()
                self.close()

        except NameError:
            print(NameError)
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Critical)
            msg.setText("لا يوجد اتصال بالانترنت")
            msg.setWindowTitle("فشل")
            retval = msg.exec_()
            sys.exit()


def open_li():
    app3 = QApplication(sys.argv)
    app3.setApplicationName("PPSPS Version 1.0")
    frm_li = Frm_li()
    frm_li.show()
    app3.exec_()


def check_license():
    if not (os.path.exists("key")):
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Critical)
        msg.setText("هناك ملفات ناقصة من البرنامج")
        msg.setWindowTitle("فشل")
        retval = msg.exec_()
        sys.exit()

    mb = uuid.UUID(int=uuid.getnode())
    # Check DB Connection
    try:
        mydb = mysql.connector.connect(
            host="bm1cehzgazdmczql5toa-mysql.services.clever-cloud.com",
            user="ubucfxpk856ntdmg",
            password="s0U5zDcArkKxKVl1xGBh",
            database="bm1cehzgazdmczql5toa")
        mycursor = mydb.cursor()
        mycursor.execute(
            "SELECT * FROM Users WHERE board=%s AND DATE(en_date) >= CURDATE() ", [str(mb), ])
        myresult = mycursor.fetchall()
        mycursor.close()
        mydb.close()
        if(myresult.__len__() > 0):
            # Verified
            li = open("li.txt", "w")
            key = open("key", "r")
            k = key.readlines()
            li_en = Fernet(str.encode(str(k))).encrypt(
                str.encode(myresult[0][0]))
            mb_en = Fernet(str.encode(str(k))).encrypt(str.encode(str(mb)))
            li.writelines([li_en.decode(), '\n', mb_en.decode()])
            key.close()
            li.close()
            return True
            # open app
        else:
            if(os.path.exists("li.txt")):
                os.remove("li.txt")
            open_li()
            return False

    except:
        print("Failed to connect to database")
        if not (os.path.exists("li.txt")):
            open_li()
            return False
        li = open("li.txt", "r")
        key = open("key", "r")
        k = key.readlines()
        l = li.readlines()

        li_de = Fernet(str.encode(str(k))).decrypt(
            str.encode(l[0][:l[0].__len__()])).decode()
        mb_de = Fernet(str.encode(str(k))).decrypt(str.encode(l[1])).decode()

        if not(mb_de == str(mb)):
            open_li()
            return False
        return True


if not (check_license()):
    sys.exit()
# ------------

# ---------------------------------------------------- INTRO


class Frm_intro(QDialog):
    # Load UI
    def __init__(self):
        super(Frm_intro, self).__init__()
        loadUi("UI/frm_intro.ui", self)
        self.mediaPlayer = QMediaPlayer(None, QMediaPlayer.VideoSurface)
        videoWidget = QVideoWidget()
        self.mediaPlayer.setVideoOutput(videoWidget)
        # self.mediaPlayer.setMedia(QMediaContent(QUrl.fromLocalFile(fileName)))
        self.lay.addWidget(videoWidget)
        vidpath = os.getcwd() + "/ppsps.mp4"
        self.mediaPlayer.setMedia(QMediaContent(
            QUrl.fromLocalFile(os.path.relpath(vidpath))))
        self.setWindowFlag(QtCore.Qt.FramelessWindowHint)

    def play(self):
        self.mediaPlayer.play()


def close_intro():
    time.sleep(7)
    intro.close()

# ----------------------------------------------------


class Frm0(QDialog):
    # Load UI
    def __init__(self):
        super(Frm0, self).__init__()
        loadUi("UI/frm0.ui", self)
        self.label.setPixmap(QPixmap(u"UI/background.jpg"))
        self.btn_start.clicked.connect(partial(set_wid, 0))
        self.btn_exit.clicked.connect(app.exit)


class Frm1(QDialog):
    # Load UI
    def __init__(self):
        super(Frm1, self).__init__()
        loadUi("UI/frm1.ui", self)
        self.label.setPixmap(QPixmap(u"UI/background.jpg"))
        hide_lblerror(self.pnl_main)
        self.btn_choose.clicked.connect(self.chooseimg)
        self.btn_next.clicked.connect(self.next)

    # Choose Image(Logo) button event
    def chooseimg(self):
        path = QFileDialog.getOpenFileName(self, 'Choose Logo', '',
                                           'Image files (*.png *.jpg *.gif)')
        if path != ('', ''):
            self.txt_logo.setPlainText(path[0])

    # Next button event
    def next(self):
        # Check empty
        # TODO: Add new Texts
        if not(
            validate_txt(self.txt_name1, self.lbl_name1) &
            validate_txt(self.txt_email, self.lbl_email, True) &
            validate_txt(self.txt_address1, self.lbl_address1) &
            validate_txt(self.txt_qul, self.lbl_qul) &
            validate_txt(self.txt_res, self.lbl_res) &
            validate_txt(self.txt_phone, self.lbl_phone, False, True) &
            validate_txt(self.txt_logo, self.lbl_logo, False, False, True) &
            validate_txt(self.txt_name2, self.lbl_name2) &
            validate_txt(self.txt_address2, self.lbl_address2) &
            validate_txt(self.txt_city, self.lbl_city)
        ):
            return

        # Copy Data..
        appData["frm1"] = {
            "name1": self.txt_name1.toPlainText(),
            "email": self.txt_email.toPlainText(),
            "phone": self.txt_phone.toPlainText(),
            "address1": self.txt_address1.toPlainText(),
            "qul": self.txt_qul.toPlainText(),
            "res": self.txt_res.toPlainText(),
            "logo": self.txt_logo.toPlainText(),
            "name2": self.txt_name2.toPlainText(),
            "address2": self.txt_address2.toPlainText(),
            "city": self.txt_city.toPlainText(),
        }

        set_wid(1)

# ----------------------------------------------------


class Frm2(QDialog):
    def __init__(self):
        super(Frm2, self).__init__()
        loadUi("UI/frm2.ui", self)
        self.label.setPixmap(QPixmap(u"UI/background.jpg"))
        hide_lblerror(self.pnl_main)
        self.btn_back.clicked.connect(self.back)
        self.btn_next.clicked.connect(self.next)

    # Back button event
    def back(self):
        set_wid(0)

    # next button event
    def next(self):
        # Check empty
        if not(validate_txt(self.txt_des, self.lbl_des)):
            return

        # Copy Data..
        pdfs = ""
        for widget in self.chs.children():
            if (isinstance(widget, QCheckBox)):
                if(widget.isChecked()):
                    i = int(widget.objectName()[2:]) - 1
                    pdfs += "  •  " + PDFs[i][0] + "\n"
                    PDFs[i][1] = True

        appData["frm2"] = {"des": self.txt_des.toPlainText(), "PDFs": pdfs}

        set_wid(2)

# ----------------------------------------------------


class Frm3(QDialog):
    def __init__(self):
        super(Frm3, self).__init__()
        loadUi("UI/frm3.ui", self)
        self.label.setPixmap(QPixmap(u"UI/background.jpg"))
        hide_lblerror(self.pnl_main)
        self.btn_back.clicked.connect(self.back)
        self.btn_next.clicked.connect(self.next)

    # Back button event
    def back(self):
        set_wid(1)

    def next(self):
        # Check empty
        if not(
            validate_txt(self.txt_name, self.lbl_name) &
            validate_txt(self.txt_tel, self.lbl_tel, False, True)
        ):
            return

        # Copy Data..
        appData["frm3"] = {
            "name": self.txt_name.toPlainText(),
            "tel": self.txt_tel.toPlainText(),
            "date1": self.date1.date().toString('d MMMM, yyyy')
        }

        set_wid(3)

# ----------------------------------------------------


class Frm4(QDialog):
    def __init__(self):
        super(Frm4, self).__init__()
        loadUi("UI/frm4.ui", self)
        self.label.setPixmap(QPixmap(u"UI/background.jpg"))
        # Add items to combobox
        for i in PDFs:
            self.cb1.addItem(i[0])
        # Update Default dates to now
        self.date1.setDateTime(QtCore.QDateTime.currentDateTime())
        self.date2.setDateTime(QtCore.QDateTime.currentDateTime())
        hide_lblerror(self.pnl_main)
        self.btn_next.clicked.connect(self.next)
        self.btn_back.clicked.connect(self.back)

    # Next button event
    def next(self):
        # check empty txts
        if not(validate_txt(self.txt_period, self.lbl_period) &
               validate_txt(self.txt_wf1, self.lbl_wf1, num=True) &
               validate_txt(self.txt_wf2, self.lbl_wf2, num=True)):
            return

        # check Min Max and date
        if not(min_max(self.date2, self.date1, self.lbl_date, self.lbl_date) &
               min_max(self.txt_wf2, self.txt_wf1, self.lbl_wf2, self.lbl_wf1)):
            return

        # Copy Data..
        appData["frm4"] = {
            "cb1": self.cb1.currentText(),
            "ch1": self.ch1.isChecked(),
            "date1": self.date1.date().toString('d MMMM, yyyy'),
            "date2": self.date2.date().toString('d MMMM, yyyy'),
            "min": int(self.txt_wf1.toPlainText()),
            "max": int(self.txt_wf2.toPlainText()),
            "period": self.txt_period.toPlainText()
        }
        set_wid(4)

    # Back button event
    def back(self):
        set_wid(2)

# ---------------------------------------------------


class Frm5(QDialog):
    def __init__(self):
        super(Frm5, self).__init__()
        loadUi("UI/frm5.ui", self)
        self.label.setPixmap(QPixmap(u"UI/background.jpg"))
        # Update Default dates to now
        hide_lblerror(self.pnl_main)
        self.btn_next.clicked.connect(self.next)
        self.btn_back.clicked.connect(self.back)

    # Next button event
    def next(self):
        # check empty txts
        if not(
            validate_txt(self.txt_num, self.lbl_num, num=True) &
            validate_txt(self.txt_supply, self.lbl_supply)
        ):
            return

        # Check Min Max
        if not(min_max(self.time1_2, self.time1_1, self.lbl_time, self.lbl_time) &
           min_max(self.time2_2, self.time2_1, self.lbl_time, self.lbl_time)):
            return

        # Copy Data..
        appData["frm5"] = {
            "num": int(self.txt_num.toPlainText()),
            "time1_1": self.time1_1.time().toString(),
            "time1_2": self.time1_2.time().toString(),
            "time2_1": self.time2_1.time().toString(),
            "time2_2": self.time2_2.time().toString(),
            "supply": self.txt_supply.toPlainText()
        }
        set_wid(5)

    # Back button event
    def back(self):
        set_wid(3)

# -----------------------------------------------------


class Frm6(QDialog):
    def __init__(self):
        super(Frm6, self).__init__()
        loadUi("UI/frm6.ui", self)
        self.label.setPixmap(QPixmap(u"UI/background.jpg"))
        hide_lblerror(self.pnl_main)
        self.btn_next.clicked.connect(self.next)
        self.btn_back.clicked.connect(self.back)

    # Next button event
    def next(self):
        # check empty txts
        if not(validate_txt(self.txt_num, self.lbl_num, num=True)):
            return

        # Set Number of companies
        global companies
        companies = int(self.txt_num.toPlainText())

        appData["frm6"] = {"Num": companies}

        if (companies > 0):
            set_wid(6)
        else:
            set_wid(7)

    # Back button event
    def back(self):
        print(companies)
        set_wid(4)

# ---------------------------------------------------


class Frm7(QDialog):
    def __init__(self):
        super(Frm7, self).__init__()
        loadUi("UI/frm7.ui", self)
        self.label.setPixmap(QPixmap(u"UI/background.jpg"))
        # Update Default dates to now
        hide_lblerror(self.pnl_main)
        self.btn_next.clicked.connect(self.next)
        self.btn_back.clicked.connect(self.back)

    # Next button event
    def next(self):
        # check empty txts
        if not(
            validate_txt(self.txt_name, self.lbl_name) &
            validate_txt(self.txt_address, self.lbl_address) &
            validate_txt(self.txt_des, self.lbl_des)
        ):
            return

        # Add data to array
        global cur_comp
        comps[cur_comp] = {"name": self.txt_name.toPlainText(),
                           "address": self.txt_address.toPlainText(),
                           "des": self.txt_des.toPlainText()}

        # Control Multiple Entery
        if(cur_comp + 1 == companies):
            if cur_plat == "Windows":
                rt = RichText()
                for i in comps:
                    # Name
                    rt.add("Nom : ", color="#808080",
                           font="Arial (Body CS)", size=32)
                    rt.add(comps[i]["name"] + "\n", color="#000000",
                           font="Arial (Body CS)", size=32, bold=True)
                    # Address
                    rt.add("Adresse : ", color="#808080",
                           font="Arial (Body CS)", size=32)
                    rt.add(comps[i]["address"] + "\n", color="#000000",
                           font="Arial (Body CS)", size=32, bold=True)
                    # Des
                    rt.add("Nature des travaux sous-traités : ", color="#808080",
                           font="Arial (Body CS)", size=32)
                    rt.add(comps[i]["des"] + "\n\n", color="#000000",
                           font="Arial (Body CS)", size=32, bold=True)
                appData["frm7"] = {"comps": rt}
            else:
                rt = ""
                for i in comps:
                    rt += "Nom : " + comps[i]["name"] + "\n"
                    rt += "Adresse : " + comps[i]["address"] + "\n"
                    rt += "Nature des travaux sous-traités : " + \
                        comps[i]["des"] + "\n"
                appData["frm7"] = {"comps": rt}

            set_wid(7)
        else:
            cur_comp += 1
            self.label_num.setText(f"Enterprise (  {str(cur_comp + 1)}  )")
            self.txt_name.setPlainText("")
            self.txt_address.setPlainText("")
            self.txt_des.setPlainText("")

    # Back button event
    def back(self):
        # Control Multiple Entery
        global cur_comp
        if(cur_comp > 0):
            cur_comp -= 1
            self.label_num.setText(str(cur_comp + 1))
            self.txt_name.setPlainText(comps[cur_comp]["name"])
            self.txt_address.setPlainText(comps[cur_comp]["address"])
            self.txt_des.setPlainText(comps[cur_comp]["des"])
        else:
            set_wid(5)


# -----------------------------------------------------


class Frm8(QDialog):
    def __init__(self):
        super(Frm8, self).__init__()
        loadUi("UI/frm8.ui", self)
        self.label.setPixmap(QPixmap(u"UI/background.jpg"))
        # Update Default dates to now
        hide_lblerror(self.pnl_main)
        self.btn_next.clicked.connect(self.next)
        self.btn_back.clicked.connect(self.back)

    # Next button event
    def next(self):
        # check empty txts
        if not(
            validate_txt(self.txt_tel1, self.lbl_tel1, phone=True) &
            validate_txt(self.txt_tel2, self.lbl_tel2, phone=True) &
            validate_txt(self.txt_tel3, self.lbl_tel3, phone=True) &
            validate_txt(self.txt_name1, self.lbl_name1) &
            validate_txt(self.txt_name2, self.lbl_name2) &
            validate_txt(self.txt_name3, self.lbl_name3) &
            validate_txt(self.txt_address1, self.lbl_address1) &
            validate_txt(self.txt_address2, self.lbl_address2) &
            validate_txt(self.txt_address3, self.lbl_address3)
        ):
            return

        # Copy Data
        appData["frm8"] = {
            "name1": self.txt_name1.toPlainText(),
            "tel1": self.txt_tel1.toPlainText(),
            "address1": self.txt_address1.toPlainText(),
            "name2": self.txt_name2.toPlainText(),
            "tel2": self.txt_tel2.toPlainText(),
            "address2": self.txt_address2.toPlainText(),
            "name3": self.txt_name3.toPlainText(),
            "tel3": self.txt_tel3.toPlainText(),
            "address3": self.txt_address3.toPlainText(),
            "ch1": self.ch1.isChecked()
        }
        set_wid(8)

    # Back button event
    def back(self):
        if(companies > 0):
            set_wid(6)
        else:
            set_wid(5)


# -----------------------------------------------------


class Frm9(QDialog):
    def __init__(self):
        super(Frm9, self).__init__()
        loadUi("UI/frm9.ui", self)
        self.label.setPixmap(QPixmap(u"UI/background.jpg"))
        # Update Default dates to now
        hide_lblerror(self.pnl_main)
        self.btn_next.clicked.connect(self.next)
        self.btn_back.clicked.connect(self.back)

    # Next button event
    def next(self):
        # check empty txts
        if not(
            validate_txt(self.txt_1, self.lbl_1) &
            validate_txt(self.txt_2, self.lbl_2) &
            validate_txt(self.txt_3, self.lbl_3) &
            validate_txt(self.txt_4, self.lbl_4)
        ):
            return

        # Copy Data
        appData["frm9"] = {
            "_1": self.txt_1.toPlainText(),
            "_2": self.txt_2.toPlainText(),
            "_3": self.txt_3.toPlainText(),
            "_4": self.txt_4.toPlainText()
        }

        set_wid(9)

    # Back button event
    def back(self):
        set_wid(7)


# -----------------------------------------------------


class Frm10(QDialog):
    def __init__(self):
        super(Frm10, self).__init__()
        loadUi("UI/frm10.ui", self)
        self.label.setPixmap(QPixmap(u"UI/background.jpg"))
        # Update Default dates to now
        hide_lblerror(self.pnl_main)
        self.txt_other.setVisible(False)
        self.btn_next.clicked.connect(self.next)
        self.btn_back.clicked.connect(self.back)
        self.ch5.stateChanged.connect(self.txt_others)

    # Next button event
    def next(self):
        # check empty txts
        if not (validate_txt(self.txt1, self.lbl_1) &
           validate_txt(self.txt2, self.lbl_2) &
           validate_txt(self.txt_h1, self.lbl_h) &
           validate_txt(self.txt_h2, self.lbl_h)):
            return

        appData["frm10"] = {
            "txt1": self.txt1.toPlainText(),
            "txt2": self.txt1.toPlainText(),
            "h1": self.txt_h1.toPlainText(),
            "h2": self.txt_h2.toPlainText(),
            "park": self.ch_park.isChecked(),
            "plan": self.ch_plan.isChecked()
        }
        set_wid(10)
        t = threading.Thread(target=holder.widget(11).create_PPSPS)
        t.start()

    # Back button event
    def back(self):
        set_wid(8)

    # Show-Hide txt Others
    def txt_others(self):
        if(self.ch5.isChecked()):
            self.txt_other.setVisible(True)
        else:
            self.txt_other.setVisible(False)

# -----------------------------------------------------


class Frm_wait(QDialog):
    def __init__(self):
        super(Frm_wait, self).__init__()
        loadUi("UI/frm_wait.ui", self)
        self.label.setPixmap(QPixmap(u"UI/background.jpg"))
        # Update Default dates to now

    def update_progress(self, step):
        self.lbl_progress.setText(str(round(step/8 * 100)) + "%")
        sizePolicy = QtWidgets.QSizePolicy(
            QtWidgets.QSizePolicy.Policy.Preferred, QtWidgets.QSizePolicy.Policy.Preferred)
        sizePolicy.setHorizontalStretch(step)
        sizePolicy2 = QtWidgets.QSizePolicy(
            QtWidgets.QSizePolicy.Policy.Preferred, QtWidgets.QSizePolicy.Policy.Preferred)
        sizePolicy2.setHorizontalStretch(8 - step)
        self.w_done.setSizePolicy(sizePolicy)
        self.w_wait.setSizePolicy(sizePolicy2)

    def create_PPSPS(self):

        dic = {}

        # Add Normal Date..
        self.update_progress(1)
        for frm in appData:
            for data in appData[frm]:
                s = frm + data
                if(isinstance(appData[frm][data], bool)):
                    if(appData[frm][data]):
                        dic[s] = "Oui"
                    else:
                        dic[s] = "Non"
                else:
                    dic[s] = appData[frm][data]

        doc = DocxTemplate("Base.docx")

        self.update_progress(2)
        # Image Logo
        img = Image.open(appData["frm1"]["logo"])
        # Calculate new size
        h = float(img.height)
        w = float(img.width)
        mh = 185.0
        mw = 251.0
        while (h > mh or w > mw):
            if(h > mh):
                tmph = h - mh
                tmph = (tmph / h)
                h = mh
                w = w - (w * tmph)
            elif(w > mw):
                tmpw = w - mw
                tmpw = (tmpw / w)
                w = mw
                h = h - h * tmpw
        # Create temp folder
        if not os.path.exists("tmpfiles/"):
            os.makedirs("tmpfiles/")
        # resize
        img.resize((int(w), int(h))).save("tmpfiles/logo.png")
        dic["logo"] = InlineImage(doc, "tmpfiles/logo.png")
        self.update_progress(3)
        # Render Values
        doc.render(dic)
        self.update_progress(4)
        # Save doc
        doc.save("tmpfiles/tmp.docx")
        self.update_progress(5)
        # -- Convert to pdf
        # Check name
        global name
        name = dic["frm1name1"]
        i = 1
        if not os.path.exists("backup/"):
            os.makedirs("backup/")
        while(os.path.exists(f"backup/{name}/")):
            name += "(" + str(i) + ")"
        os.makedirs(f"backup/{name}/SPLITED/")
        # Convert Docx to PDF
        if cur_plat == "Windows":
            try:
                docx2pdf.convert(f'./tmpfiles/tmp.docx',
                                 f'./backup/{name}/SPLITED/{name}.pdf')
            except:
                msg = QMessageBox()
                msg.setIcon(QMessageBox.Critical)
                msg.setText("تأكد من تثبيت برنامج microsoft word")
                msg.setWindowTitle("فشل")
                retval = msg.exec_()
        elif cur_plat == "Linux":
            try:
                subprocess.check_output(
                    ['libreoffice', '--convert-to', 'pdf', f'./tmpfiles/tmp.docx'])
                shutil.move(f'tmp.pdf', f'./backup/{name}/SPLITED/{name}.pdf')
            except:
                msg = QMessageBox()
                msg.setIcon(QMessageBox.Critical)
                msg.setText("تأكد من تثبيت برنامج LIBRE_OFFICE")
                msg.setWindowTitle("فشل")
                retval = msg.exec_()
        elif cur_plat == "Darwin":
            try:
                subprocess.check_output(
                    ['libreoffice', '--convert-to', 'pdf', f'./tmpfiles/tmp.docx'])
                shutil.move(f'tmp.pdf', f'./backup/{name}/SPLITED/{name}.pdf')
            except:
                docx2pdf.convert(f'./tmpfiles/tmp.docx',
                                 f'./backup/{name}/SPLITED/{name}.pdf')
        self.update_progress(6)
        shutil.rmtree('./tmpfiles')

        set_wid(11)

        self.update_progress(7)
        # copy static pdf
        for i in PDFs:
            if i[1]:
                shutil.copyfile(
                    f"staticPDFs/{i[2]}", f"backup/{name}/SPLITED/{i[0]}.pdf")

        self.update_progress(8)
        # Merge all pdfs together
        merger = PdfMerger()
        merger.append(f'./backup/{name}/SPLITED/{name}.pdf')
        for file in glob.glob(f'./backup/{name}/SPLITED/*.pdf'):
            if file == f'./backup/{name}/SPLITED\{name}.pdf':
                continue
            merger.append(file)
        merger.write(f"./backup/{name}/{name}.pdf")
        merger.close()
        return

# -----------------------------------------------------


class Frm_end(QDialog):
    def __init__(self):
        super(Frm_end, self).__init__()
        loadUi("UI/frm_end.ui", self)
        self.label.setPixmap(QPixmap(u"./Ui/background.jpg"))
        self.btn_save.clicked.connect(self.save_click)
        self.btn_print.clicked.connect(self.print_click)
        self.btn_exit.clicked.connect(self.exit_click)
        self.btn_restart.clicked.connect(self.restart_click)

    def save_click(self):
        path = str(QFileDialog.getExistingDirectory(self, "Select Directory"))
        msg = QMessageBox()
        try:
            shutil.copytree(
                os.getcwd() + f"/backup/{name}/", path, dirs_exist_ok=True)
            msg.setIcon(QMessageBox.Information)
            msg.setText("Enregistré")
            msg.setWindowTitle("Succès")
        except:
            msg.setIcon(QMessageBox.Critical)
            msg.setText("Échec de l'enregistrement du PDF")
            msg.setWindowTitle("Échouer")
        retval = msg.exec_()

    def print_click(self):
        if cur_plat == "Linux":
            subprocess.call(
                ['xdg-open', os.path.abspath(f"backup/{name}/{name}.pdf")])
        elif cur_plat == "Windows":
            os.startfile(os.path.abspath(f"backup/{name}/{name}.pdf"))
        else:
            subprocess.call('open', os.path.abspath(
                f"backup/{name}/{name}.pdf"))

    def exit_click(self):
        app.exit()

    def restart_click(self):
        restart()


def frms():

    # Form 0
    frm0 = Frm0()
    holder.addWidget(frm0)
    # Form 1
    frm1 = Frm1()
    holder.addWidget(frm1)
    # Form 2
    frm2 = Frm2()
    holder.addWidget(frm2)
    # Form 3
    frm3 = Frm3()
    holder.addWidget(frm3)
    # Form 4
    frm4 = Frm4()
    holder.addWidget(frm4)
    # Form 5
    frm5 = Frm5()
    holder.addWidget(frm5)
    # Form 6
    frm6 = Frm6()
    holder.addWidget(frm6)
    # Form 7
    frm7 = Frm7()
    holder.addWidget(frm7)
    # Form 8
    frm8 = Frm8()
    holder.addWidget(frm8)
    # Form 9
    frm9 = Frm9()
    holder.addWidget(frm9)
    # Form 10
    frm10 = Frm10()
    holder.addWidget(frm10)
    # Form wait
    frm_wait = Frm_wait()
    holder.addWidget(frm_wait)
    # Form wait
    frm_end = Frm_end()
    holder.addWidget(frm_end)


# -------------------------------------------
# main

app2 = QApplication(sys.argv)
intro = Frm_intro()
intro.resize(1000, 650)
intro.show()
t = threading.Thread(target=intro.play)
t.start()
t2 = threading.Thread(target=close_intro)
t2.start()
app2.exec_()


app = QApplication(sys.argv)
app.setApplicationName("PPSPS Version 1.0")
# Set Icons
app.setWindowIcon(QIcon("./icons/icon (1).ico"))
# Forms Holder
holder = QStackedWidget()

holder.resize(1082, 720)

frms()

holder.show()

sys.exit(app.exec())
