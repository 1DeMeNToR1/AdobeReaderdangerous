import sys
from PyQt5 import QtCore, QtGui, QtWidgets, Qt, uic
from PyQt5.QtWidgets import *
import traceback
import pandas as pd
import time

def log_uncaught_exceptions(ex_cls, ex, tb): #error catcher
    text = '{}: {}:\n'.format(ex_cls.__name__, ex)
    text += ''.join(traceback.format_tb(tb))
    print(text)
    QtWidgets.QMessageBox.critical(None, 'Error', text)
    sys.exit()
sys.excepthook = log_uncaught_exceptions

lines = []


diction = {}

Form, _ = uic.loadUiType('main.ui')
filter, _ = uic.loadUiType('filter.ui')
load, _ = uic.loadUiType('load.ui')


class win_filter(QMainWindow, filter):
    def __init__(self, parent=None):
        super().__init__(parent, QtCore.Qt.Window)
        self.parent = parent
        self.setupUi(self)


class main_ui(QMainWindow, Form):
    def __init__(self):
        super(main_ui, self).__init__()
        self.setupUi(self)
        self.initUI()


    def initUI(self):
        self.build_diagr.clicked.connect(lambda: self.open_filt())
        self.load_file.clicked.connect(lambda: self.load_file_fun())

    def open_filt(self):
        ui2 = win_filter(self)
        ui2.show()

    def load_file_fun(self):
        col = [4, 7, 8, 9, 12, 14, 15, 16, 21]

        file, __ = QFileDialog.getOpenFileName(self, 'Open file')
        start_time = time.time()

        top_players = pd.read_excel(file, usecols=col)
        counter = -1
        for name in top_players['Unnamed: 4'].tolist():
            counter += 1
            if "Adobe Reader" in str(name):
                print(counter)
                lines.append(counter)
                diction[counter] = "Adobe Reader"
        for i in lines:
            programm = top_players['Unnamed: 4'].tolist()[i]
            os = top_players['Unnamed: 7'].tolist()[i]
            class_yazvim = top_players['Unnamed: 8'].tolist()[i]
            date =  top_players['Unnamed: 9'].tolist()[i]
            yr_opas =  top_players['Unnamed: 12'].tolist()[i]
            status_yaz =  top_players['Unnamed: 14'].tolist()[i]
            nal_exploit =  top_players['Unnamed: 15'].tolist()[i]
            ystr_yaz =  top_players['Unnamed: 16'].tolist()[i]
            type_error =  top_players['Unnamed: 21'].tolist()[i]
            diction[i] = (programm, os , class_yazvim, date, yr_opas, status_yaz, nal_exploit, ystr_yaz, type_error)

        krit_yaz = 0
        high_yaz = 0
        med_yaz = 0
        low_yaz = 0

        krit_yaz_ystr = 0
        high_yaz_ystr = 0
        med_yaz_ystr = 0
        low_yaz_ystr = 0
        none1 = 0
        none2 = 0
        none3 = 0
        none4 = 0

        krit_yaz_podtv = 0
        high_yaz_podtv = 0
        med_yaz_podtv = 0
        low_yaz_podtv = 0
        krit_yaz_potenc = 0
        high_yaz_potenc = 0
        med_yaz_potenc = 0
        low_yaz_potenc = 0

        class_yazvim_arch = 0
        class_yazvim_code = 0
        class_yazvim_many = 0
        class_yazvim_none = 0

        y2014A = 0
        y2014Y = 0
        y2015A = 0
        y2015Y = 0
        y2016A = 0
        y2016Y = 0
        y2017A = 0
        y2017Y = 0
        y2018A = 0
        y2018Y = 0
        y2019A = 0
        y2019Y = 0
        y2020A = 0
        y2020Y = 0
        y2021A = 0
        y2021Y = 0


        codes_to_yr_opas = {"Критический уровень опасности ":krit_yaz, "Высокий уровень опасности ":high_yaz, "Средний уровень опасности ":med_yaz, "Низкий уровень опасности ":low_yaz}
        codes_to_ystr_yaz = {"Уязвимость устранена":{"Критический уровень опасности ":krit_yaz_ystr, "Высокий уровень опасности ":high_yaz_ystr, "Средний уровень опасности ":med_yaz_ystr, "Низкий уровень опасности ":low_yaz_ystr},
                             'Информация об устранении отсутствует':{"Критический уровень опасности ":none1, "Высокий уровень опасности ":none2, "Средний уровень опасности ":none3, "Низкий уровень опасности ":none4}}
        codes_to_status_yaz = {"Подтверждена":{"Критический уровень опасности ":krit_yaz_podtv, "Высокий уровень опасности ":high_yaz_podtv, "Средний уровень опасности ":med_yaz_podtv, "Низкий уровень опасности ":low_yaz_podtv},
                               "Потенциальная":{"Критический уровень опасности ":krit_yaz_potenc, "Высокий уровень опасности ":high_yaz_potenc, "Средний уровень опасности ":med_yaz_potenc, "Низкий уровень опасности ":low_yaz_potenc}}
        codes_to_class_yazvim = {"Уязвимость архитектуры":class_yazvim_arch, "Уязвимость кода":class_yazvim_code, "Уязвимость многофакторная":class_yazvim_many, "nan":class_yazvim_none}
        code_to_date_ystr = {"2005":none4,"2014":{"Уязвимость устранена":y2014Y, 'Информация об устранении отсутствует':none4}, "2015":{"Уязвимость устранена":y2015Y, 'Информация об устранении отсутствует':none4}, "2016":{"Уязвимость устранена":y2016Y, 'Информация об устранении отсутствует':none4},"2017":{"Уязвимость устранена":y2017Y, 'Информация об устранении отсутствует':none4}, "2018":{"Уязвимость устранена":y2018Y, 'Информация об устранении отсутствует':none4},
                             "2019":{"Уязвимость устранена":y2019Y, 'Информация об устранении отсутствует':none4},"2020":{"Уязвимость устранена":y2020Y, 'Информация об устранении отсутствует':none4}, "2021":{"Уязвимость устранена":y2021Y, 'Информация об устранении отсутствует':none4}}
        code_to_date_obn = {"2014":y2014Y, "2015":y2015Y,"2016":y2016Y,"2017":y2017Y, "2018":y2018Y,"2019":y2019Y, "2020":y2020Y,"2021":y2021Y}

        for i in diction:
            codes_to_yr_opas[diction[i][4].split("(")[0]] += 1
            codes_to_ystr_yaz[diction[i][7].split("(")[0]][diction[i][4].split("(")[0]] += 1
            codes_to_status_yaz[diction[i][5].split(" ")[0]][diction[i][4].split("(")[0]] += 1
            codes_to_class_yazvim[str(diction[i][2])] += 1
            print(str(diction[i]))
            code_to_date_ystr[str(diction[i][3].split(".")[2])][str(diction[i][7].split("(")[0])] += 1
            code_to_date_obn[str(diction[i][3].split(".")[2])] += 1


        self.krit_obn_labl.setText(str(codes_to_yr_opas["Критический уровень опасности "]))
        self.hig_obn_labl.setText(str(codes_to_yr_opas["Высокий уровень опасности "]))
        self.med_obn_labl.setText(str(codes_to_yr_opas["Средний уровень опасности "]))
        self.low_obn_labl.setText(str(codes_to_yr_opas["Низкий уровень опасности "]))

        self.krit_ystr_labl.setText(str(codes_to_ystr_yaz["Уязвимость устранена"]["Критический уровень опасности "]))
        self.hig_ystr_labl.setText(str(codes_to_ystr_yaz["Уязвимость устранена"]["Высокий уровень опасности "]))
        self.med_ystr_labl.setText(str(codes_to_ystr_yaz["Уязвимость устранена"]["Средний уровень опасности "]))
        self.low_ystr_labl.setText(str(codes_to_ystr_yaz["Уязвимость устранена"]["Низкий уровень опасности "]))

        self.krit_podtvpot_labl.setText(str(codes_to_status_yaz["Подтверждена"]["Критический уровень опасности "])+"/"+str(codes_to_status_yaz["Потенциальная"]["Критический уровень опасности "]))
        self.hig_podtvpot_labl.setText(str(codes_to_status_yaz["Подтверждена"]["Высокий уровень опасности "])+"/"+str(codes_to_status_yaz["Потенциальная"]["Высокий уровень опасности "]))
        self.mid_podtvpot_labl.setText(str(codes_to_status_yaz["Подтверждена"]["Средний уровень опасности "])+"/"+str(codes_to_status_yaz["Потенциальная"]["Средний уровень опасности "]))
        self.low_podtvpot_labl.setText(str(codes_to_status_yaz["Подтверждена"]["Низкий уровень опасности "])+"/"+str(codes_to_status_yaz["Потенциальная"]["Низкий уровень опасности "]))

        self.yiaz_ache.setText(str(codes_to_class_yazvim["Уязвимость архитектуры"]))
        self.yiaz_code.setText(str(codes_to_class_yazvim["Уязвимость кода"]))
        self.yiaz_manyfact.setText(str(codes_to_class_yazvim["Уязвимость многофакторная"]))
        self.yiaz_none.setText(str(codes_to_class_yazvim["nan"]))

        self.obn_2014.setText(str(code_to_date_obn["2014"]))
        self.obn_2015.setText(str(code_to_date_obn["2015"]))
        self.obn_2016.setText(str(code_to_date_obn["2016"]))
        self.obn_2017.setText(str(code_to_date_obn["2017"]))
        self.obn_2018.setText(str(code_to_date_obn["2018"]))
        self.obn_2019.setText(str(code_to_date_obn["2019"]))
        self.obn_2020.setText(str(code_to_date_obn["2020"]))
        self.obn_2021.setText(str(code_to_date_obn["2021"]))
        #
        self.ystr_2014.setText(str(code_to_date_ystr["2014"]["Уязвимость устранена"]))
        self.ystr_2015.setText(str(code_to_date_ystr["2015"]["Уязвимость устранена"]))
        self.ystr_2016.setText(str(code_to_date_ystr["2016"]["Уязвимость устранена"]))
        self.ystr_2017.setText(str(code_to_date_ystr["2017"]["Уязвимость устранена"]))
        self.ystr_2018.setText(str(code_to_date_ystr["2018"]["Уязвимость устранена"]))
        self.ystr_2019.setText(str(code_to_date_ystr["2019"]["Уязвимость устранена"]))
        self.ystr_2020.setText(str(code_to_date_ystr["2020"]["Уязвимость устранена"]))
        self.ystr_2021.setText(str(code_to_date_ystr["2021"]["Уязвимость устранена"]))




        chtenie = str((time.time() - start_time))
        self.message = QMessageBox.information(self, 'На чтение затрачено', 'На чтение затрачено '+ chtenie + ' секунд')



if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    ui1 = main_ui()
    ui1.show()
    sys.exit(app.exec_())