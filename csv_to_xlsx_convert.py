import sys
from PyQt5.QtWidgets import QApplication, QWidget, QFileDialog, QMessageBox, QDesktopWidget
import pandas as pd
from pandas import ExcelWriter

class App(QWidget):
    def __init__(self):
        super().__init__()
        self.title = 'TeamsCSVToXLSX Convert'
        self.initUI()

    def initUI(self):
        self.setWindowTitle(self.title)
        qr = self.frameGeometry()
        cp = QDesktopWidget().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())
        self.openFileNameDialog()
        self.show()

    def openFileNameDialog(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        fileName, _ = QFileDialog.getOpenFileName(self, "Choice a exported grade file of Microsoft Teams", "",
                                                  "CSV (*.csv)", options=options)
        if fileName:
            path = str(fileName)[:int(str(fileName).rfind("/"))+1]
            dados = pd.read_csv(fileName)
            
            dados[dados.columns[0]] += " " + dados[dados.columns[1]]

            list_col = [1,2]
            for i in range(4, len(dados.columns.values), 3):
                list_col.append(i)
                list_col.append(i+1)

            dados.drop(dados.columns[list_col], axis=1, inplace=True)

            list_2 = [list(dados.columns.values)[0]] + list(dados.columns.values)[1:][::-1]
            dados = dados[list_2]

            resultado = [0 for _ in range(dados.shape[0])]
            for i in range(dados.shape[0]):
                for j in range(len(dados.columns)):
                    try:
                        resultado[i] += int(dados.iloc[i][j])
                    except:
                        pass
            dados['Sum'] = resultado

            writer = ExcelWriter(path+'converted_grades.xlsx')
            dados.to_excel(writer, 'grades')
            writer.save()
            msg = QMessageBox()
            msg.setWindowTitle("Sucess!")
            msg.setIcon(QMessageBox.Information)
            msg.setText("SUCCESS! File saved with the name 'converted_grades.xlsx', in the same place of the source file! ")
            msg.exec()

        sys.exit()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = App()
    sys.exit(app.exec_())
