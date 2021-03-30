import sys
from PyQt5.QtWidgets import QApplication, QWidget, QFileDialog, QMessageBox, QDesktopWidget
import pandas as pd
from pandas import ExcelWriter

class App(QWidget):
    def __init__(self):
        super().__init__()
        self.title = 'TeamsCSVToXLSX Converter'
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
        fileName, _ = QFileDialog.getOpenFileName(self, "Choice an file of Microsoft Teams", "",
                                                  "CSV (*.csv)", options=options)
        if fileName:
            path = str(fileName)[:int(str(fileName).rfind("/"))+1]
            dados = pd.read_csv(fileName)
            dados["Nome"] = dados["First Name"] + " " + dados["Last Name"]
            dados.drop(dados.filter(regex='Point|Feedback|First Name|Last Name|Email Address').columns, axis=1,
                       inplace=True)
            dados = dados[list(dados.columns.values)[::-1]]

            resultado = [0 for _ in range(dados.shape[0])]
            for i in range(dados.shape[0]):
                for j in range(len(dados.columns)):
                    try:
                        resultado[i] += int(dados.iloc[i][j])
                    except:
                        pass
            dados['Soma'] = resultado

            writer = ExcelWriter(path+'converted_gradesbook.xlsx')
            dados.to_excel(writer, 'grades')
            writer.save()
            msg = QMessageBox()
            msg.setWindowTitle("Sucess")
            msg.setIcon(QMessageBox.Information)
            msg.setText("SUCCESS! File saved with the name 'converted_gradesbook.xlsx', in the same place of the source file! ")
            msg.exec()

        sys.exit()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = App()
    sys.exit(app.exec_())