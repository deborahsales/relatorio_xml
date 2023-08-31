#pyuic5 -x Window.ui -o Window.py
#pyinstaller --windowed Main.py -i logo_IP.ico
#cxfreeze Main.py --target-dir "Gerador Relatório de Uso" --icon=logo_IP.ico --base-name=WIN32GUI

from PyQt5 import QtWidgets
from Window import Ui_MainWindow
from xml_codigo import alunos

class MainWindow(QtWidgets.QMainWindow):
    def __init__(self):
      super(MainWindow, self).__init__()
      self.ui = Ui_MainWindow()
      self.ui.setupUi(self)
      self.setup_program_ui()
    
    def setup_program_ui(self):
      self.ui.botao_extraido.clicked.connect(lambda: self.abrir_activity())  
      self.ui.botao_gerar.clicked.connect(lambda: self.salvar_arquivo())

    def abrir_activity(self):
      self._arquivo_activity = QtWidgets.QFileDialog.getOpenFileName()[0]
      self.ui.activity.setText(self._arquivo_activity)
    
    def salvar_arquivo(self):      
      self._nome_arquivo = QtWidgets.QFileDialog.getSaveFileName()[0]
      alunos(self._arquivo_activity,self._nome_arquivo)
      
      QtWidgets.QMessageBox.about(self,
             'Concluído', f'Relatório gerado com sucesso!!')
      
      
      self.ui.activity.setText('')
      