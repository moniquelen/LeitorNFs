import os
import shutil
from pathlib import Path
from PyQt5 import QtWidgets, QtGui, QtCore
from ler_pdfs import process_pdfs_in_folder

BASE_DIR = Path(__file__).parent

FOLDER_NFS = str(BASE_DIR / "NFs")
FOLDER_XMLS = str(BASE_DIR / "XMLs")
OUTPUT_EXCEL = str(BASE_DIR / "output.xlsx")
BACKGROUND_IMAGE = str(BASE_DIR / "img" / "background.svg")
LOGO_IMAGE = str(BASE_DIR / "img" / "imgLogo.png")
ICON_SUCCESS = str(BASE_DIR / "img" / "iconSucess.png")
ICON_ERROR = str(BASE_DIR / "img" / "iconError.png")
ICON_DELETE = str(BASE_DIR / "img" / "iconDelete.png")
ICON_NFS = str(BASE_DIR / "img" / "imgNFs.png")

class DropArea(QtWidgets.QLabel):
    def __init__(self, main_window, parent=None):
        super().__init__(parent)
        self.main_window = main_window
        self.setAcceptDrops(True)
        self.setAlignment(QtCore.Qt.AlignCenter)
        self.setFixedSize(350, 200)
        
        
        self.setStyleSheet("""
            QLabel {
                background-color: rgba(255, 255, 255, 0.9);
                border: 2px dashed #2a477f;
                border-radius: 20px;
                color: #666;
                font-size: 16px;
            }
            QLabel:hover {
                border-color: #62bfc5;
                background-color: rgba(245, 245, 245, 0.95);
            }
        """)
        self.setAlignment(QtCore.Qt.AlignCenter)
        
        self.texto = QtWidgets.QLabel("Arraste e solte arquivos PDF aqui\nou clique para selecionar", self)
        self.texto.setAlignment(QtCore.Qt.AlignCenter)
        self.texto.setStyleSheet("font-size: 14px; color: #888;")
        self.texto.setGeometry(0, 0, 350, 200)

    def mousePressEvent(self, event):
        self.selecionar_arquivos()

    def selecionar_arquivos(self):
        arquivos, _ = QtWidgets.QFileDialog.getOpenFileNames(self, "Selecionar Arquivos PDF", "", "Arquivos PDF (*.pdf)")
        if arquivos:
            self.main_window.arquivos = arquivos
            self.texto.setText(f"{len(arquivos)} arquivo(s) selecionado(s)")
            self.main_window.botao_submit.setEnabled(True)
            self.main_window.botao_reset.show()

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            event.ignore()


    def dropEvent(self, event):
        urls = event.mimeData().urls()
        arquivos = [url.toLocalFile() for url in urls if url.toLocalFile().endswith('.pdf')]
        if arquivos:
            self.main_window.arquivos = arquivos
            self.texto.setText(f"{len(arquivos)} arquivo(s) selecionado(s)")
            self.main_window.botao_submit.setEnabled(True)
            self.main_window.botao_reset.show()

class MainWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Leitor de Notas Fiscais")
        self.setWindowIcon(QtGui.QIcon(LOGO_IMAGE))
        self.setGeometry(100, 100, 1024, 768)
        self.setMinimumSize(800, 600)
        self.arquivos = []
        self.excel_path = ""

        # Configuração do background
        self.background_label = QtWidgets.QLabel(self)
        self.background_label.setAlignment(QtCore.Qt.AlignCenter)
        self.atualizar_background()
        
        # Widget central e layout
        central_widget = QtWidgets.QWidget(self)
        self.setCentralWidget(central_widget)
        layout = QtWidgets.QVBoxLayout(central_widget)
        layout.setContentsMargins(40, 40, 40, 40)
        layout.setSpacing(30)

        # Título
        titulo_label = QtWidgets.QLabel("Leitor de Notas Fiscais")
        titulo_label.setAlignment(QtCore.Qt.AlignCenter)
        titulo_label.setStyleSheet("""
            QLabel {
                font-size: 36px;
                font-weight: bold;
                color: white;
            }
        """)
        
        shadow = QtWidgets.QGraphicsDropShadowEffect()
        shadow.setBlurRadius(10)
        shadow.setOffset(2, 2)
        shadow.setColor(QtGui.QColor(0, 0, 0, 150))
        titulo_label.setGraphicsEffect(shadow)
        layout.addWidget(titulo_label)

        
        # Container para área de drop
        drop_container = QtWidgets.QWidget()
        drop_container.setFixedSize(400, 220) 
        drop_layout = QtWidgets.QVBoxLayout(drop_container)
        drop_layout.setContentsMargins(0, 0, 0, 0)

        # Área de drop
        self.drop_area = DropArea(self)
        drop_layout.addWidget(self.drop_area, alignment=QtCore.Qt.AlignCenter)

        # Adicionar a imagem dentro da DropArea
        self.image_label = QtWidgets.QLabel(self.drop_area)
        self.image_label.setPixmap(QtGui.QPixmap(ICON_NFS).scaled(60, 60, QtCore.Qt.KeepAspectRatio, QtCore.Qt.SmoothTransformation))
        self.image_label.setAlignment(QtCore.Qt.AlignCenter)
        self.image_label.setStyleSheet("background: none; border: none;")

        # Ajustar a ordem dos widgets
        self.drop_area.layout = QtWidgets.QVBoxLayout(self.drop_area)
        self.drop_area.layout.addWidget(self.image_label, alignment=QtCore.Qt.AlignCenter)

        # Texto indicativo
        self.drop_area.texto.setParent(None)
        self.drop_area.texto = QtWidgets.QLabel("Arraste e solte arquivos PDF aqui\nou clique para selecionar", self.drop_area)
        self.drop_area.texto.setAlignment(QtCore.Qt.AlignCenter)
        self.drop_area.texto.setStyleSheet("font-size: 14px; color: #888; border: none; background: none;")
        self.drop_area.layout.addWidget(self.drop_area.texto, alignment=QtCore.Qt.AlignCenter)

        drop_layout.addWidget(self.drop_area)

        # Botão de reset 
        self.botao_reset = QtWidgets.QPushButton(drop_container)
        self.botao_reset.setFixedSize(28, 28)
        self.botao_reset.setIcon(QtGui.QIcon(ICON_DELETE))
        self.botao_reset.setIconSize(QtCore.QSize(16, 16))
        self.botao_reset.setToolTip("Remover todos os arquivos")
        self.botao_reset.setStyleSheet("""
            QPushButton {
                background-color: #c30721;
                border-radius: 14px;
                border: none;
            }
            QPushButton:hover {
                background-color: #c0392b;
            }
        """)
        self.botao_reset.setContentsMargins(6, 6, 6, 6)
        self.botao_reset.hide()
        self.botao_reset.move(310, 20)
        self.botao_reset.clicked.connect(self.resetar_arquivos)

        layout.addWidget(drop_container, alignment=QtCore.Qt.AlignCenter)


        # Container para botões
        button_container = QtWidgets.QWidget()
        button_layout = QtWidgets.QHBoxLayout(button_container)
        button_layout.setContentsMargins(0, 0, 0, 0)
        button_layout.setSpacing(20)
        button_layout.addStretch()
        
        # Botões principais
        self.botao_submit = self.criar_botao("Processar NFs", "#74a3c5", False)
        self.botao_submit.clicked.connect(self.executar_automacao)
        button_layout.addWidget(self.botao_submit)

        self.botao_download = self.criar_botao("Exportar Excel", "#74a3c5", False)
        self.botao_download.clicked.connect(self.baixar_excel)
        button_layout.addWidget(self.botao_download)

        self.botao_abrir = self.criar_botao("Abrir Relatório", "#74a3c5", False)
        self.botao_abrir.clicked.connect(self.abrir_relatorio)
        button_layout.addWidget(self.botao_abrir)
        
        button_layout.addStretch()
        layout.addWidget(button_container)

        # Logo
        self.logo_label = QtWidgets.QLabel(self)
        logo_size = 80
        self.logo_label.setPixmap(QtGui.QPixmap(LOGO_IMAGE).scaled(
            logo_size, logo_size, 
            QtCore.Qt.KeepAspectRatio, 
            QtCore.Qt.SmoothTransformation))
        self.logo_label.setGeometry(
            self.width() - logo_size - 20,
            self.height() - logo_size - 20,
            logo_size,
            logo_size
        )
        self.logo_label.setStyleSheet("background: transparent;")

    def resetar_arquivos(self):
        self.arquivos = []
        self.drop_area.texto.setText("Arraste e solte arquivos PDF aqui\nou clique para selecionar")
        self.botao_submit.setEnabled(False)
        self.botao_reset.hide()
        self.botao_download.setEnabled(False)
        self.botao_abrir.setEnabled(False)

    def atualizar_background(self):
        pixmap = QtGui.QPixmap(BACKGROUND_IMAGE)
        scaled = pixmap.scaled(self.size(), QtCore.Qt.KeepAspectRatioByExpanding, QtCore.Qt.SmoothTransformation)
        self.background_label.setPixmap(scaled)
        self.background_label.setGeometry(0, 0, self.width(), self.height())

    def criar_botao(self, texto, cor, enabled=True):
        botao = QtWidgets.QPushButton(texto)
        botao.setEnabled(enabled)
        botao.setFixedSize(180, 45)
        botao.setStyleSheet(f"""
            QPushButton {{
                background-color: {cor};
                color: white;
                border: none;
                border-radius: 12px;
                font-size: 15px;
            }}
            QPushButton:hover {{
                background-color: {self.escurecer_cor(cor)};
            }}
            QPushButton:disabled {{
                background-color: #d9e6ef;
                color: #7f8c8d;
            }}
        """)
        return botao

    def escurecer_cor(self, hex_color, factor=0.8):
        color = QtGui.QColor(hex_color)
        h, s, v, _ = color.getHsv()
        new_v = int(v * factor)
        color.setHsv(h, s, new_v)
        return color.name()

    def resizeEvent(self, event):
        self.atualizar_background()
        logo_size = 80
        self.logo_label.setGeometry(
            self.width() - logo_size - 20,
            self.height() - logo_size - 20,
            logo_size,
            logo_size
        )
        super().resizeEvent(event)

    def executar_automacao(self):
        try:
            os.makedirs(FOLDER_NFS, exist_ok=True)
            os.makedirs(FOLDER_XMLS, exist_ok=True)

            for arquivo in self.arquivos:
                shutil.copy(arquivo, FOLDER_NFS)

            process_pdfs_in_folder(FOLDER_NFS, FOLDER_XMLS, OUTPUT_EXCEL)
            self.mostrar_mensagem("Sucesso", "Processamento concluído! O arquivo Excel está pronto para download.", ICON_SUCCESS)
            self.botao_download.setEnabled(True)

        except Exception as e:
            self.mostrar_mensagem("Erro", f"Erro ao executar a automação: {str(e)}", ICON_ERROR)

    def baixar_excel(self):
        destino, _ = QtWidgets.QFileDialog.getSaveFileName(self, "Salvar arquivo Excel", "", "Arquivos Excel (*.xlsx)")
        if destino:
            try:
                shutil.copy(OUTPUT_EXCEL, destino)
                self.excel_path = destino
                self.botao_abrir.setEnabled(True)
                self.mostrar_mensagem("Sucesso", "Arquivo Excel salvo com sucesso!", ICON_SUCCESS)
            except Exception as e:
                self.mostrar_mensagem("Erro", f"Erro ao salvar o arquivo Excel: {str(e)}", ICON_ERROR)

    def abrir_relatorio(self):
        try:
            if os.path.exists(self.excel_path):
                os.startfile(self.excel_path)
            else:
                self.mostrar_mensagem("Erro", "Arquivo Excel não encontrado!", ICON_ERROR)
        except Exception as e:
            self.mostrar_mensagem("Erro", f"Erro ao abrir o arquivo: {str(e)}", ICON_ERROR)

    def mostrar_mensagem(self, titulo, mensagem, icone):
        msg = QtWidgets.QMessageBox(self)
        msg.setWindowTitle(titulo)
        msg.setText(mensagem)
        msg.setIconPixmap(QtGui.QPixmap(icone).scaled(64, 64, QtCore.Qt.KeepAspectRatio))
        msg.exec_()

if __name__ == "__main__":
    app = QtWidgets.QApplication([])
    window = MainWindow()
    window.show()
    app.exec_()