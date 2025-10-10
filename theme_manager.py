import sys
from PyQt5.QtCore import QTimer, QObject, pyqtSignal

# Variável para armazenar o estado atual do tema
CURRENT_THEME = ""

def is_windows_dark_mode():
    """
    Verifica no registro do Windows se o tema de aplicativos está escuro.
    Retorna True se estiver escuro, False se estiver claro ou se a chave não for encontrada.
    """
    if sys.platform != 'win32':
        # Esta função é específica para Windows. Retorna False para outros sistemas.
        return False
        
    try:
        import winreg
        # Chave do registro que armazena a preferência de tema para apps
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r'Software\Microsoft\Windows\CurrentVersion\Themes\Personalize')
        # O valor 'AppsUseLightTheme' é 0 para escuro e 1 para claro.
        value, _ = winreg.QueryValueEx(key, 'AppsUseLightTheme')
        return value == 0
    except (FileNotFoundError, ImportError):
        # Se a chave não existir ou o módulo winreg não for encontrado, assume o tema claro.
        return False

def load_stylesheet(theme_name):
    """Carrega o arquivo QSS correspondente ao nome do tema."""
    try:
        with open(f'{theme_name}.qss', 'r') as f:
            return f.read()
    except FileNotFoundError:
        print(f"Aviso: Arquivo de tema '{theme_name}.qss' não encontrado.")
        return ""

class ThemeManager(QObject):
    # Sinal que será emitido quando o tema mudar
    theme_changed = pyqtSignal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.current_theme = self.get_current_theme_name()
        
        # Cria um timer para verificar a mudança de tema a cada 2 segundos
        self.timer = QTimer(self)
        self.timer.setInterval(2000) # 2000 ms = 2 segundos
        self.timer.timeout.connect(self.check_for_theme_change)
        self.timer.start()

    def get_current_theme_name(self):
        """Retorna 'dark' ou 'light' com base na configuração do sistema."""
        return "dark" if is_windows_dark_mode() else "light"

    def check_for_theme_change(self):
        """Verifica se o tema do sistema mudou e emite um sinal se necessário."""
        new_theme = self.get_current_theme_name()
        if new_theme != self.current_theme:
            self.current_theme = new_theme
            print(f"Tema do sistema alterado para: {new_theme}")
            self.theme_changed.emit(new_theme)

    def apply_initial_theme(self, app):
        """Aplica o tema detectado quando o aplicativo é iniciado."""
        print(f"Aplicando tema inicial: {self.current_theme}")
        stylesheet = load_stylesheet(self.current_theme)
        app.setStyleSheet(stylesheet)