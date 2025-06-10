import win32com.client
import time
import subprocess
import logging
import os
import sys
from typing import Optional

# Configuração de Logging
def setup_logging(log_name: str):
    """
    Configura o sistema de logging padrão do Python para exibir logs no console
    e salvar em um arquivo dentro de uma pasta 'Logs' no diretório atual.

    Args:
        log_name (str): Um nome base para o arquivo de log. A data e hora
                             serão adicionadas a este nome.
    """
    # --- Configuração do Logging ---
    log_formatter = logging.Formatter(
        '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )

    # Diretório de logs relativo ao diretório de execução atual
    current_directory = os.getcwd()
    log_directory = os.path.join(current_directory, "Logs")

    try:
        os.makedirs(log_directory, exist_ok=True)
    except OSError as e:
        print(f"ERRO CRÍTICO: Não foi possível criar o diretório de logs {log_directory}: {e}", file=sys.stderr)
        sys.exit(f"Não foi possível criar o diretório de logs: {e}")

    timestamp = time.strftime("%Y-%m-%d_%H-%M-%S")
    log_file_name = f"{log_name}_{timestamp}.log"
    log_path = os.path.join(log_directory, log_file_name)

    print(f"Diretório de execução: {current_directory}")
    print(f"Salvando logs em: {log_path}")
    print(f"-----------------------------")

    root_logger = logging.getLogger()
    root_logger.setLevel(logging.INFO)

    if root_logger.hasHandlers():
        root_logger.handlers.clear()

    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setFormatter(log_formatter)
    console_handler.setLevel(logging.INFO)
    root_logger.addHandler(console_handler)

    try:
        file_handler = logging.FileHandler(log_path, mode='a', encoding='utf-8')
        file_handler.setFormatter(log_formatter)
        file_handler.setLevel(logging.DEBUG)
        root_logger.addHandler(file_handler)
        logging.info(f"Configuração de logging concluída. Arquivo de log: {log_path}")

    except Exception as e:
        logging.error(f"ERRO CRÍTICO: Falha ao configurar FileHandler para {log_path}: {e}", exc_info=True)

logger = logging.getLogger(__name__)

class SAPAutomatorError(Exception):
    """Exceção base para erros relacionados à automação SAP.
       Serve como um tipo genérico para capturar qualquer erro da nossa automação."""
    pass

class SAPConnectionError(SAPAutomatorError):
    """Erro específico que ocorre durante a tentativa de conexão ou login no SAP.
       Herda de SAPAutomatorError, permitindo capturar erros de conexão de forma mais específica."""
    pass

class SAPExecutionError(SAPAutomatorError):
    """Erro que ocorre durante a execução de ações no SAP *após* o login bem-sucedido (ex: erro em uma transação).
       Herda de SAPAutomatorError."""
    pass

class SAPAutomator:
    """
    Classe para automatizar interações com o SAP GUI Scripting.

    Esta classe encapsula toda a lógica para:
    - Abrir o SAP Logon.
    - Conectar a um sistema SAP específico.
    - Realizar o login.
    - Fornecer acesso à sessão SAP ativa para interações posteriores.
    """

    def __init__(self, sap_exe_path: str, system_name: str, client: str, language: str, user: str, password: str):
        """
        Inicializa o SAPAutomator com os parâmetros de configuração necessários.

        Args:
            sap_exe_path (str): Caminho completo para o executável saplogon.exe.
            system_name (str): Nome/Descrição do sistema SAP como aparece no SAP Logon. Usado para identificar qual conexão abrir.
            client (str): Mandante SAP (ex: "300").
            language (str): Idioma de login (ex: "PT", "EN").
            user (str): Nome de usuário SAP.
            password (str): Senha SAP.
        """

        self.sap_exe_path = sap_exe_path
        self.system_name = system_name
        self.client = client
        self.language = language
        self.user = user
        self.password = password

        self.session: Optional[win32com.client.CDispatch] = None # Armazenará o objeto da sessão SAP ativa.
        self._sap_gui_app: Optional[win32com.client.CDispatch] = None # Armazenará o objeto principal do SAP GUI Scripting.
        self._connection: Optional[win32com.client.CDispatch] = None # Armazenará o objeto da conexão específica ao sistema.

        logger.info("SAPAutomator inicializado.")

    def _open_sap_logon(self, timeout_seconds: int = 30) -> None:
        """
        Abre o SAP Logon (saplogon.exe) e aguarda ativamente até que o
        motor de scripting do SAP GUI esteja pronto para ser usado.

        Args:
            timeout_seconds (int): Tempo máximo (em segundos) para esperar o SAP GUI Scripting Engine ficar disponível. Default é 30 segundos.

        Raises:
            SAPConnectionError: Se o executável não for encontrado, ocorrer um erro ao abri-lo,
                                ou se o motor de scripting não ficar disponível dentro do timeout.
        """

        logger.info(f"Tentando abrir o SAP Logon em: {self.sap_exe_path}")
        try:
            subprocess.Popen(self.sap_exe_path)
        except FileNotFoundError:
            logger.error(f"Erro: Executável SAP não encontrado em {self.sap_exe_path}")
            raise SAPConnectionError(f"Executável SAP não encontrado em {self.sap_exe_path}")
        except Exception as erro:
            logger.error(f"Erro inesperado ao tentar abrir o SAP Logon: {erro}")
            raise SAPConnectionError(f"Erro inesperado ao tentar abrir o SAP Logon: {erro}")

        start_time = time.time() # Marca o tempo de início da espera.
        # Loop continua enquanto o tempo decorrido for menor que o timeout.
        while time.time() - start_time < timeout_seconds:
            try:
                self._sap_gui_app = win32com.client.GetObject("SAPGUI")
                if self._sap_gui_app:
                    logger.info("SAP GUI Scripting Engine obtido com sucesso.")
                    return
            except Exception:
                time.sleep(1)

        logger.error(f"Timeout: SAP GUI Scripting Engine não ficou disponível em {timeout_seconds} segundos.")
        raise SAPConnectionError(f"Timeout: SAP GUI Scripting Engine não ficou disponível em {timeout_seconds} segundos.")

    def _connect_to_system(self) -> None:
        """
        Usa o objeto SAP GUI Scripting Engine (obtido anteriormente ou tentando obter agora)
        para se conectar ao sistema SAP específico definido em 'self.system_name'.

        Raises:
            SAPConnectionError: Se não for possível obter o motor de scripting, ou se a conexão
                                ao sistema especificado falhar, ou se nenhuma sessão for criada.

        """

        try:
            application = self._sap_gui_app.GetScriptingEngine
            logger.info(f"Tentando conectar ao ambiente: {self.system_name}")
            self._connection = application.OpenConnection(self.system_name, True)

            if not self._connection:
                 raise SAPConnectionError(f"Falha ao abrir conexão para o sistema {self.system_name}. Objeto de conexão é nulo.")

            if self._connection.Sessions.Count > 0:
                self.session = self._connection.Children(0)
                logger.info("Sessão SAP obtida com sucesso.")
            else:
                logger.warning(f"Nenhuma sessão encontrada imediatamente para {self.system_name}. Aguardando...")
                time.sleep(3)

                if self._connection.Sessions.Count > 0:
                     self.session = self._connection.Children(0)
                     logger.info("Sessão SAP obtida após espera adicional.")
                else:
                     raise SAPConnectionError(f"Nenhuma sessão encontrada para a conexão {self.system_name} após a abertura e espera.")

        except Exception as erro:
            logger.error(f"Ocorreu um erro ao realizar a conexão com o SAP ({self.system_name}): {erro}")
            raise SAPConnectionError(f"Erro ao conectar ao sistema SAP {self.system_name}: {erro}")

    def _login(self) -> None:
        """
        Preenche os campos de Mandante, Usuário, Senha e Idioma na tela de login
        da sessão SAP ativa e pressiona Enter para logar.

        Raises:
            SAPConnectionError: Se a sessão não estiver disponível (self.session é None) ou
                                se ocorrer um erro genérico durante o login (ex: falha de comunicação COM).
            ValueError: Se um elemento específico da tela de login (campo de texto, botão)
                        não for encontrado via findById, indicando que a tela pode estar diferente
                        do esperado ou o seletor está incorreto.
        """
        if not self.session:
            logger.error("Tentativa de login sem uma sessão SAP ativa.")
            raise SAPConnectionError("Sessão SAP não está inicializada para login.")

        try:
            if not hasattr(self.session, 'findById'):
                 raise SAPConnectionError("Objeto de sessão inválido ou não suporta 'findById'. Ocorreu um problema na conexão.")

            try:
                # Tenta localizar o campo de texto do mandante usando seu ID.
                mandt_field = self.session.findById("wnd[0]/usr/txtRSYST-MANDT")
            except Exception:
                 logger.warning("Tela de login padrão não detectada ou campo Mandante não encontrado.")
                 raise SAPConnectionError("Mandante não encontrado na tela")

            logger.info(f"Realizando login no mandante {self.client} com usuário {self.user}")
            mandt_field.text = self.client
            self.session.findById("wnd[0]/usr/txtRSYST-BNAME").text = self.user
            self.session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = self.password
            self.session.findById("wnd[0]/usr/txtRSYST-LANGU").text = self.language
            self.session.findById("wnd[0]/tbar[0]/btn[0]").press()

            time.sleep(1) 
            status_bar = self.session.findById("wnd[0]/sbar")
            # Verifica o tipo da mensagem. 'E' (Error), 'A' (Abort), 'X' (Exit) indicam problemas.
            if status_bar.MessageType.lower() in ('e', 'a', 'x'):
                 error_message = status_bar.Text
                 logger.error(f"Erro de login SAP detectado na barra de status: {error_message}")
                 raise SAPConnectionError(f"Erro de login SAP: {error_message}")

            logger.info("Login SAP realizado com sucesso.")

        except Exception as erro:
            logger.error(f"Ocorreu um erro ao realizar o Login SAP: {erro}", exc_info=True) # exc_info=True inclui o traceback no log para depuração.
            if "findById" in str(erro) or "element" in str(erro).lower():
                 raise ValueError(f"Erro ao encontrar elemento da GUI durante o login: {erro}")
            else:
                 raise SAPConnectionError(f"Erro inesperado durante o login SAP: {erro}")

    def initialize_connection(self, open_new_logon: bool = True) -> Optional[win32com.client.CDispatch]:
        """
        Orquestra o processo completo: abre o SAP Logon (opcionalmente),
        conecta ao sistema e realiza o login.

        Args:
             open_new_logon (bool): Se True (padrão), tenta abrir uma nova instância do saplogon.exe
                                   usando _open_sap_logon. Se False, assume que o SAP Logon
                                   já está aberto e tenta conectar diretamente.

        Returns:
            Optional[win32com.client.CDispatch]: Retorna o objeto da sessão SAP ativa se todo o
                                                 processo for bem-sucedido. Retorna None ou lança exceção em caso de falha.
                                                 (Com a implementação atual, exceções são preferidas sobre retornar None).

        Raises:
            SAPConnectionError: Se qualquer etapa do processo de abertura, conexão ou login falhar.
            ValueError: Se elementos da GUI de login não forem encontrados.
            Exception: Captura e relança qualquer outra exceção inesperada.
        """

        try:
            if open_new_logon:
                self._open_sap_logon()

            self._connect_to_system()
            self._login()

            logger.info("Inicialização da conexão SAP concluída com sucesso.")
            return self.session

        except (SAPConnectionError, ValueError, Exception) as e:
            logger.error(f"Falha ao inicializar a conexão SAP: {e}")
            self.session = None
            self._connection = None
            self._sap_gui_app = None
            raise e

    def get_session(self) -> Optional[win32com.client.CDispatch]:
        """Retorna o objeto de sessão SAP ativo (self.session).
           Útil se o chamador precisar acessar a sessão depois da inicialização.
        """
        if not self.session:
             logger.warning("Tentativa de obter sessão, mas não está inicializada ou a conexão falhou anteriormente.")
        return self.session

    def close_connection(self) -> None:
        """
        Tenta fechar a conexão SAP e a sessão associada de forma limpa.
        Nota: Geralmente não fecha o SAP Logon em si, apenas a conexão específica.
        """
        try:
            # Verifica se existe uma sessão ativa.
            if self.session:
                logger.info("Fechando a sessão SAP.")
                self.session.findById("wnd[0]").close() #Fechar a conexão geralmente fecha as sessões filhas.

            # Verifica se existe uma conexão ativa.
            if self._connection:
                 logger.info("Fechando a conexão SAP.")
                 self._connection.CloseConnection()

            self.session = None
            self._connection = None

            logger.info("Conexão SAP fechada.")
        except Exception as e:
            logger.error(f"Erro ao tentar fechar a conexão SAP: {e}")
            self.session = None
            self._connection = None

    def check_msgBox(self, session):
        """
        Verifica continuamente a existência da janela de mensagem "wnd[1]" (pop-up)
        e pressiona Enter (VKey 0) se ela existir.

        Args:
            session: O objeto de sessão ativa do SAP.

        Returns:
            bool: True se uma mensagem foi encontrada e tratada, False caso contrário.
                (Retorno adicionado para possível lógica condicional)
        """
        msg_handled = False
        try:
            while True:
                msgBox = session.findById("wnd[1]", False)
                if msgBox:
                    #logging.info(f"Janela de mensagem wnd[1] encontrada (Tipo: {msgBox.Type}). Pressionando Enter.")
                    msgBox.sendVKey(0) #Enter
                    msg_handled = True
                else:
                    break
        except Exception as e:
            sbar_text = ""
            try:
                sbar = session.findById("wnd[0]/sbar", False)
                if sbar:
                    sbar_text = sbar.text
            except:
                pass
            logging.error(f"Erro na função check_msgBox (Statusbar: '{sbar_text}'): {e}")
            # raise # Descomente se o script deve parar aqui

        return msg_handled
