# SAP Automator
*Automação de SAP GUI com Python*

---

Este projeto consiste em um módulo Python que robustece e simplifica a automação de tarefas no **SAP GUI para Windows**. Ele utiliza a API de Scripting nativa do SAP e encapsula toda a lógica de conexão, login, gerenciamento de sessão e tratamento de erros em uma classe `SAPAutomator` fácil de usar.

O script também inclui uma função de configuração de **logging** que gera registros detalhados da execução, tanto no console quanto em arquivos, facilitando a depuração e o rastreamento de processos.

---

## Funcionalidades Principais
- **Conexão Simplificada:** Abre o SAP Logon e conecta-se ao sistema desejado com uma única chamada de método.
- **Login Automático:** Preenche automaticamente as credenciais (mandante, usuário, senha, idioma) e efetua o login.
- **Gerenciamento de Sessão:** Fornece acesso direto ao objeto de sessão do SAP para interações subsequentes.
- **Tratamento de Erros Robusto:** Utiliza exceções personalizadas para diferenciar erros de conexão de erros de execução.
- **Manuseio de Pop-ups:** Inclui uma função para detectar e fechar caixas de diálogo que possam aparecer durante a automação.
- **Logging Integrado:** Cria arquivos de log com timestamp, além de exibir as informações no console.
- **Fechamento Seguro:** Garante que a conexão com o SAP seja encerrada de forma limpa, mesmo em caso de erros.

---

## Pré-requisitos
1.  Python 3.x
2.  Biblioteca `pywin32`
3.  SAP GUI para Windows
4.  SAP GUI Scripting Habilitado (servidor e cliente)

### Como Habilitar o Scripting no Cliente SAP
1.  Abra o **SAP Logon**.
2.  Clique no ícone no canto superior esquerdo e vá em **Opções**.
3.  Navegue até **Acessibilidade e Scripts > Scripts**.
4.  Em **Configurações do Usuário**, marque a opção **Ativar scripts**.
5.  Desmarque as opções **Notificar quando um script se conectar ao SAP GUI** e **Notificar quando um script abrir uma conexão**.
6.  Clique em **OK** para salvar.

---

## Instalação da Dependência
Para instalar a biblioteca `pywin32`, utilize o pip:
```bash
pip install pywin32
```

---

## Como Usar
1.  Salve o código do módulo de automação em um arquivo Python, por exemplo, `sap_automation.py`.
2.  Crie um segundo arquivo, por exemplo, `main.py`, para executar sua automação. Importe os componentes necessários do arquivo `sap_automation.py`.

Abaixo está um exemplo de como seria o seu arquivo `main.py`:

```python
# main.py
import logging
# Importe as classes e funções do seu módulo
from sap_automation import SAPAutomator, setup_logging, SAPConnectionError

def run_automation():
    """
    Função principal que executa o processo de automação.
    """
    # 1. Configura o sistema de logging
    setup_logging("relatorio_mara")

    # --- Configurações da Conexão ---
    sap_exe_path = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
    system_name = "S4H - SAP S/4HANA"
    client = "100"
    language = "PT"
    user = "seu_usuario_sap"
    password = "sua_senha_sap"

    automator = None
    try:
        # 2. Inicializa o automator
        automator = SAPAutomator(sap_exe_path, system_name, client, language, user, password)

        # 3. Conecta e loga no SAP
        session = automator.initialize_connection()
        logging.info("Conexão e login realizados com sucesso!")

        # 4. Automatiza tarefas
        logging.info("Iniciando a transação SE16...")
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nSE16"
        session.findById("wnd[0]").sendVKey(0)

        session.findById("wnd[0]/usr/ctxtDATABROWSE-TABLENAME").text = "MARA"
        session.findById("wnd[0]/tbar[1]/btn[8]").press()

        logging.info("Transação SE16 executada com sucesso.")

    except SAPConnectionError as e:
        logging.error(f"Erro de Conexão/Login: {e}")
    except Exception as e:
        logging.error(f"Ocorreu um erro inesperado: {e}", exc_info=True)
    finally:
        # 5. Garante o fechamento da conexão
        if automator:
            logging.info("Fechando a conexão SAP.")
            automator.close_connection()
            logging.info("Processo finalizado.")

if __name__ == "__main__":
    run_automation()
```

---

## Estrutura do Módulo
- **`setup_logging(log_name)`**: Função que configura o logging para console e arquivo.
- **Exceções Customizadas**:
    - `SAPAutomatorError`: Exceção base.
    - `SAPConnectionError`: Para erros de conexão ou login.
    - `SAPExecutionError`: Para erros pós-login (em desenvolvimento).
- **Classe `SAPAutomator`**:
    - `__init__(...)`: Inicializa com os parâmetros de conexão.
    - `initialize_connection()`: Orquestra todo o processo de conexão e login.
    - `get_session()`: Retorna a sessão ativa.
    - `close_connection()`: Fecha a sessão e a conexão de forma segura.
    - `check_msgBox()`: Função utilitária para tratar pop-ups.

---

## Licença
Este projeto é de código aberto. Sinta-se à vontade para usá-lo e modificá-lo conforme necessário.
