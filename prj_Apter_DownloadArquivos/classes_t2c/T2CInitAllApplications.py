from botcity.web import WebBot, Browser
from botcity.core import DesktopBot
from prj_Apter_DownloadArquivos.classes_t2c.utils.T2CMaestro import T2CMaestro, LogLevel, ErrorType
from prj_Apter_DownloadArquivos.classes_t2c.utils.T2CExceptions import BusinessRuleException
from prj_Apter_DownloadArquivos.classes_t2c.sqlite.T2CSqliteQueue import T2CSqliteQueue
from prj_Apter_DownloadArquivos.classes_t2c.email.T2CSendEmail import T2CSendEmail

from prj_Apter_DownloadArquivos.classes_t2c.sqlserver.T2CSqlAnaliticoSintetico import T2CSqlAnaliticoSintetico


class T2CInitAllApplications:
    """
    Classe feita para ser invocada principalmente no começo de um processo, para iniciar os processos necessários para a automação.
    """

    #Iniciando a classe, pedindo um dicionário config e o bot que vai ser usado e enviando uma exceção caso nenhum for informado
    def __init__(self, arg_dictConfig:dict, arg_clssMaestro:T2CMaestro, arg_botWebbot:WebBot=None, arg_botDesktopbot:DesktopBot=None, arg_clssSqliteQueue:T2CSqliteQueue=None):
        """
        Inicializa a classe T2CInitAllApplications.

        Parâmetros:
        - arg_dictConfig (dict): dicionário de configuração.
        - arg_clssMaestro (T2CMaestro): instância de T2CMaestro.
        - arg_botWebbot (WebBot): instância de WebBot (opcional, default=None).
        - arg_botDesktopbot (DesktopBot): instância de DesktopBot (opcional, default=None).
        - arg_clssSqliteQueue (T2CSqliteQueue): instância de T2CSqliteQueue (opcional, default=None).

        Raises:
        - Exception: caso nenhum bot seja fornecido.
        """
        
        if(arg_botWebbot is None and arg_botDesktopbot is None): raise Exception("Não foi possível inicializar a classe, forneça pelo menos um bot")
        else:
            self.var_botWebbot = arg_botWebbot
            self.var_botDesktopbot = arg_botDesktopbot
            self.var_dictConfig = arg_dictConfig
            self.var_clssMaestro = arg_clssMaestro
            self.var_clssSqliteQueue = arg_clssSqliteQueue


    def execute(self, arg_boolFirstRun=False, arg_clssT2CSql:T2CSqlAnaliticoSintetico=None, arg_clssSendEmail:T2CSendEmail=None):
        """
        Executa a inicialização dos aplicativos necessários.

        Parâmetros:
        - arg_boolFirstRun (bool): indica se é a primeira execução (default=False).
        - arg_clssSqliteQueue (T2CSqliteQueue): instância da classe T2CSqliteQueue (opcional, default=None).

        Raises:
        - BusinessRuleException: em caso de erro de regra de negócio.
        - Exception: em caso de erro geral.

        Observação:
        - Edite o valor da variável `var_intMaxTentativas` no arquivo Config.xlsx.
        """

        if(arg_boolFirstRun):
            var_clssItem = arg_clssSendEmail
            var_clssItem.read_email_Inbox(arg_clssT2CSql=arg_clssT2CSql)

        #Edite o valor dessa variável a no arquivo Config.xlsx
        var_intMaxTentativas = self.var_dictConfig["MaxRetryNumber"]
        
        for var_intTentativa in range(var_intMaxTentativas):
            try:
                self.var_clssMaestro.write_log("Iniciando aplicativos, tentativa " + (var_intTentativa+1).__str__())
                #Insira aqui seu código para iniciar os aplicativos
            except BusinessRuleException as exception:
                self.var_clssMaestro.write_log(arg_strMensagemLog="Erro de negócio: " + str(exception), arg_enumLogLevel=LogLevel.ERROR, arg_enumErrorType=ErrorType.BUSINESS_ERROR)

                raise
            except Exception as exception:
                self.var_clssMaestro.write_log(arg_strMensagemLog="Erro, tentativa " + (var_intTentativa+1).__str__() + ": " + str(exception), arg_enumLogLevel=LogLevel.ERROR, arg_enumErrorType=ErrorType.APP_ERROR)

                if(var_intTentativa+1 == var_intMaxTentativas): raise
                else: 
                    #inclua aqui seu código para tentar novamente
                    
                    continue
            else:
                self.var_clssMaestro.write_log("Aplicativos iniciados, continuando processamento...")
                
                break
            
