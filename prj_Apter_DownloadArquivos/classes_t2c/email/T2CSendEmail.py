from prj_Apter_DownloadArquivos.classes_t2c.sqlserver.T2CSqlAnaliticoSintetico import T2CSqlAnaliticoSintetico
from prj_Apter_DownloadArquivos.classes_t2c.utils.T2CMaestro import T2CMaestro, LogLevel, ErrorType
from imap_tools import MailBox, AND, MailMessageFlags
from botcity.plugins.email import BotEmailPlugin
from pathlib import Path
import datetime
import regex



ROOT_DIR = Path(__file__).parent.parent.parent.__str__()

class T2CSendEmail:
    """
    Classe responsável pelo envio de email, enviando usuario, senha e qual servidor vai ser usado
    """
    def __init__(self, arg_strNomeProcesso:str, arg_dictConfig:dict, arg_clssMaestro:T2CMaestro):
        """
        Inicializa a classe T2CSendEmail.

        Parâmetros:
        - arg_strNomeProcesso (str): nome do processo.
        - arg_strEmailServerSmtp (str): endereço do servidor SMTP.
        - arg_intEmailPortaSmtp (int): porta do servidor SMTP.
        - arg_strUsuario (str): nome de usuário para autenticação.
        - arg_strSenha (str): senha para autenticação.
        - arg_clssMaestro (T2CMaestro): instância de T2CMaestro.
        """
        self.var_strEmailServerSmtp = arg_dictConfig['EmailServerSmtp']
        self.var_intEmailPortaSmtp = arg_dictConfig['EmailPortaSmtp']
        self.var_strEmailServerImap = arg_dictConfig['EmailServerImap']
        self.var_intEmailPortaImap = arg_dictConfig['EmailPortaImap']
        self.var_strUsuario = arg_dictConfig['EmailUsuario']
        self.var_strSenha = arg_dictConfig['EmailSenha']
        self.var_clssMaestro = arg_clssMaestro
        self.var_strNomeProcesso = arg_strNomeProcesso

    
    def send_email_inicial(self, arg_strEnvioPara:str, arg_strCC:str=None, arg_strBCC:str=None):
        """
        Envia o email inicial do robô, apenas precisando informar quem deve receber (separado por ;) e o nome do robô

        Parâmetros:
        - arg_strEnvioPara (str): destinatários separados por ';'.
        - arg_strCC (str): destinatários em cópia separados por ';'. (opcional, default=None)
        - arg_strBCC (str): destinatários em cópia oculta separados por ';'. (opcional, default=None)
        """

        #Lendo template
        var_fileTemplate = open(ROOT_DIR + "\\resources\\templates\\Email_Inicio.txt", "r")
        var_strEmailTexto = var_fileTemplate.read()
        var_fileTemplate.close()

        var_strEmailTexto = var_strEmailTexto.replace("*NOME_ROBO*", self.var_strNomeProcesso)
        var_strEmailAssunto = "Inicio execução: " + self.var_strNomeProcesso
        
        var_listEnvioPara = arg_strEnvioPara.split(";") if(arg_strEnvioPara is not None) else []
        var_listCC = arg_strCC.split(";") if(arg_strCC is not None) else []
        var_listBCC = arg_strBCC.split(";") if(arg_strBCC is not None) else []

        #Configurando classe do email, separado para emitir logs diferentes
        try:
            self.var_clssMaestro.write_log("Configurando classe de email para envio")
            var_clssEmailController = BotEmailPlugin()
            var_clssEmailController.configure_smtp(self.var_strEmailServerSmtp, self.var_intEmailPortaSmtp)
            var_clssEmailController.login(self.var_strUsuario, self.var_strSenha)
            self.var_clssMaestro.write_log("Email configurado")
        except Exception:
            self.var_clssMaestro.write_log(arg_strMensagemLog="Erro Configurando email: " + str(Exception), arg_enumLogLevel=LogLevel.ERROR, arg_enumErrorType=ErrorType.APP_ERROR)
            raise

        #Envia o email inicial
        try:
            self.var_clssMaestro.write_log("Enviando email inicial")
            var_clssEmailController.send_message(text_content=var_strEmailTexto, 
                                                 to_addrs=var_listEnvioPara, 
                                                 cc_addrs=var_listCC, 
                                                 bcc_addrs=var_listBCC, 
                                                 use_html=True, 
                                                 subject=var_strEmailAssunto)
            self.var_clssMaestro.write_log("Email enviado com sucesso")
        except Exception:
            self.var_clssMaestro.write_log(arg_strMensagemLog="Erro enviando email inicial: " + str(Exception), arg_enumLogLevel=LogLevel.ERROR, arg_enumErrorType=ErrorType.APP_ERROR)
            raise
  
    def send_email_final(self, arg_strHorarioInicio:str, arg_strHorarioFim:str, arg_strEnvioPara:str, arg_strCC:str=None, arg_strBCC:str=None, arg_listAnexos:list=None, arg_boolSucesso:bool=True):
        """
        Envio do e-mail de finalização do robô.
        Recebendo o horário do início da execução, o horário final, para quem é necessário enviar (separado por ;) e os relatórios finais
        
        Parâmetros:
        - arg_strHorarioInicio (str): horário de início.
        - arg_strHorarioFim (str): horário de fim.
        - arg_strEnvioPara (str): destinatários separados por ';'.
        - arg_strCC (str): destinatários em cópia separados por ';'. (opcional, default=None)
        - arg_strBCC (str): destinatários em cópia oculta separados por ';'. (opcional, default=None)
        - arg_listAnexos (list): lista de anexos (opcional, default=None).
        - arg_boolSucesso (bool): indica se a execução foi bem-sucedida (opcional, default=True).
        """
        
        #Lendo template
        var_fileTemplate = open(ROOT_DIR + "\\resources\\templates\\Email_Final.txt", "r")
        var_strEmailTexto = var_fileTemplate.read()
        var_fileTemplate.close()

        var_strStatusFinalizacao = "com sucesso" if arg_boolSucesso else "com erros"
        var_strEmailTexto = var_strEmailTexto.replace("*NOME_ROBO*", self.var_strNomeProcesso).replace("*DATAHORA_INI*", arg_strHorarioInicio).replace("*DATAHORA_FIM*", arg_strHorarioFim).replace("*FINALIZACAO*", var_strStatusFinalizacao)
        var_strEmailTexto.replace("*NOME_ROBO*", self.var_strNomeProcesso)
        var_strEmailAssunto = "Finalização da execução: " + self.var_strNomeProcesso

        var_listEnvioPara = arg_strEnvioPara.split(";") if(arg_strEnvioPara is not None) else []
        var_listCC = arg_strCC.split(";") if(arg_strCC is not None) else []
        var_listBCC = arg_strBCC.split(";") if(arg_strBCC is not None) else []

        #Configurando classe do email, separado para emitir logs diferentes
        try:
            self.var_clssMaestro.write_log("Configurando classe de email para envio")
            var_clssEmailController = BotEmailPlugin()
            var_clssEmailController.configure_smtp(self.var_strEmailServerSmtp, self.var_intEmailPortaSmtp)
            var_clssEmailController.login(self.var_strUsuario, self.var_strSenha)
            self.var_clssMaestro.write_log("Email configurado")
        except Exception:
            self.var_clssMaestro.write_log(arg_strMensagemLog="Erro Configurando email: " + str(Exception), arg_enumLogLevel=LogLevel.ERROR, arg_enumErrorType=ErrorType.APP_ERROR)
            raise
        
        #Envia o email de finalização
        try:
            self.var_clssMaestro.write_log("Enviando email de finalização")
            var_clssEmailController.send_message(text_content=var_strEmailTexto, 
                                                 to_addrs=var_listEnvioPara, 
                                                 cc_addrs=var_listCC, 
                                                 bcc_addrs=var_listBCC, 
                                                 attachments=arg_listAnexos, 
                                                 use_html=True, 
                                                 subject=var_strEmailAssunto)
            self.var_clssMaestro.write_log("Email enviado com sucesso")
        except Exception:
            self.var_clssMaestro.write_log(arg_strMensagemLog="Erro enviando email de finalização: " + str(Exception), arg_enumLogLevel=LogLevel.ERROR, arg_enumErrorType=ErrorType.APP_ERROR)
            raise

    # def send_email_erro(self, arg_strEnvioPara:str, arg_listAnexos:list, arg_strDetalhesErro:str, arg_boolBusiness:bool=False, arg_strCC:str=None, arg_strBCC:str=None):
    #     """
    #     Envio do e-mail em casos de erro durante a execução do robô.

    #     Parâmetros:
    #     - arg_strEnvioPara (str): destinatários separados por ';'.
    #     - arg_listAnexos (list): lista de anexos.
    #     - arg_strDetalhesErro (str): detalhes do erro.
    #     - arg_boolBusiness (bool): indica se o erro é de regra de negócio. (opcional, default=False)
    #     - arg_strCC (str): destinatários em cópia separados por ';'. (opcional, default=None)
    #     - arg_strBCC (str): destinatários em cópia oculta separados por ';'. (opcional, default=None)
    #     """
        
    #     #Lendo template
    #     var_fileTemplate = open(ROOT_DIR + "\\resources\\templates\\Email_ErroEncontrado.txt", "r")
    #     var_strEmailTexto = var_fileTemplate.read()
    #     var_fileTemplate.close()

    #     var_strEmailTexto = var_strEmailTexto.replace("*NOME_ROBO*", self.var_strNomeProcesso).replace("*ERRO_DETALHES*", arg_strDetalhesErro)
    #     var_strEmailTexto = var_strEmailTexto.replace("*ERRO_TIPO*", "ERRO DE REGRA DE NEGÓCIO") if arg_boolBusiness else var_strEmailTexto.replace("*ERRO_TIPO*", "ERRO INESPERADO")
    #     var_strEmailAssunto = "Erro durante a execução: " + self.var_strNomeProcesso
        
    #     var_listEnvioPara = arg_strEnvioPara.split(";") if(arg_strEnvioPara is not None) else []
    #     var_listCC = arg_strCC.split(";") if(arg_strCC is not None) else []
    #     var_listBCC = arg_strBCC.split(";") if(arg_strBCC is not None) else []

    #     #Configurando classe do email, separado para emitir logs diferentes
    #     try:
    #         self.var_clssMaestro.write_log("Configurando classe de email para envio")
    #         var_clssEmailController = BotEmailPlugin()
    #         var_clssEmailController.configure_smtp(self.var_strEmailServerSmtp, self.var_intEmailPortaSmtp)
    #         var_clssEmailController.login(self.var_strUsuario, self.var_strSenha)
    #         self.var_clssMaestro.write_log("Email configurado")
    #     except Exception:
    #         self.var_clssMaestro.write_log("Erro Configurando email:")
    #         self.var_clssMaestro.write_log(str(Exception))
    #         raise

    #     #Envia o email inicial
    #     try:
    #         self.var_clssMaestro.write_log("Enviando email de erro")
    #         var_clssEmailController.send_message(text_content=var_strEmailTexto, 
    #                                              to_addrs=var_listEnvioPara, 
    #                                              cc_addrs=var_listCC, 
    #                                              bcc_addrs=var_listBCC, 
    #                                              use_html=True, 
    #                                              subject=var_strEmailAssunto,
    #                                              attachments=arg_listAnexos)
    #         self.var_clssMaestro.write_log("Email enviado com sucesso")
    #     except Exception:
    #         self.var_clssMaestro.write_log(arg_strMensagemLog="Erro enviando email de erro: " + str(Exception), arg_enumLogLevel=LogLevel.ERROR, arg_enumErrorType=ErrorType.APP_ERROR)
    #         raise

    def send_email_erro(self, arg_strItemErro:str,arg_strEnvioPara:str,arg_listAnexos:list ,arg_strUrldrive:str,arg_strDetalhesErro:list, arg_strDataInit:str, arg_strDataFim:str, arg_dctQueueItem:str,arg_strCC:str=None):
        """
        Envio do e-mail em casos de erro durante a execução do robô.

        Parâmetros:
        - arg_strEnvioPara (str): destinatários separados por ';'.
        - arg_listAnexos (list): lista de anexos.
        - arg_strDetalhesErro (str): detalhes do erro.
        - arg_boolBusiness (bool): indica se o erro é de regra de negócio. (opcional, default=False)
        - arg_strCC (str): destinatários em cópia separados por ';'. (opcional, default=None)
        """    
        
        #Lendo template
        var_fileTemplate = open(ROOT_DIR + "\\resources\\templates\\Email_Erro.txt", "r")
        var_strEmailTexto = var_fileTemplate.read()
        var_fileTemplate.close()

        var_lstDescErro = arg_strDetalhesErro

        var_strEmailTexto = var_strEmailTexto.replace("*TIPO_ERRO*", var_lstDescErro[0])
        var_strEmailTexto = var_strEmailTexto.replace("*ERRO_DETALHES*", var_lstDescErro[1])
        var_strEmailTexto = var_strEmailTexto.replace("*ITEM ERRO*", arg_strItemErro)

        var_strEmailTexto = var_strEmailTexto.replace("*REQUISICAO*",arg_dctQueueItem)

        var_strEmailTexto = var_strEmailTexto.replace("*URL_ARQUIVO*",arg_strUrldrive)

        arg_strHorarioInicio = arg_strDataInit.replace(' ',' às ')
        arg_strHorarioFim = arg_strDataFim.replace(' ',' às ')

        var_strEmailTexto = var_strEmailTexto.replace("*DATAHORA_INI*", arg_strHorarioInicio).replace("*DATAHORA_FIM*", arg_strHorarioFim)

        var_strEmailAssunto = "[ERRO] PROCESSAMENTO RPA - DOWNLOAD ARQUIVO"
        
        var_listEnvioPara = arg_strEnvioPara.split(";") if(arg_strEnvioPara is not None) else []
        var_listCC = arg_strCC.split(";") if(arg_strCC is not None) else []

        #Configurando classe do email, separado para emitir logs diferentes
        try:
            self.var_clssMaestro.write_log("Configurando classe de email para envio")
            var_clssEmailController = BotEmailPlugin()
            var_clssEmailController.configure_smtp(self.var_strEmailServerSmtp, self.var_intEmailPortaSmtp)
            var_clssEmailController.login(self.var_strUsuario, self.var_strSenha)
            self.var_clssMaestro.write_log("Email configurado")

        except Exception:
            self.var_clssMaestro.write_log("Erro Configurando email:")
            self.var_clssMaestro.write_log(str(Exception))


        #Envia o email inicial
        try:
            self.var_clssMaestro.write_log("Enviando email de erro")
            var_clssEmailController.send_message(text_content=var_strEmailTexto, 
                                                 to_addrs=var_listEnvioPara, 
                                                 cc_addrs=var_listCC, 
                                                 use_html=True, 
                                                 subject=var_strEmailAssunto,
                                                 attachments=arg_listAnexos)
            
            self.var_clssMaestro.write_log("Email enviado com sucesso")

        except Exception:
            self.var_clssMaestro.write_log(arg_strMensagemLog="Erro enviando email de erro: " + str(Exception), arg_enumLogLevel=LogLevel.ERROR, arg_enumErrorType=ErrorType.APP_ERROR)

    def send_email(self, arg_strCorpoEmail:str, arg_strEnvioPara:str, arg_strCC:str, arg_strBCC:str, arg_strAssunto:str, arg_listAnexos:list=None, arg_boolHtml:bool=False):        
        """
        Simplesmente envia um email normal. Pode ser usado em vários lugares no código, porém é necessário informar um corpo para o email 

        Parâmetros:
        - arg_strCorpoEmail (str): corpo do e-mail.
        - arg_strEnvioPara (str): destinatários separados por ';'.
        - arg_strCC (str): destinatários em cópia separados por ';'.
        - arg_strBCC (str): destinatários em cópia oculta separados por ';'
        - arg_strAssunto (str): assunto do e-mail.
        - arg_listAnexos (list): lista de anexos. (opcional, default=None)
        - arg_boolHtml (bool): indica se o corpo do e-mail é HTML. (default=False)
        """
        var_listEnvioPara = arg_strEnvioPara.split(";") if(arg_strEnvioPara is not None) else []
        var_listCC = arg_strCC.split(";") if(arg_strCC is not None) else []
        var_listBCC = arg_strBCC.split(";") if(arg_strBCC is not None) else []

        #Configurando classe do email, separado para emitir logs diferentes
        try:
            self.var_clssMaestro.write_log("Configurando classe de email para envio")
            var_clssEmailController = BotEmailPlugin()
            var_clssEmailController.configure_smtp(self.var_strEmailServerSmtp, self.var_intEmailPortaSmtp)
            var_clssEmailController.login(self.var_strUsuario, self.var_strSenha)
            self.var_clssMaestro.write_log("Email configurado")
        except Exception:
            self.var_clssMaestro.write_log(arg_strMensagemLog="Erro Configurando email: " + str(Exception), arg_enumLogLevel=LogLevel.ERROR, arg_enumErrorType=ErrorType.APP_ERROR)
            raise

        #Envia o email customizado
        try:
            self.var_clssMaestro.write_log("Enviando email customizado")
            var_clssEmailController.send_message(text_content=arg_strCorpoEmail, 
                                                 to_addrs=var_listEnvioPara, 
                                                 attachments=arg_listAnexos, 
                                                 use_html=arg_boolHtml, 
                                                 subject=arg_strAssunto)
            self.var_clssMaestro.write_log("Email enviado com sucesso")
        except Exception:
            self.var_clssMaestro.write_log(arg_strMensagemLog="Erro enviando email customizado", arg_enumLogLevel=LogLevel.ERROR, arg_enumErrorType=ErrorType.APP_ERROR)
            raise
    
    def send_email_sucesso(self, arg_strHorarioInicio:str, arg_strHorarioFim:str, arg_strEnvioPara:str, arg_dctQueueItem:str, arg_strUrldrive:str,arg_strCC:str=None,arg_strTextoItem:str=None):
        """
        Envio do e-mail de finalização do robô.
        Recebendo o horário do início da execução, o horário final, para quem é necessário enviar (separado por ;) e os relatórios finais
        
        Parâmetros:
        - arg_strHorarioInicio (str): horário de início.
        - arg_strHorarioFim (str): horário de fim.
        - arg_strEnvioPara (str): destinatários separados por ';'.
        - arg_strRequisicao (str): Descrição da Requisição.
        - arg_strCC (str): destinatários em cópia separados por ';'. (opcional, default=None)
        """
        
        #Lendo template
        var_fileTemplate = open(ROOT_DIR + "\\resources\\templates\\Email_Sucesso.txt", "r")
        var_strEmailTexto = var_fileTemplate.read()
        var_fileTemplate.close()

        var_strEmailTexto = var_strEmailTexto.replace("*SOLICITACAO*", arg_strTextoItem)
        var_strEmailTexto = var_strEmailTexto.replace("*URL_ARQUIVO*", arg_strUrldrive)  
        arg_strHorarioInicio = arg_strHorarioInicio.replace(' ',' às ')
        arg_strHorarioFim = arg_strHorarioFim.replace(' ',' às ')
        var_strEmailTexto = var_strEmailTexto.replace("*DATAHORA_INI*", arg_strHorarioInicio).replace("*DATAHORA_FIM*", arg_strHorarioFim)
        var_strEmailAssunto = "PROCESSAMENTO RPA - DOWNLOAD ARQUIVOS"

        var_listEnvioPara = arg_strEnvioPara.split(";") if(arg_strEnvioPara is not None) else []
        var_listCC = arg_strCC.split(";") if(arg_strCC is not None) else []

        #Configurando classe do email, separado para emitir logs diferentes
        try:
            self.var_clssMaestro.write_log("Configurando classe de email para envio")
            var_clssEmailController = BotEmailPlugin()
            var_clssEmailController.configure_smtp(self.var_strEmailServerSmtp, self.var_intEmailPortaSmtp)
            var_clssEmailController.login(self.var_strUsuario, self.var_strSenha)
            self.var_clssMaestro.write_log("Email configurado")
        except Exception:
            self.var_clssMaestro.write_log(arg_strMensagemLog="Erro Configurando email: " + str(Exception), arg_enumLogLevel=LogLevel.ERROR, arg_enumErrorType=ErrorType.APP_ERROR)
            raise
        
        #Envia o email de finalização
        try:
            self.var_clssMaestro.write_log("Enviando email de finalização")
            var_clssEmailController.send_message(text_content=var_strEmailTexto, 
                                                 to_addrs=var_listEnvioPara, 
                                                 cc_addrs=var_listCC, 
                                                #  attachments = arg_strCaminhoExcel,
                                                 use_html=True, 
                                                 subject=var_strEmailAssunto)
            self.var_clssMaestro.write_log("Email enviado com sucesso")
        except Exception:
            self.var_clssMaestro.write_log(arg_strMensagemLog="Erro enviando email de finalização: " + str(Exception), arg_enumLogLevel=LogLevel.ERROR, arg_enumErrorType=ErrorType.APP_ERROR)
            raise


    # Lê os e-mails de Download Arquivos que estão no Inbox
    def read_email_Inbox(self, arg_clssT2CSql:T2CSqlAnaliticoSintetico) -> None:

        var_lstEmails = []

        # Inicializando conexão com o serviço IMAP
        var_mbMailBox = MailBox(host=self.var_strEmailServerImap, port=self.var_intEmailPortaImap)
        
        var_bmbBaseMailBox = var_mbMailBox.login(self.var_strUsuario, self.var_strSenha)

        var_bmbBaseMailBox.folder.set("INBOX")
        
        #COLOCAR ERRO CASO NÃO CONSIGA LOGAR NO EMAIL 

        # Lê os e-mails no Inbox que tenham  'Download de Arquivos' no assunto e Move para a pasta 'Onda 2'
        for email in var_bmbBaseMailBox.fetch(criteria=AND('SUBJECT ": Download de Arquivos| Tipo"', seen=False), mark_seen=False, bulk=True):
            var_bmbBaseMailBox.move([email.uid], "Onda 2 - Download de Arquivos") 
            var_LogMaestro = self.var_clssMaestro.write_log("Realizou a captura de e-mail marcado como não lido.",arg_enumLogLevel=LogLevel.INFO)
                    
        var_bmbBaseMailBox.folder.set("Onda 2 - Download de Arquivos")

        for email in var_bmbBaseMailBox.fetch(criteria=AND('SUBJECT ": Download de Arquivos| Tipo"', seen=False), mark_seen=False, bulk=True):

            var_strText = email.text.rsplit('|', 1)[0] + "|"

            try:
                self.var_clssMaestro.write_log(var_strText)

                var_lstEmails.append(regex.search(r"(?<=TIPO\:)\s+(.*?)(?=\|)", var_strText, flags=regex.IGNORECASE).group().strip())

                var_lstEmails.append(regex.search(r"(?<=PROJETO\:)\s+(.*?)(?=\|)", var_strText, flags=regex.IGNORECASE).group().strip().replace("/","-").replace("\\","-"))

                var_lstEmails.append(regex.search(r"(?<=EMAIL\:)\s+(.*?)(?=\|)", var_strText, flags=regex.IGNORECASE).group().strip())

                var_lstEmails.append(regex.search(r"(?<=CNPJ DO CLIENTE\:)\s+(.*?)(?=\|)", var_strText, flags=regex.IGNORECASE).group().strip().replace(".","").replace("/","").replace("-",""))

                var_lstEmails.append(regex.search(r"(?<=PROCURAÇÃO ELETRÔNICA\:)\s+(.*?)(?=\|)", var_strText, flags=regex.IGNORECASE).group().strip())

                var_strAuxDtInicio          = regex.search(r"(?<=PERÍODO INICIAL\:)\s+(.*?)(?=\|)", var_strText, flags=regex.IGNORECASE).group().strip().upper().replace("ST","").replace("ND","").replace("RD","").replace("TH","").replace("AUGU", "AUGUST")

                var_lstEmails.append(datetime.datetime.strftime( datetime.datetime.strptime(var_strAuxDtInicio, '%B %d %Y'),  '%d-%m-%Y'))

                var_strAuxDtFinal           = regex.search(r"(?<=PERÍODO FINAL\:)\s+(.*?)(?=\|)", var_strText, flags=regex.IGNORECASE).group().strip().upper().replace("ST","").replace("ND","").replace("RD","").replace("TH","").replace("AUGU", "AUGUST")

                var_lstEmails.append(datetime.datetime.strftime( datetime.datetime.strptime(var_strAuxDtFinal, '%B %d %Y'),  '%d-%m-%Y'))

                var_lstEmails.append(regex.search(r"(?<=EFD-CONTRIBUIÇÕES\:)\s+(.*?)(?=\|)", var_strText, flags=regex.IGNORECASE).group().strip().upper())

                var_lstEmails.append(regex.search(r"(?<=EFD\-ICMS\/IPI\:)\s+(.*?)(?=\|)", var_strText, flags=regex.IGNORECASE).group().strip().upper())
                
                var_lstEmails.append(regex.search(r"(?<=ECD\:)\s+(.*?)(?=\|)", var_strText, flags=regex.IGNORECASE).group().strip().upper())

                var_lstEmails.append(regex.search(r"(?<=ECF\:)\s+(.*?)(?=\|)", var_strText, flags=regex.IGNORECASE).group().strip().upper())

                var_lstEmails.append(regex.search(r"(?<=DCTF\:)\s+(.*?)(?=\|)", var_strText, flags=regex.IGNORECASE).group().strip().upper())

                var_lstEmails.append(regex.search(r"(?<=DARF\:)\s+(.*?)(?=\|)", var_strText, flags=regex.IGNORECASE).group().strip().upper())

                var_lstEmails.append(var_strText)

                var_strInfos = str(var_lstEmails).replace("[", "").replace("]", "")

                #Insere no banco Apter 
                arg_clssT2CSql.insert_controle_arquivos(arg_strInfosProjeto=var_strInfos)

                var_strInfos = None
                var_lstEmails = []

                var_bmbBaseMailBox.flag( email.uid, MailMessageFlags.SEEN,  True)

                
            except Exception as exception:
                self.var_clssMaestro.write_log("Erro ao inserir as informações da requisição -"+ var_strText +"- no Banco de Dados. \n" + str(exception),
                                            arg_enumLogLevel=LogLevel.ERROR)


        var_bmbBaseMailBox.logout()


        

