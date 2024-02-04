from botcity.web import WebBot, Browser
from botcity.core import DesktopBot
# Uncomment the line below for integrations with BotMaestro
# Using the Maestro SDK
from botcity.maestro import *
from prj_Apter_DownloadArquivos.classes_t2c.sqlserver.T2CSqlAnaliticoSintetico import T2CSqlAnaliticoSintetico
from prj_Apter_DownloadArquivos.classes_t2c.utils.T2CMaestro import T2CMaestro, LogLevel, ErrorType
from prj_Apter_DownloadArquivos.classes_t2c.T2CCloseAllApplications import T2CCloseAllApplications
from prj_Apter_DownloadArquivos.classes_t2c.T2CInitAllApplications import T2CInitAllApplications
from prj_Apter_DownloadArquivos.classes_t2c.utils.T2CExceptions import BusinessRuleException
from prj_Apter_DownloadArquivos.classes_t2c.utils.T2CScreenRecorder import T2CScreenRecorder
from prj_Apter_DownloadArquivos.classes_t2c.T2CKillAllProcesses import T2CKillAllProcesses
from prj_Apter_DownloadArquivos.classes_t2c.relatorios.T2CRelatorios import T2CRelatorios
from prj_Apter_DownloadArquivos.classes_t2c.T2CInitAllSettings import T2CInitAllSettings
from prj_Apter_DownloadArquivos.classes_t2c.sqlite.T2CSqliteQueue import T2CSqliteQueue
from prj_Apter_DownloadArquivos.classes_t2c.email.T2CSendEmail import T2CSendEmail
from prj_Apter_DownloadArquivos.classes_t2c.T2CProcess import T2CProcess
from webdriver_manager import chrome
import datetime

class Bot(WebBot):
    """
    Classe que utiliza as funcionalidades da classe WebBot.
    """
    def action(self, execution=None):
        """
        Método principal para execução do bot.

        Parâmetros:
        - execution (objeto): objeto de execução (opcional, default=None).

        Raises:
        - Exception: Caso nenhum bot seja fornecido.
        """

        # Configurar se deve ou não rodar no modo headless
        self.headless = False
        
        #Carregando o arquivo Config num dicionário
        var_dictConfig = T2CInitAllSettings().load_config()

        #Iniciando a classe para controle do maestro
        var_clssMaestro = T2CMaestro(arg_clssExecution=execution, arg_dictConfig=var_dictConfig)
        
        #RunnerId é o nome do runner se estiver rodando pelo maestro, mas se estiver rodando localmente, ele pega o nome da máquina
        var_strNomeMaquina:str = var_clssMaestro.var_strRunnerId

        #Iniciando variáveis de controle
        var_dateDatahoraInicio = datetime.datetime.now()
        var_strDatahoraInicio = var_dateDatahoraInicio.strftime("%d/%m/%Y %H:%M:%S")
        var_strNomeProcesso:str = var_dictConfig["NomeProcesso"]
        var_expTypeException = None
       
        #Variáveis contadores já levadas em conta no framework
        var_intQtdeItensProcessados = 0
        var_intQtdeItensAppException = 0
        var_intQtdeItensConsecutiveExceptions = 0
        var_intQtdeItensBusinessException = 0
        var_intQtdeItensSucesso = 0
        
        #Variáveis que precisam ser levadas em conta pelo desenvolvedor (somar contadores e indicar uso ou não)
        var_boolUsaCaptcha = False
        var_intQtdeCaptcha = 0

        # Instantiate a DesktopBot
        desktop_bot = DesktopBot()

        #CONFIGURAÇÕES DE BROWSER
        self.browser = Browser.CHROME

        #Descomente para baixar e usar a versão do webdriver do seu navegador escolhido
        self.driver_path = chrome.ChromeDriverManager().install()              #Chrome

        #Iniciando classes
        var_strCaminhoBd = var_dictConfig["CaminhoBancoSqlite"] if(var_dictConfig.__contains__("CaminhoBancoSqlite")) else None
        var_strTabelaFila = var_dictConfig["FilaProcessamento"] if(var_dictConfig.__contains__("FilaProcessamento")) else "tbl_Fila_Processamento"
        var_clssSqliteQueue = T2CSqliteQueue(arg_strCaminhoBd=var_strCaminhoBd, arg_strTabelaFila=var_strTabelaFila, arg_strNomeMaquina=var_strNomeMaquina, arg_clssMaestro=var_clssMaestro, arg_strMaxRetry= var_dictConfig["MaxRetryNumber"].__str__())
        var_clssSqlAnaliticoSintetico = T2CSqlAnaliticoSintetico(arg_clssMaestro=var_clssMaestro, arg_dictConfig=var_dictConfig)

        var_clssRelatorios = T2CRelatorios(arg_dictConfig=var_dictConfig)
        var_clssInitAllApplications = T2CInitAllApplications(arg_dictConfig=var_dictConfig, arg_botWebbot=self, arg_botDesktopbot=desktop_bot, arg_clssMaestro=var_clssMaestro, arg_clssSqliteQueue=var_clssSqliteQueue)
        var_clssCloseAllApplications = T2CCloseAllApplications(arg_dictConfig=var_dictConfig, arg_botWebbot=self, arg_botDesktopbot=desktop_bot, arg_clssMaestro=var_clssMaestro)
        var_clssKillAllProcesses = T2CKillAllProcesses(arg_dictConfig=var_dictConfig, arg_botWebbot=self, arg_botDesktopbot=desktop_bot, arg_clssMaestro=var_clssMaestro)
        var_clssProcess = T2CProcess(arg_dictConfig=var_dictConfig, arg_botWebbot=self, arg_botDesktopbot=desktop_bot, arg_clssMaestro=var_clssMaestro)
        var_clssScreenRecorder = T2CScreenRecorder(arg_strNomeProcesso=var_strNomeProcesso, arg_clssMaestro=var_clssMaestro, arg_dictConfig=var_dictConfig)

        #Se for indicando que é necessário incluir linhas no SQL Server, ou verificado que não está rodando em debug (vs code ou test task), insere a primeira aqui
        if(var_dictConfig["NecesSQLServer"].upper() == "SIM" and not var_clssMaestro.var_boolIsTestTask):
            var_clssSqlAnaliticoSintetico.insert_linha_inicio_sintetico(arg_strNomeMaquina=var_strNomeMaquina, 
                                                                        arg_dateInicioExecucao=var_dateDatahoraInicio)
            
        if(var_dictConfig["NecesSQLServer"].upper() == "SIM"):
            var_clssSqlAnaliticoSintetico.insert_linha_inicio_sintetico(arg_strNomeMaquina=var_strNomeMaquina, 
                                                                        arg_dateInicioExecucao=var_dateDatahoraInicio)
        #CLASSE COMENTADA POR PRECISAR DE PARÂMETROS CUSTOMIZADOS
        #DESCOMENTAR EM DESENVOLVIMENTO, apenas o que for necessário
        var_clssSendEmail = T2CSendEmail(arg_dictConfig= var_dictConfig,arg_strNomeProcesso=var_strNomeProcesso, arg_clssMaestro=var_clssMaestro)
        #var_clssSendEmailOutlook = T2CSendEmailOutlook(arg_strNomeProcesso=var_strNomeProcesso, arg_clssMaestro=var_clssMaestro)

        #Enviando email de inicialização
        if(var_dictConfig["EmailInicial"].upper() == "SIM"):
            var_clssSendEmail.send_email_inicial(arg_strEnvioPara=var_dictConfig["EmailDestinatarios"])
            #var_clssSendEmailOutlook.send_email_inicial(arg_strEnvioPara=var_dictConfig["EmailDestinatarios"])

        var_clssMaestro.write_log("Iniciando processamento: " + var_strNomeProcesso)

        #Se for indicado que é necessário Gravar a tela do processo, inicia aqui.
        if(var_dictConfig["GravarTela"].upper() == "SIM"):
            var_clssScreenRecorder.iniciar_gravacao()
        
        try:
            var_clssInitAllApplications.execute(arg_boolFirstRun=True,arg_clssT2CSql=var_clssSqlAnaliticoSintetico,arg_clssSendEmail=var_clssSendEmail)

            #Popula itens no banco interno para gerenciamento da fila 
            var_clssSqlAnaliticoSintetico.populate_process_queue(var_clssSqliteQueue)

            var_strQtdItensQueue = str( var_clssSqliteQueue.var_intItemsQueue )

            var_clssMaestro.write_log("Quantidade de Requisições Inseridas na fila: " + var_strQtdItensQueue )

        except BusinessRuleException as exception:
            #incluindo erro no relatório sintético, com horário de finalização
            var_dateDatahoraFim = datetime.datetime.now()
            var_strDatahoraFim = var_dateDatahoraFim.strftime("%d/%m/%Y %H:%M:%S")
            var_strTotalProcessamento = str(var_dateDatahoraFim - var_dateDatahoraInicio)
            var_strTotalProcessamento = var_strTotalProcessamento.split('.')[0]

            var_clssRelatorios.inserir_linha_sintetico(arg_listValores=[var_strNomeProcesso, 
                                                                        var_strDatahoraInicio, 
                                                                        var_strDatahoraFim, 
                                                                        var_strTotalProcessamento, 
                                                                        var_intQtdeItensProcessados, 
                                                                        var_intQtdeItensSucesso, 
                                                                        var_intQtdeItensBusinessException, 
                                                                        var_intQtdeItensAppException, 
                                                                        var_strNomeMaquina])

            #Tirando print do erro antes de fechar
            var_strCaminhoScreenshot = var_dictConfig["CaminhoExceptionScreenshots"] + "ExceptionScreenshot_" + datetime.datetime.now().strftime("%d%m%Y_%H%M%S%f") + ".png"
            desktop_bot.save_screenshot(path=var_strCaminhoScreenshot)
            
            #Se for indicando que é necessário incluir linhas no SQL Server, ou verificado que não está rodando em debug (vs code ou test task), insere a primeira aqui
            if(var_dictConfig["NecesSQLServer"].upper() == "SIM" and not var_clssMaestro.var_boolIsTestTask):
                var_clssSqlAnaliticoSintetico.update_linha_fim_sintetico(arg_dateFimExecucao=var_dateDatahoraFim, 
                                                                        #  arg_intQtdeCaptcha=var_intQtdeCaptcha,
                                                                         arg_intTotalItens=var_intQtdeItensProcessados, 
                                                                         arg_intTotalItensBusinessEx=var_intQtdeItensBusinessException,
                                                                         arg_intTotalItensAppEx=var_intQtdeItensAppException, 
                                                                         arg_intTotalItensSucesso=var_intQtdeItensSucesso)

            #Enviando o email final por causa do erro na inicialização
            #COMENTADA PARA EVITAR ERROS QUANDO CLASSE NÂO INICIALIZADA, DESCOMENTAR A QUE PRECISAR
            #if(var_dictConfig["EmailFinal"].upper() == "SIM"):
                #var_clssSendEmail.send_email_final(arg_strHorarioInicio=var_strDatahoraInicio, arg_strHorarioFim=var_strDatahoraFim, arg_strNomeProcesso=var_strNomeProcesso, arg_strEnvioPara=var_dictConfig["EmailDestinatarios"], arg_listAnexos=[var_clssRelatorios.var_strPathRelatorioAnalitico, var_clssRelatorios.var_strPathRelatorioSintetico], arg_boolSucesso=False)
                #var_clssSendEmailOutlook.send_email_final(arg_strHorarioInicio=var_strDatahoraInicio, arg_strHorarioFim=var_strDatahoraFim, arg_strEnvioPara=var_dictConfig["EmailDestinatarios"], arg_listAnexos=[var_clssRelatorios.var_strPathRelatorioAnalitico, var_clssRelatorios.var_strPathRelatorioSintetico, var_strCaminhoScreenshot], arg_boolSucesso=False)

            # Finaliza o Gravador de Tela
            if(var_dictConfig["GravarTela"].upper() == "SIM"):
                var_clssScreenRecorder.finalizar_gravacao()

            var_clssMaestro.finish_task(arg_boolSucesso=False, arg_strMensagem="Task finalizada na inicialização por erro de negócio, verifique os logs de execução.")
            raise
 
        except Exception as exception:
            #incluindo erro no relatório sintético, com horário de finalização
            var_dateDatahoraFim = datetime.datetime.now()
            var_strDatahoraFim = var_dateDatahoraFim.strftime("%d/%m/%Y %H:%M:%S")
            var_strTotalProcessamento = str(var_dateDatahoraFim - var_dateDatahoraInicio)
            var_strTotalProcessamento = var_strTotalProcessamento.split('.')[0]

            var_clssRelatorios.inserir_linha_sintetico(arg_listValores=[var_strNomeProcesso, 
                                                                        var_strDatahoraInicio, 
                                                                        var_strDatahoraFim, 
                                                                        var_strTotalProcessamento, 
                                                                        var_intQtdeItensProcessados, 
                                                                        var_intQtdeItensSucesso, 
                                                                        var_intQtdeItensBusinessException, 
                                                                        var_intQtdeItensAppException, 
                                                                        var_strNomeMaquina])

            #Tirando print do erro antes de fechar
            var_strCaminhoScreenshot = var_dictConfig["CaminhoExceptionScreenshots"] + "ExceptionScreenshot_" + datetime.datetime.now().strftime("%d%m%Y_%H%M%S%f") + ".png"
            desktop_bot.save_screenshot(path=var_strCaminhoScreenshot)
            
            #Se for indicando que é necessário incluir linhas no SQL Server, faz update com os dados finais aqui
            if(var_dictConfig["NecesSQLServer"].upper() == "SIM"):
                var_clssSqlAnaliticoSintetico.update_linha_fim_sintetico(arg_dateFimExecucao=var_dateDatahoraFim, 
                                                                        #  arg_intQtdeCaptcha=var_intQtdeCaptcha, 
                                                                         arg_intTotalItens=var_intQtdeItensProcessados, 
                                                                         arg_intTotalItensBusinessEx=var_intQtdeItensBusinessException, 
                                                                         arg_intTotalItensAppEx=var_intQtdeItensAppException, 
                                                                         arg_intTotalItensSucesso=var_intQtdeItensSucesso)


            #Enviando o email final por causa do erro na inicialização
            #COMENTADA PARA EVITAR ERROS QUANDO CLASSE NÂO INICIALIZADA, DESCOMENTAR A QUE PRECISAR
            #if(var_dictConfig["EmailFibnal"].upper() == "SIM"):
                #var_clssSendEmail.send_email_final(arg_strHorarioInicio=var_strDatahoraInicio, arg_strHorarioFim=var_strDatahoraFim, arg_strNomeProcesso=var_strNomeProcesso, arg_strEnvioPara=var_dictConfig["EmailDestinatarios"], arg_listAnexos=[var_clssRelatorios.var_strPathRelatorioAnalitico, var_clssRelatorios.var_strPathRelatorioSintetico], arg_boolSucesso=False)
                #var_clssSendEmailOutlook.send_email_final(arg_strHorarioInicio=var_strDatahoraInicio, arg_strHorarioFim=var_strDatahoraFim, arg_strEnvioPara=var_dictConfig["EmailDestinatarios"], arg_listAnexos=[var_clssRelatorios.var_strPathRelatorioAnalitico, var_clssRelatorios.var_strPathRelatorioSintetico, var_strCaminhoScreenshot], arg_boolSucesso=False)
            
            # Finaliza o Gravador de Tela
            if(var_dictConfig["GravarTela"].upper() == "SIM"):
                var_clssScreenRecorder.finalizar_gravacao()

            var_clssMaestro.finish_task(arg_boolSucesso=False, arg_strMensagem="Task finalizada na inicialização por erro desconhecido, verifique os logs de execução.")
            raise

        #--------------------------------- INICIO DO PROCESS ---------------------------------

        #Verifica se foi solicitado a interrupção da execução no Maestro
        var_boolInterrupted = var_clssMaestro.is_interrupted()
        if(var_boolInterrupted): var_dctQueueItem = dict()
        else: var_dctQueueItem = var_clssSqliteQueue.get_next_queue_item()

        #Processamento continua até a classe informar que não existe itens novos para processar
        if( bool(var_dctQueueItem) ): var_clssMaestro.write_log("Iniciando processamento dos itens encontrados na fila")
        else: var_clssMaestro.write_log("Não existem itens para processamento")

        while( bool(var_dctQueueItem) ):
            try:
                var_expTypeException = None
                var_intQtdeItensProcessados += 1
                var_dateDatahoraInicio_Item = datetime.datetime.now()
                var_strDatahoraInicio_Item = var_dateDatahoraInicio_Item.strftime("%d/%m/%Y %H:%M:%S")

                var_clssMaestro.write_log(arg_strMensagemLog= "Executando Item " + str(var_intQtdeItensProcessados) + " de " + var_strQtdItensQueue)

                var_clssMaestro.write_log(arg_strMensagemLog= "item Referência: " + str(var_dctQueueItem['referencia']) + 
                                          
                                          ". Tentativa: " + str(int(str(var_dctQueueItem['retry'])) + 1))

                #Executando o process
                # var_lstReturnProcess = var_clssProcess.execute(arg_dctQueueItem= var_dctQueueItem)
                var_strLinkDrive,var_strItemErro,var_strErro,var_lstPrintErro = var_clssProcess.execute(arg_dctQueueItem= var_dctQueueItem)
                var_intQtdeItensConsecutiveExceptions = 0

                if var_strErro != None:
                    var_lstDescErro = str(var_strErro).split(". Mensagem: ")


            except BusinessRuleException as exception:
                var_lstDescErro = str(exception).split(". Mensagem: ")
                var_expTypeException = exception

                var_intQtdeItensConsecutiveExceptions = 0
                var_intQtdeItensBusinessException += 1

                #Incluindo linha no relatório analítico sinalizando erro de aplicação
                var_dateDatahoraFim_Item = datetime.datetime.now()
                var_strDatahoraFim_Item = var_dateDatahoraFim_Item.strftime("%d/%m/%Y %H:%M:%S")

                var_clssRelatorios.inserir_linha_analitico(arg_listValores=[var_strDatahoraInicio_Item, 
                                                                            var_strDatahoraFim_Item, 
                                                                            var_dctQueueItem['id'], 
                                                                            var_dctQueueItem['referencia'], 
                                                                            var_strNomeMaquina, 
                                                                            "ERRO - REGRA DE NEGÓCIO", 
                                                                            var_lstDescErro[1]])
                
                #Se for indicando que é necessário incluir linhas no SQL Server, faz o insert com os dados do item aqui
                if(var_dictConfig["NecesSQLServer"].upper() == "SIM"):                    
                    var_clssSqlAnaliticoSintetico.insert_linha_analitico(arg_dctItemFila=var_dctQueueItem, 
                                                                        arg_strNomeProcesso=var_strNomeProcesso, 
                                                                        arg_dateInicioItem=var_dateDatahoraInicio_Item, 
                                                                        arg_dateFimItem=var_dateDatahoraFim_Item, 
                                                                        arg_strNomeFila=var_strTabelaFila, 
                                                                        arg_strStatusItem="ERRO", 
                                                                        arg_strTipoExcecao=var_lstDescErro[0], 
                                                                        arg_strDescricaoExcecao=var_lstDescErro[1])
                
                #Marcando item como erro de business
                var_clssSqliteQueue.update_status_item(arg_intIndex=int(var_dctQueueItem['id']), 
                                                       arg_strNovoStatus="BUSINESS ERROR", 
                                                       arg_strObs=var_lstDescErro[1])
                
                #Enviando email
                if(var_dictConfig["EmailCadaErro"].upper() == "SIM"):
                    var_clssSendEmail.send_email_erro(arg_boolBusiness=True, 
                                                      arg_strDetalhesErro=str(exception),
                                                      arg_strEnvioPara=var_dictConfig["EmailDestinatarios"])
                    
                    #var_clssSendEmailOutlook.send_email_erro(arg_boolBusiness=True, arg_strDetalhesErro=str(exception), arg_strEnvioPara=var_dictConfig["EmailDestinatarios"])             

            except Exception as exception:
                var_lstDescErro = exception.__str__().split(". Mensagem: ")
                var_expTypeException = exception
                var_intQtdeItensAppException +=1
                
                #Incluindo linha no relatório analítico sinalizando erro de aplicação
                var_dateDatahoraFim_Item = datetime.datetime.now()
                var_strDatahoraFim_Item = var_dateDatahoraFim_Item.strftime("%d/%m/%Y %H:%M:%S")
                var_clssRelatorios.inserir_linha_analitico(arg_listValores=[var_strDatahoraInicio_Item, 
                                                                            var_strDatahoraFim_Item, 
                                                                            var_dctQueueItem['id'], 
                                                                            var_dctQueueItem['referencia'], 
                                                                            var_strNomeMaquina, 
                                                                            "ERRO - APLICAÇÃO", 
                                                                            var_lstDescErro[0]])

                #Se for indicando que é necessário incluir linhas no SQL Server, faz o insert com os dados do item aqui
                if(var_dictConfig["NecesSQLServer"].upper() == "SIM"):                    
                    var_clssSqlAnaliticoSintetico.insert_linha_analitico(arg_dctItemFila=var_dctQueueItem, 
                                                                            arg_strNomeProcesso=var_strNomeProcesso, 
                                                                            arg_dateInicioItem=var_dateDatahoraInicio_Item, 
                                                                            arg_dateFimItem=var_dateDatahoraFim_Item, 
                                                                            arg_strNomeFila=var_strTabelaFila, 
                                                                            arg_strStatusItem="ERRO", 
                                                                            arg_strTipoExcecao=var_lstDescErro[0], 
                                                                            arg_strDescricaoExcecao=var_lstDescErro[1])

                #Coloca o item na fila novamente para reprocessamento. Ou marca o item como falha, caso atinja o máximo de tentativas
                var_clssSqliteQueue.retry_queue_item(arg_dctQueueItem=var_dctQueueItem, 
                                                     arg_strStatus= "APP ERROR", 
                                                     arg_strObs=var_lstDescErro[1])
                
                if( str(var_dctQueueItem['retry']) == str( var_dictConfig["MaxRetryNumber"] ) ):
                    pass
                    if(var_dictConfig["EmailCadaErro"].upper() == "SIM"):
                        var_clssSendEmail.send_email_erro(arg_boolBusiness=False, 
                                                          arg_strDetalhesErro=str(exception), 
                                                          arg_strEnvioPara=var_dictConfig["EmailDestinatarios"])
                        # var_clssSendEmailOutlook.send_email_erro(arg_boolBusiness=True, arg_strDetalhesErro=str(exception), arg_strEnvioPara=var_dictConfig["EmailDestinatarios"])             
               
            else:

                #Inluindo linha no relatório analítico sinalizando sucesso
                var_intQtdeItensSucesso += 1
                var_dateDatahoraFim_Item = datetime.datetime.now()
                var_strDatahoraFim_Item = var_dateDatahoraFim_Item.strftime("%d/%m/%Y %H:%M:%S")
                var_clssRelatorios.inserir_linha_analitico(arg_listValores=[var_strDatahoraInicio_Item, 
                                                                            var_strDatahoraFim_Item, 
                                                                            var_dctQueueItem['id'], 
                                                                            var_dctQueueItem['referencia'], 
                                                                            var_strNomeMaquina, 
                                                                            "SUCESSO", 
                                                                            ""])
                
                #Se for indicando que é necessário incluir linhas no SQL Server, faz o insert com os dados do item aqui
                if(var_dictConfig["NecesSQLServer"].upper() == "SIM"): 
                    var_clssSqlAnaliticoSintetico.insert_linha_analitico(arg_dctItemFila=var_dctQueueItem, 
                                                                        arg_strNomeProcesso=var_strNomeProcesso, 
                                                                        arg_dateInicioItem=var_dateDatahoraInicio_Item, 
                                                                        arg_dateFimItem=var_dateDatahoraFim_Item, 
                                                                        arg_strNomeFila=var_strTabelaFila, 
                                                                        arg_strStatusItem="SUCESSO")         

                #Marcando item como finalizado com sucesso
                var_clssSqliteQueue.update_status_item(arg_intIndex=int(var_dctQueueItem['id']), 
                                                       arg_strNovoStatus="SUCCESS")
                
                #Enviando email
                #if(var_dictConfig["EmailCadaErro"].upper() == "SIM"):
                    #var_clssSendEmail.send_email_erro(arg_boolBusiness=True, arg_strDetalhesErro=str(exception), arg_strEnvioPara=var_dictConfig["EmailDestinatarios"])
                    #var_clssSendEmailOutlook.send_email_erro(arg_boolBusiness=True, arg_strDetalhesErro=str(exception), arg_strEnvioPara=var_dictConfig["EmailDestinatarios"])             


                if(var_dictConfig['EmailFinal'].upper() == "SIM"):

                    var_dateDatahoraFim_Item = datetime.datetime.now()
                    var_strDatahoraFim_Item = var_dateDatahoraFim_Item.strftime("%d/%m/%Y %H:%M:%S")

                    if var_strItemErro == "":
                        var_clssSendEmail.send_email_sucesso(arg_strEnvioPara=var_dctQueueItem["email"], 
                                                            arg_strHorarioInicio=var_strDatahoraInicio, 
                                                            arg_strHorarioFim=var_strDatahoraFim_Item, 
                                                            arg_dctQueueItem = var_dctQueueItem["id_download"], 
                                                            arg_strUrldrive=var_strLinkDrive,
                                                            arg_strTextoItem = var_dctQueueItem["tarefa_infos"],
                                                            arg_strCC = var_dictConfig["EmailDestinatariosSucesso"])
                    else:
                        var_clssSendEmail.send_email_erro(arg_strDetalhesErro=var_lstDescErro, 
                                                        arg_strItemErro = var_strItemErro,
                                                        arg_strEnvioPara=var_dictConfig["EmailDestinatarios"], 
                                                        arg_strCC = var_dctQueueItem["email"],
                                                        arg_strDataInit=var_strDatahoraInicio, 
                                                        arg_strDataFim=var_strDatahoraFim_Item,
                                                        arg_listAnexos = var_lstPrintErro,
                                                        arg_strUrldrive=var_strLinkDrive,
                                                        arg_dctQueueItem = var_dctQueueItem["tarefa_infos"])

            finally:
                #Verifica se o item processado saiu do loop for com Exceção
                if type(var_expTypeException) == Exception:
                    #Incrementa +1 devido a exceção do Item do Processado
                    var_intQtdeItensConsecutiveExceptions += 1
                    if(var_intQtdeItensConsecutiveExceptions >= var_dictConfig['MaxConsecutiveSystemExceptions']): break
                
                #Verifica se foi solicitado a interrupção da execução no Maestro
                var_boolInterrupted = var_clssMaestro.is_interrupted()
                if(var_boolInterrupted == True): var_dctQueueItem = dict()
                #Processamento continua até a classe informar que não existe itens novos para processar
                else: var_dctQueueItem = var_clssSqliteQueue.get_next_queue_item()

        #--------------------------------- FIM DO PROCESS ---------------------------------

        #--------------------------------- INICIO DO END ---------------------------------

        #Fechando aplicativos no final do processamento
        try:
            var_clssCloseAllApplications.execute()
        except Exception:
            var_clssMaestro.write_log(arg_strMensagemLog="Fechando aplicativos pelo KillAllProcesses", arg_enumLogLevel=LogLevel.WARN)
            var_clssKillAllProcesses.execute()

        #Preenchendo linha no relatório sintético, incluindo horário de finalização
        var_clssMaestro.write_log("Inserindo linhas no relatório analítico e enviando email")
        var_dateDatahoraFim = datetime.datetime.now()
        var_strDatahoraFim = var_dateDatahoraFim.strftime("%d/%m/%Y %H:%M:%S")
        var_strTotalProcessamento = str(var_dateDatahoraFim - var_dateDatahoraInicio)
        var_strTotalProcessamento = var_strTotalProcessamento.split('.')[0]

        var_clssRelatorios.inserir_linha_sintetico(arg_listValores=[var_strNomeProcesso, 
                                                                    var_strDatahoraInicio, 
                                                                    var_strDatahoraFim, 
                                                                    var_strTotalProcessamento, 
                                                                    var_intQtdeItensProcessados, 
                                                                    var_intQtdeItensSucesso, 
                                                                    var_intQtdeItensBusinessException, 
                                                                    var_intQtdeItensAppException, 
                                                                    var_strNomeMaquina])

        #Se for indicando que é necessário incluir linhas no SQL Server, ou verificado que não está rodando em debug (vs code ou test task), insere a primeira aqui
        if(var_dictConfig["NecesSQLServer"].upper() == "SIM" and not var_clssMaestro.var_boolIsTestTask):
                var_clssSqlAnaliticoSintetico.update_linha_fim_sintetico(arg_dateFimExecucao=var_dateDatahoraFim, 
                                                                        #  arg_intQtdeCaptcha=var_intQtdeCaptcha, 
                                                                         arg_intTotalItens=var_intQtdeItensProcessados, 
                                                                         arg_intTotalItensBusinessEx=var_intQtdeItensBusinessException, 
                                                                         arg_intTotalItensAppEx=var_intQtdeItensAppException, 
                                                                         arg_intTotalItensSucesso=var_intQtdeItensSucesso)



        #Enviando email final com os relatórios analítico e sintético
        # if(var_dictConfig["EmailFinal"].upper() == "SIM"):   
        #     var_clssSendEmail.send_email_final(arg_strHorarioInicio=var_strDatahoraInicio, arg_strHorarioFim=var_strDatahoraFim, arg_strEnvioPara=var_dictConfig["EmailDestinatarios"], arg_listAnexos=[var_clssRelatorios.var_strPathRelatorioAnalitico, var_clssRelatorios.var_strPathRelatorioSintetico], arg_boolSucesso=True)
            #var_clssSendEmailOutlook.send_email_final(arg_strHorarioInicio=var_strDatahoraInicio, arg_strHorarioFim=var_strDatahoraFim, arg_strEnvioPara=var_dictConfig["EmailDestinatarios"], arg_listAnexos=[var_clssRelatorios.var_strPathRelatorioAnalitico, var_clssRelatorios.var_strPathRelatorioSintetico], arg_boolSucesso=True)

        # Finaliza o Gravador de Tela
        if(var_dictConfig["GravarTela"].upper() == "SIM"):
            var_clssScreenRecorder.finalizar_gravacao()
        
        var_clssMaestro.write_log("Finalizando processamento.")
        if not type(var_expTypeException) == Exception:
            var_clssMaestro.finish_task(arg_boolSucesso=True, arg_strMensagem="Task finished OK.")
        else: 
            var_clssMaestro.finish_task(arg_boolSucesso=False, arg_strMensagem="Task failed.")        
            

        #--------------------------------- FIM DO END ---------------------------------

if __name__ == '__main__':
    Bot.main()
