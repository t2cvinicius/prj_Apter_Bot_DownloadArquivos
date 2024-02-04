from prj_Apter_DownloadArquivos.classes_t2c.sqlserver.T2CSqlAnaliticoSintetico import T2CSqlAnaliticoSintetico
from prj_Apter_DownloadArquivos.classes_t2c.utils.T2CMaestro import T2CMaestro, LogLevel, ErrorType
from prj_Apter_DownloadArquivos.classes_t2c.utils.T2CExceptions import BusinessRuleException
from prj_Apter_DownloadArquivos.classes_t2c.receitanet.ReceitaNet import ReceitaNet
from prj_Apter_DownloadArquivos.classes_t2c.utils.Planilhamento import Planilhamento
from prj_Apter_DownloadArquivos.classes_t2c.utils.FilesTreatment import Files
from prj_Apter_DownloadArquivos.classes_t2c.utils.GoogleDrive import GoogleDrive
from botcity.web import WebBot, Browser
from botcity.core import DesktopBot

import traceback
import datetime
import shutil
import time
import os 

#Classe responsável pelo processamento principal, necessário preencher com o seu código no método execute
class T2CProcess:
    """
    Classe responsável pelo processamento principal.
    """
    #Iniciando a classe, pedindo um dicionário config e o bot que vai ser usado e enviando uma exceção caso nenhum for informado
    def __init__(self, arg_dictConfig:dict, arg_clssMaestro:T2CMaestro, arg_botWebbot:WebBot=None, arg_botDesktopbot:DesktopBot=None):
        """
        Inicializa a classe T2CProcess.

        Parâmetros:
        - arg_dictConfig (dict): dicionário de configuração.
        - arg_clssMaestro (T2CMaestro): instância de T2CMaestro.
        - arg_botWebbot (WebBot): instância de WebBot (opcional, default=None)
        - arg_botDesktopbot (DesktopBot): instância de DesktopBot (opcional, default=None)

        Raises:
        - Exception: caso algum bot não seja fornecido.
        """
        if(arg_botWebbot is None or arg_botDesktopbot is None): raise Exception("Não foi possível inicializar a classe")
        else:
            self.var_botWebbot = arg_botWebbot
            self.var_botDesktopbot = arg_botDesktopbot
            self.var_dictConfig = arg_dictConfig
            self.var_clssMaestro = arg_clssMaestro
            
    #Parte principal do código, deve ser preenchida pelo desenvolvedor
    def execute(self, arg_dctQueueItem:dict ):
        """
        Método principal para execução do código.

        Parâmetros:
        - arg_dctQueueItem (dict): NamedTuple com o item da fila a ser processado.

        """

        var_strProjectName = str(arg_dctQueueItem['cnpj'])[0:8] + "_" + str(arg_dctQueueItem['tipo_projeto']).replace(" ", "_") + "_" + datetime.datetime.now().strftime("%d%m%Y_%H%M%S")
        var_strPathProject = self.var_dictConfig['path_base_arquivos'] + var_strProjectName  + "\\"
        var_strPathExcel = var_strPathProject + var_strProjectName +"_Book de Dados.xlsx"

        var_dateInicio = datetime.datetime.strptime(arg_dctQueueItem['data_inicio'],'%d-%m-%Y')
        var_dateFim  = datetime.datetime.strptime(arg_dctQueueItem['data_fim'], '%d-%m-%Y')
        
        var_lstPrintErro = []
        var_strItemErro = ""
        var_strErro = None
        var_strException = None

        try:
            var_clssReceitaNet = ReceitaNet( arg_clssMaestro=self.var_clssMaestro,arg_botDesktopBot=self.var_botDesktopbot, arg_dctConfig=self.var_dictConfig, arg_dctQueue=arg_dctQueueItem)
            var_clssPlanilhamento = Planilhamento(arg_clssMaestro=self.var_clssMaestro,arg_strPathExcel=var_strPathExcel,arg_strPathDownload=self.var_dictConfig ["path_download_receitanet"] ,arg_strPathProject= var_strPathProject, arg_dctQueue= arg_dctQueueItem)
            var_clssT2CSql = T2CSqlAnaliticoSintetico(arg_clssMaestro=self.var_clssMaestro, arg_dictConfig= self.var_dictConfig)
            var_clssGoogleDrive = GoogleDrive(arg_clssMaestro=self.var_clssMaestro,arg_dictConfig=self.var_dictConfig)

            # if( arg_dctQueueItem['email'] == '' or not str(arg_dctQueueItem['email']).__contains__("@apter.com.br") ):
            if arg_dctQueueItem['email'] == '' or '@apter.com.br' not in str(arg_dctQueueItem['email']).lower():
                raise BusinessRuleException("ERRO. PROCESS. Mensagem: O campo que se refere ao Email Gerente, não foi informado ou está fora do padrão.")
            

            Files().create_Folder(arg_strPath= var_strPathProject)
            
            origem = r"C:\\robo\BotCity\\Projetos\\prj_Apter_Bot_DownloadArquivos\\prj_Apter_DownloadArquivos\\resources\\templates\\BOOK DE DADOS_TEMPLATE.xlsx"
            destino = var_strPathProject
            shutil.copy(origem,destino)
            os.rename(os.path.join(var_strPathProject, "BOOK DE DADOS_TEMPLATE.xlsx"), os.path.join(var_strPathProject, var_strProjectName +"_Book de Dados.xlsx"))
            var_strPathExcel = var_strPathProject + var_strProjectName +"_Book de Dados.xlsx"

            #Cartao CNPJ
            # var_LogMaestro = self.var_clssMaestro.write_log("Iniciando Download Cartão CNPJ",arg_enumLogLevel=LogLevel.INFO)
            """
            Codigo do Download CNPJ 
            """
            
            var_lstAuxFilesReceita = ['efd_contribuicoes', 'efd_fiscal', 'ecd', 'ecf']

            var_lstFilesReceitaNet:list[str] = [key for key, values in arg_dctQueueItem.items() if str(values).upper() == 'SIM' and key in var_lstAuxFilesReceita] #Pega todas as linhas que contem Sim para saber oque deve processar#

            if(arg_dctQueueItem['dctf'].upper() == 'SIM'):
                try:
                    var_strTipo = 'dctf'

                    var_lstTable = list[dict]

                    var_lstLinesToDownload = var_clssT2CSql.get_lines_to_remove_dctf(var_lstTable=var_lstTable, arg_dateDtInicio=var_dateInicio, arg_dateDtFim=var_dateFim)
                    
                except BusinessRuleException as exception:
                    pass

                except Exception as exception:
                    pass


            if( arg_dctQueueItem['darf'].upper() == 'SIM'):
                try:
                    var_strTipo = 'darf'                  
                
                except BusinessRuleException as exception:
                    pass

                except Exception as exception:
                    pass

            if( bool(var_lstFilesReceitaNet) ):
                for tipo in var_lstFilesReceitaNet:
                    try:

                        if( tipo == 'efd_contribuicoes'): 
                            var_strFolderName = 'EFD Contribuições'

                        elif( tipo == 'efd_fiscal'): 
                            var_strFolderName = 'EFD Fiscal'
                        
                        else: 
                            var_strFolderName = tipo.upper()
                        
                        var_strPathTipoArquivo = var_strPathProject + var_strFolderName + "\\"

                        Files().create_Folder(arg_strPath= var_strPathTipoArquivo)
                        
                        var_LogMaestro = self.var_clssMaestro.write_log(f"Iniciando o processamento {tipo}.",arg_enumLogLevel=LogLevel.INFO)

                        var_clssReceitaNet.open_receitanet()

                        var_clssReceitaNet.verify_download_queue()

                        var_clssReceitaNet.fill_filters(arg_strTipo= tipo)

                        var_strPathTable = var_clssReceitaNet.extract_table()

                        var_dfTable = var_clssPlanilhamento.read_table_receita(arg_strPathTable= var_strPathTable, arg_strTipo= tipo)
                        
                        # if( not tipo == "efd_fiscal" ):
                        # var_lstLinesToRemove = list( var_dfTable.loc[ var_dfTable['remove_dpt'] | var_dfTable['remove_out'] ]['index'] )
                        var_lstLinesToRemove = list( var_dfTable.loc[ var_dfTable['remove_dpt']]['index'] )
                        var_lstLinesToRemove.sort()
                        var_clssReceitaNet.deselect_files(var_lstLinesToRemove)

                        var_clssReceitaNet.download_arquivos()

                        var_clssReceitaNet.close_receitanet()

                        var_clssPlanilhamento.planilhamento_receitanet(arg_strTipo= tipo, arg_dfTable= var_dfTable, arg_strPathTipo=var_strPathTipoArquivo)
            
                        var_clssT2CSql.update_status_arquivo(arg_strTipo = tipo, arg_strIdDownload = arg_dctQueueItem['id_download'],arg_strStatus = 'FEITO')

                    #Exception mapeado erro de negocio (NÃO PARA O BOT, SOMENTE REGISTRA)   
                    except BusinessRuleException as exception:
                        var_strItemErro += f"{tipo.upper()} | "
                        self.var_clssMaestro.write_log(arg_strMensagemLog=str(exception),
                            arg_enumErrorType=ErrorType.BUSINESS_ERROR,
                            arg_enumLogLevel=LogLevel.ERROR)


                    #Exception para o processo (Erro em caso de seletor/reframe)       
                    except Exception as exception:
                        var_strItemErro += f"{tipo.upper()} | "

                        if exception.__str__().__contains__(". Mensagem: "):
                            var_strErro = exception.__str__()

                        elif exception.__str__().__contains__("Element not found: "):
                            var_strErro = "ERRO SELETOR. RECEITANET. Mensagem: " + traceback.format_exc()
                            
                        else:
                            var_strErro = "ERRO WORKFLOW. RECEITANET. Mensagem: Erro não mapeado durante execução do ReceitaNetBX.\n" + traceback.format_exc()
                    
                        self.var_clssMaestro.write_log(arg_strMensagemLog=str(var_strErro),
                            arg_enumErrorType=ErrorType.APP_ERROR,
                            arg_enumLogLevel=LogLevel.ERROR)
                        
                    finally:

                        if tipo.lower() in var_strItemErro.lower():

                            var_strPathErroArquivo = var_strPathProject + "ERRO" + "\\"
                            Files().create_Folder(arg_strPath= var_strPathErroArquivo)

                            var_CaminhoPrint = (var_strPathErroArquivo + tipo.upper()  + "_" +datetime.datetime.now().strftime("%d%m%Y_%H%M%S")+".png")
                            self.var_botDesktopbot.screenshot(var_CaminhoPrint)
                            var_lstPrintErro.append(var_CaminhoPrint)
                            var_CaminhoPrint = ""

                        var_clssReceitaNet.close_receitanet()

        except BusinessRuleException as exception:

            var_clssReceitaNet.close_receitanet()
            self.var_clssMaestro.write_log(arg_strMensagemLog=str(exception),
                                           arg_enumErrorType=ErrorType.BUSINESS_ERROR,
                                           arg_enumLogLevel=LogLevel.ERROR)

            #Para o processo
            raise BusinessRuleException( str(exception) )


        except Exception as exception:
            var_clssReceitaNet.close_receitanet()

            if( str(exception).__contains__(". Mensagem: ") ):
                
                var_strException = str(exception)
            else:
                var_strException = "ERRO WORKFLOW. Mensagem: Ocorreu um erro inesperado durante o processamento. \n" +  traceback.format_exc()

            self.var_clssMaestro.write_log(arg_strMensagemLog=var_strException,
                                           arg_enumErrorType=ErrorType.APP_ERROR,
                                           arg_enumLogLevel=LogLevel.ERROR)

            raise Exception(var_strException)

        else: 
            if var_strErro or var_strException != None:            
                var_clssT2CSql.update_status_arquivo(arg_strTipo = 'status_projeto', arg_strIdDownload = arg_dctQueueItem['id_download'],arg_strStatus = 'ERRO')
            else:
                var_clssT2CSql.update_status_arquivo(arg_strTipo = 'status_projeto', arg_strIdDownload = arg_dctQueueItem['id_download'],arg_strStatus = 'FEITO')

            self.var_clssMaestro.write_log(arg_strMensagemLog="Download de Arquivos finalizado.")
            var_strLinkDrive = var_clssGoogleDrive.get_link(arg_strPathName = var_strProjectName)       
            return var_strLinkDrive,var_strItemErro,var_strErro,var_lstPrintErro

        
