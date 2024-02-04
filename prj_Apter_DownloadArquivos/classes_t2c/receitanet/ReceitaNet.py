from prj_Apter_DownloadArquivos.classes_t2c.utils.T2CMaestro import T2CMaestro, LogLevel, ErrorType
from prj_Apter_DownloadArquivos.classes_t2c.utils.T2CExceptions import BusinessRuleException, ElementNotFound
from prj_Apter_DownloadArquivos.classes_t2c.utils.FilesTreatment import Files

from pywinauto.application import Application
from botcity.core import DesktopBot
from pathlib import Path
import subprocess
import unidecode 
import time


ROOT_DIR = Path(__file__).parent.parent.parent.__str__() + "\\"


class ReceitaNet:
    def __init__(self, arg_clssMaestro:T2CMaestro,arg_botDesktopBot:DesktopBot, arg_dctConfig:dict, arg_dctQueue:dict) -> None:
        self.bot = arg_botDesktopBot
        self.var_clssMaestro = arg_clssMaestro

        self.var_strPathRobotExe = arg_dctConfig['path_robotexe']
        self.var_strPathReceitaApp = arg_dctConfig['path_app_receitanet']
        self.var_strPathDownloadReceitanet = arg_dctConfig['path_download_receitanet']
        self.var_strCNPJ = str( arg_dctQueue['cnpj'] )
        self.var_strDataInicio = str(arg_dctQueue['data_inicio']).replace("-", "")
        self.var_strDataFim = str(arg_dctQueue['data_fim']).replace("-", "")
        self.var_strRoot_dir = str(arg_dctConfig['Root_dir'])

        
        var_strPathCertificados = arg_dctConfig['path_certificados']
        var_strCertificado = unidecode.unidecode( str(arg_dctQueue['procuracao']) ).upper().replace(":", "!")

        self.var_strCertificado = Files().get_Files_By_ID(arg_strPath= var_strPathCertificados, arg_strIdentificador= var_strCertificado)[0]
        self.var_strPassCert = self.var_strCertificado.split("-")[1].split(".")[0].strip()


    def open_receitanet(self) -> None:

        # self.close_receitanet()
        var_LogMaestro = self.var_clssMaestro.write_log("Aguardando iniciar ReceitaNetBx",arg_enumLogLevel=LogLevel.INFO)

        self.bot.execute(self.var_strPathReceitaApp)
        
        cont = 0
        var_aux = False
        while cont < 10 and var_aux == False :
            time.sleep(5)
            if self.bot.find( "Btn_Certificado", matching=0.97, waiting_time=300):
                var_LogMaestro = self.var_clssMaestro.write_log("ReceitaNetBx iniciado.",arg_enumLogLevel=LogLevel.INFO)
                var_aux = True
        cont += 1

        try: 
            if not self.bot.find( "Btn_Certificado", matching=0.97, waiting_time=300):
                self.__not_found("Btn_Certificado")
            time.sleep(1)
            self.bot.click(wait_after=500)        
            var_LogMaestro = self.var_clssMaestro.write_log("Click em 'Buscar Certificada'.",arg_enumLogLevel=LogLevel.INFO)


            if not self.bot.find( "Btn_Escolha", matching=0.97, waiting_time=5000):
                self.__not_found("Btn_Escolha")
            self.bot.click_relative(-198, 5)
            self.bot.paste(self.var_strCertificado)
            var_LogMaestro = self.var_clssMaestro.write_log("Passou o caminho do arquivo.",arg_enumLogLevel=LogLevel.INFO)


            if not self.bot.find( "Btn_Escolha", matching=0.97, waiting_time=5000):
                self.__not_found("Btn_Escolha")
            self.bot.click(wait_after=500)
            var_LogMaestro = self.var_clssMaestro.write_log("Click em 'Escolha'.",arg_enumLogLevel=LogLevel.INFO)


            if not self.bot.find( "Btn_OK", matching=0.97, waiting_time=5000):
                self.__not_found("Btn_OK")
            self.bot.double_click_relative(27, -27)
            self.bot.paste(self.var_strPassCert)
            var_LogMaestro = self.var_clssMaestro.write_log("Informou a senha do certificado digital.",arg_enumLogLevel=LogLevel.INFO)


            if not self.bot.find( "Btn_OK", matching=0.97, waiting_time=5000):
                self.__not_found("Btn_OK")
            self.bot.click(wait_after=500)

            if self.bot.find( "img_Erro_Pesquisa_Filtro", matching=0.97, waiting_time=2000):
                var_LogMaestro = self.var_clssMaestro.write_log("Erro ao validar Certificado.",arg_enumLogLevel=LogLevel.INFO)
                raise Exception("Erro ao Validar Certificado.")

            self.__set_focus()

            if not self.bot.find( "Slct_Contribuinte", matching=0.97, waiting_time=5000):
                self.__not_found("Slct_Contribuinte")
            self.bot.click(wait_after=500)
            self.bot.click_relative(31, 39, wait_after=500)
            var_LogMaestro = self.var_clssMaestro.write_log("Selecionado perfil de acesso.",arg_enumLogLevel=LogLevel.INFO)


            if not self.bot.find( "Slct_CPF", matching=0.97, waiting_time=5000):
                self.__not_found("Slct_CPF")
            self.bot.click(wait_after=500)
            self.bot.click_relative(12, 39, wait_after=500)
            self.bot.double_click_relative(87, 6)
            self.bot.type_keys(self.var_strCNPJ)
            var_LogMaestro = self.var_clssMaestro.write_log("Alterou de CPF para CNPJ e realizou o preenchimento.",arg_enumLogLevel=LogLevel.INFO)
            
            if not self.bot.find_text( "Btn_Entrar", waiting_time=5000):
                self.__not_found("Btn_Entrar")
            self.bot.double_click()
            var_LogMaestro = self.var_clssMaestro.write_log("Click em 'Entrar'",arg_enumLogLevel=LogLevel.INFO)
            
            #Logica de erro                        
            if self.bot.find( "CNPJ_Invalido", matching=0.97, waiting_time=10000) != None:
                raise Exception("O CNPJ informado é invalido.")

            if self.bot.find( "img_Procurador_Receitanet", matching=0.97, waiting_time=15000):
                var_LogMaestro = self.var_clssMaestro.write_log("Abertura realizada com sucesso",arg_enumLogLevel=LogLevel.INFO)

        
        except Exception as exception:
            raise BusinessRuleException("ERRO APP. RECEITANET. Mensagem: Erro ao inicializar o Aplicativo Receitanet: \n" + str(exception))


    def close_receitanet(self) -> None:

        try:

            self.__set_focus()
            var_LogMaestro = self.var_clssMaestro.write_log("Fechando ReceitaNetBx",arg_enumLogLevel=LogLevel.INFO)
            
            if not self.bot.find( "img_Receitanet_Login", matching=0.97, waiting_time=3000):
                if self.bot.find( "Btn_Sair_Receitanet", matching=0.97, waiting_time=3000):
                    self.bot.click()
                
            else:                
                if self.bot.find( "Btn_Sair_Login", matching=0.97, waiting_time=3000):
                    self.bot.click()

            self.app.kill()
            
        except:
            var_LogMaestro = self.var_clssMaestro.write_log("Receitanet não está Aberto.",arg_enumLogLevel=LogLevel.INFO)


    def verify_download_queue(self) -> None:

        self.__set_focus()
        
        var_LogMaestro = self.var_clssMaestro.write_log("Verificando Fila de Downloads.",arg_enumLogLevel=LogLevel.INFO)
        
        try:
            if not self.bot.find( "Btn_Acompanhamento", matching=0.97, waiting_time=5000):
                self.__not_found("Btn_Acompanhamento")
            self.bot.click()
            var_LogMaestro = self.var_clssMaestro.write_log("Click em 'Acompanhamento'.",arg_enumLogLevel=LogLevel.INFO)
            
            if not self.bot.find( "Btn_Fila_Download", matching=0.97, waiting_time=5000):
                self.__not_found("Btn_Fila_Download")
            self.bot.click()
            var_LogMaestro = self.var_clssMaestro.write_log("Click em 'Fila de download'",arg_enumLogLevel=LogLevel.INFO)

            if not self.bot.find( "img_Fila_Vazia", matching=0.97, waiting_time=5000):
                var_LogMaestro = self.var_clssMaestro.write_log("Arquivos encontrados para Download.",arg_enumLogLevel=LogLevel.INFO)

                # Notifica por e-mail
                self.__wait_fila_download()
                var_LogMaestro = self.var_clssMaestro.write_log("Fila de Download Finalizada com sucesso.",arg_enumLogLevel=LogLevel.INFO)

            else:
                var_LogMaestro = self.var_clssMaestro.write_log("Fila de Donwloads já se encontra vazia.",arg_enumLogLevel=LogLevel.INFO)

        except Exception as exception:
            raise Exception("ERRO APP. RECEITANET. Mensagem: Erro ao realizar Download dos Arquivos  : \n" + str(exception))

        var_LogMaestro = self.var_clssMaestro.write_log("Limpando Diretório Padrão de Downloads do ReceitaNet.",arg_enumLogLevel=LogLevel.INFO)

        Files().clear_Folder(arg_strPath=self.var_strPathDownloadReceitanet)


    def fill_filters(self, arg_strTipo:str) -> None:

        self.__set_focus()

        if arg_strTipo == "efd_contribuicoes":
            var_strSeletorSistema = "Slct_Sistema_EFD_Contribuicoes"

        elif arg_strTipo == "efd_fiscal":
            var_strSeletorSistema = "Slct_Sistema_EFD_Fiscal"

        elif arg_strTipo == "ecd":
            var_strSeletorSistema = "Slct_Sistema_ECD"

        elif arg_strTipo == "ecf":  
            var_strSeletorSistema = "Slct_Sistema_ECF"
        

        try:
            if not self.bot.find( "Btn_Pesquisar_Receita", matching=0.97, waiting_time=5000):
                self.__not_found("Btn_Pesquisar_Receita")
            self.bot.click(wait_after=1000)
            var_LogMaestro = self.var_clssMaestro.write_log("Click em 'Pesquisa'.",arg_enumLogLevel=LogLevel.INFO)
            
            if not self.bot.find( "Slct_Filtro_Sistema", matching=0.97, waiting_time=5000):
                self.__not_found("Slct_Sistema")
            self.bot.click_relative(179, 8, wait_after=750)
            
            if not self.bot.find(var_strSeletorSistema, matching=0.97, waiting_time=5000):
                self.__not_found(var_strSeletorSistema)
            self.bot.click(wait_after=750)
            var_LogMaestro = self.var_clssMaestro.write_log("Selecionou sistema.",arg_enumLogLevel=LogLevel.INFO)
            
            if not self.bot.find( "Slct_Filtro_Arquivo", matching=0.97, waiting_time=5000):
                self.__not_found("Slct_Tipo_Arquivo")
            self.bot.click_relative(168, 7, wait_after=750)
            var_LogMaestro = self.var_clssMaestro.write_log("Selecionou tipo arquivo.",arg_enumLogLevel=LogLevel.INFO)

            if not self.bot.find( "Slct_Tipo_Arquivo_Escrituracao", matching=0.97, waiting_time=5000):
                self.__not_found("Slct_Tipo_Arquivo_Escrituracao")
            self.bot.click(wait_after=750)
 
            if not self.bot.find( "Slct_Filtro_Pesquisa", matching=0.97, waiting_time=5000):
                self.__not_found("Slct_Filtro_Pesquisa")
            self.bot.click_relative(199, 13, wait_after=750)
            var_LogMaestro = self.var_clssMaestro.write_log("Selecionou tipo pesquisa.",arg_enumLogLevel=LogLevel.INFO)

            if not self.bot.find( "Slct_Pesquisa_Periodo", matching=0.97, waiting_time=5000):
                self.__not_found("Slct_Pesquisa_Periodo")
            self.bot.click(wait_after=750)

            if self.bot.find( "Fld_DataInicio", matching=0.97, waiting_time=3000):
                self.bot.click_relative(124, 8)

            elif self.bot.find( "Fld_DataInicio_Fiscal", matching=0.97, waiting_time=2000):
                    self.bot.click_relative(348, 7)

            else: 
                self.__not_found("Fld_DataInicio")
                            
            self.bot.control_a()
            self.bot.delete()
            self.bot.type_keys(self.var_strDataInicio)
            self.bot.enter()
            var_LogMaestro = self.var_clssMaestro.write_log("Preencheu 'Data de início'.",arg_enumLogLevel=LogLevel.INFO)

            if self.bot.find( "Fld_DataFim", matching=0.97, waiting_time=5000):
                self.bot.click_relative(127, 7)
            
            elif self.bot.find( "Fld_DataFim_Fiscal", matching=0.97, waiting_time=5000):
                self.bot.click_relative(344, 7)

            else:    
                self.__not_found("Fld_DataFim")

            self.bot.control_a()
            self.bot.delete()
            self.bot.type_keys(self.var_strDataFim)
            self.bot.enter()
            var_LogMaestro = self.var_clssMaestro.write_log("Preencheu 'Data de fim'.",arg_enumLogLevel=LogLevel.INFO)
        

            if self.bot.find( "img_Erro_Procuracao", matching=0.97, waiting_time=10000) != None:
                raise Exception("ERRO APP. RECEITANET. Mensagem: Não existe procuração eletrônica para o detentor.")           
            
        
            if arg_strTipo == "efd_fiscal":
                if not self.bot.find( "box_Todos_Estabelecimentos", matching=0.97, waiting_time=10000):
                    self.__not_found("box_Todos_Estabelecimentos")
                self.bot.click_relative(226, 7)
                
                # if not self.bot.find( "box_Ultimo_Arquivo", matching=0.97, waiting_time=10000):
                #     self.__not_found("box_Ultimo_Arquivo")
                # self.bot.click_relative(320, 8)

            self.__set_focus()

            time.sleep(3)        

            if not self.bot.find( "Btn_Pesquisar_Filtros", matching=0.97, waiting_time=5000):
                self.__not_found("Btn_Pesquisar_Filtros")
            self.bot.click()
            
            time.sleep(5)        

            #Realiza segundo click para carregar mais informações 
            if not self.bot.find_text( "btn_Pesquisar2", threshold=230, waiting_time=10000):
                self.not_found("btn_Pesquisar2")
            self.bot.click()
            
            if self.bot.find( "img_Erro_Pesquisa_Filtro", matching=0.97, waiting_time=5000):
                raise BusinessRuleException("Erro ao Realizar consulta.")
            
            elif self.bot.find( "img_Aviso_Pesquisa", matching=0.97, waiting_time=5000):
                raise BusinessRuleException("Erro ao Realizar consulta.")


            if not self.bot.find( "img_Resultado_Pesquisa", matching=0.97, waiting_time=5000):
                self.__not_found("img_Resultado_Pesquisa")
            self.bot.click_relative(15, 24)

        except Exception as exception:
            raise Exception("ERRO APP. RECEITANET. Mensagem: Erro ao realizar o Preenchimento Aplicativo Receitanet: \n" + str(exception))

        
    def deselect_files(self, arg_lstLinesToRemove:list) -> None:

        self.__set_focus()

        self.bot.tab()
        self.bot.tab()
        self.bot.type_up()

        var_intAux = 1
        for index in arg_lstLinesToRemove:
            
            pressDown = index - var_intAux

            var_intAux = index

            self.__set_focus()
            for n in range(pressDown):

                self.bot.type_down()
                    
            self.bot.space()


    def extract_table(self) -> str:

        var_LogMaestro = self.var_clssMaestro.write_log("Realizando extração da tabela com UiPath ",arg_enumLogLevel=LogLevel.INFO)

        try:
            # Files().clear_Folder( arg_strPath=ROOT_DIR + "resources\\UiPath")
            Files().clear_Folder( arg_strPath=self.var_strRoot_dir + "resources\\UiPath")

            # var_strCommand = self.var_strPathRobotExe + " -file " + ROOT_DIR + "resources\\UiPath\\ExtractTableReceita\\Main.xaml "
            var_strCommand = self.var_strPathRobotExe + " -file " + self.var_strRoot_dir + "resources\\UiPath\\ExtractTableReceita\\Main.xaml "

            subprocess.run(var_strCommand)

        
        except Exception as exception:
            raise Exception("ERRO APP. RECEITANET. Mensagem: Erro ao extrair a Tabela do Aplicativo Receitanet: \n" + str(exception))

        # return ROOT_DIR + "resources\\UiPath\\NewTable.xlsx"
        return self.var_strRoot_dir + "resources\\UiPath\\NewTable.xlsx"


    def download_arquivos(self) -> None:

        self.__set_focus()

        try:
            if not self.bot.find( "Btn_Solicitar_Arquivos", matching=0.97, waiting_time=45000):
                self.__not_found("Btn_Solicitar_Arquivos")
            self.bot.click()
            var_LogMaestro = self.var_clssMaestro.write_log("Click em 'Solicitar arquivos'.",arg_enumLogLevel=LogLevel.INFO)
           
            cont = 0
            var_aux = False
            while cont < 10 and var_aux == False :
                time.sleep(5)
                if self.bot.find( "img_Numero_Pedido", matching=0.97, waiting_time=30000):
                    self.bot.enter()
                    var_LogMaestro = self.var_clssMaestro.write_log("Click em 'OK'",arg_enumLogLevel=LogLevel.INFO)
                    var_aux = True
            cont += 1

            if not self.bot.find( "Btn_Acompanhamento", matching=0.97, waiting_time=5000):
                self.__not_found("Btn_Acompanhamento")
            self.bot.click()
            var_LogMaestro = self.var_clssMaestro.write_log("Click em 'Acompanhamento'.",arg_enumLogLevel=LogLevel.INFO)

            if not self.bot.find( "Btn_Pedidos_Arquivos", matching=0.97, waiting_time=5000):
                self.__not_found("Btn_Pedidos_Arquivos")
            self.bot.click()
            var_LogMaestro = self.var_clssMaestro.write_log("Click em 'Ver pedidos e arquivos'.",arg_enumLogLevel=LogLevel.INFO)
            
            # Click Número Pedido
            cont = 0
            var_aux = False
            while cont < 10 and var_aux == False :
                time.sleep(15)
                if self.bot.find( "img_Pedido_Realizado", matching=0.97, waiting_time=5000):
                    self.bot.double_click_relative(31, 28)
                    var_LogMaestro = self.var_clssMaestro.write_log("Click no Primeiro Pedido.",arg_enumLogLevel=LogLevel.INFO)
                    var_aux = True
                else:    
                    self.__not_found("img_Pedido_Realizado")
            cont += 1

            time.sleep(75)
            
            # Teste novo do click relativo
            if not self.bot.find( "img_Arquivos_Click_Relativo", matching=0.97, waiting_time=10000):
                 self.not_found("img_Arquivos_Click_Relativo")
            self.bot.click_relative(12, 28)
            var_LogMaestro = self.var_clssMaestro.write_log("Click em 'CheckBox'.",arg_enumLogLevel=LogLevel.INFO)

            time.sleep(10)
            
            if self.bot.find( "box_Checked_AllFiles", matching=0.97, waiting_time=20000):
                time.sleep(25)        

            if not self.bot.find( "Btn_Download_Arquivos", matching=0.97, waiting_time=60000):
                self.__not_found("Btn_Download_Arquivos")
            self.bot.click()
            var_LogMaestro = self.var_clssMaestro.write_log("Click em 'Baixar'",arg_enumLogLevel=LogLevel.INFO)
            
            if not self.bot.find( "Btn_Fila_Download", matching=0.97, waiting_time=5000):
                self.__not_found("Btn_Fila_Download")
            self.bot.click()
            var_LogMaestro = self.var_clssMaestro.write_log("Click em 'Fila de download'",arg_enumLogLevel=LogLevel.INFO)

        except Exception as exception:
            raise Exception("ERRO APP. RECEITANET. Mensagem: Erro ao Selecionar o Número do Pedido: \n" + str(exception))

        self.__wait_fila_download()


    def __wait_fila_download(self) -> None:
        
        var_LogMaestro = self.var_clssMaestro.write_log(f"Aguardando finalização do download.",arg_enumLogLevel=LogLevel.INFO)
        var_contador = 0
        
        while True:

            if self.bot.find( "img_Fila_Vazia", matching=0.97, waiting_time=7500):
                break

            self.__set_focus()
            
            if self.bot.find( "SuspensoAte", matching=0.97, waiting_time=10000):
                
                time.sleep(420)

                if self.bot.find( "img_Fila_Vazia", matching=0.97, waiting_time=7500):
                    break
                else: 
                    var_contador += 1
                    if var_contador == 2:
                        raise Exception("ERRO APP. RECEITANET. Mensagem: Limite de tempo suspenso excedido")

            self.__set_focus()

        var_LogMaestro = self.var_clssMaestro.write_log(f"Sem mais itens para download.",arg_enumLogLevel=LogLevel.INFO)

    def __set_focus(self) -> None:

        try:
            self.app = Application(backend='uia').connect(title_re='Receitanet BX')
                
            var_winReceita = self.app.window(title_re='Receitanet BX')

        except:
            raise Exception("Receitanet Fechou inesperadamente durante o Download dos Arquivos.")
        
        else:
            var_winReceita.set_focus()


    def __not_found(self, label) -> None:
        raise ElementNotFound(f"Element not found: {label}")

















