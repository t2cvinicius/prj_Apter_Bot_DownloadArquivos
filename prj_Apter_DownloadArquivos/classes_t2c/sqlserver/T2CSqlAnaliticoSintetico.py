from prj_Apter_DownloadArquivos.classes_t2c.utils.T2CMaestro import T2CMaestro, LogLevel
from prj_Apter_DownloadArquivos.classes_t2c.sqlite.T2CSqliteQueue import T2CSqliteQueue

import datetime
import pyodbc

class T2CSqlAnaliticoSintetico:
    """
    Classe usada para gravar dados da execução no SQL Server, geralmente no banco da T2C.
    """

    def __init__(self, arg_clssMaestro:T2CMaestro, arg_dictConfig:dict):
        """
        Inicia uma instância da classe T2CSqlAnatilitoSintetico, realizando a criação de uma conexão e um cursor.

        Parâmetros:
            - arg_dictConfig (dict): Dicionário de configuração do framework.
            - arg_clssMaestro (T2CMaestro): Instância de uma classe T2CMaestro.
        """
        self.var_dictConfig = arg_dictConfig
        self.var_clssMaestro = arg_clssMaestro
        

    def connect(self):
        """
        Realiza a conexão com o SQL Server.
        """
        var_strDbServer = self.var_dictConfig["BdServer"]
        var_strDbDatabase = self.var_dictConfig["BdDatabase"]
        var_strDbUser = self.var_dictConfig["BdUsuario"]
        #var_strDbPassword = self.var_clssMaestro.get_credential(self.var_dictConfig["BdSenha"])

        #Troque para  em casos onde não é possível usar credenciais (developers, por exemplo)
        var_strDbPassword = self.var_dictConfig["BdSenha"]
        
        var_strConnectionString = ("DRIVER={SQL Server}" + 
                                   ";SERVER=" + var_strDbServer + 
                                   #";PORT=1443" +                      #Porta usada na conexão, para casos onde precisa usar uma específica devido firewall
                                   ";DATABASE=" + var_strDbDatabase + 
                                   ";UID=" + var_strDbUser + 
                                   ";PWD=" + var_strDbPassword)

        try:
            #Tenta criar uma conexão e instanciar um novo cursor
            self.var_sqlConn = pyodbc.connect(var_strConnectionString)
            self.var_csrCursor = self.var_sqlConn.cursor()
        except Exception as exception:
            self.var_clssMaestro.write_log("Erro ao conectar ao SQL Server")
            self.var_clssMaestro.write_log(str(exception))


    def disconnect(self):
        """
        Desconecta do SQL Server.
        """
        try:
            self.var_csrCursor.close()
            self.var_sqlConn.close()
        except Exception as exception:
            self.var_clssMaestro.write_log("Erro ao desconectar do SQL Server")
            self.var_clssMaestro.write_log(str(exception))


    def insert_linha_inicio_sintetico(self, arg_strNomeMaquina:str, arg_dateInicioExecucao:datetime):
        """
        Insere uma nova linha na tabela tbl_dados_sinteticos. Esse método é executado no começo do processo.

        Parâmetros:
            - arg_strNomeMaquina (str): Nome da máquina.
            - arg_dateInicioExecucao (datetime): Datetime representando o inicio da execução.
        """
        self.var_clssMaestro.write_log("Incluindo linha na tabela tbl_apter_resumo_execucao_robos no SQL Server")

        #Conectando ao SQL Server
        self.connect()

        try:
            #Insert inicial
            self.var_csrCursor.execute("INSERT INTO tbl_apter_resumo_execucao_robos (" + 
                                        "nome_processo" 
                                        ",nome_maquina"
                                        ",inicio_exec" 
                                        ") VALUES ("
                                        "'" + self.var_dictConfig["NomeProcesso"] + "'"        #nome_processo 
                                        ",'" + arg_strNomeMaquina + "'"                         #nome_maquina
                                        ",?"                                                    #inicio
                                        ")", arg_dateInicioExecucao)
            
            #Pegando o identity gerado e salvando numa variável
            var_rowIdentitySintetico = self.var_csrCursor.execute("SELECT @@IDENTITY").fetchone()
            self.var_strIdentitySintetico = str(var_rowIdentitySintetico[0])

            self.var_csrCursor.commit()

        except Exception as exception:
            self.var_clssMaestro.write_log("Erro ao incluir linha no SQL Server (sintético) \n" + exception.__str__(), 
                                           arg_enumLogLevel= LogLevel.ERROR)
            
        finally:
            #Desconectando do SQL Server
            self.disconnect()


    def update_linha_fim_sintetico(self,
                                   arg_intTotalItens:int,
                                   arg_intTotalItensSucesso:int,
                                   arg_intTotalItensBusinessEx:int,
                                   arg_intTotalItensAppEx:int, 
                                   arg_dateFimExecucao:datetime):
        """
        Inserindo campos novos em um update na linha na tabela tbl_dados_sinteticos. Esse método é executado no fim do processo.

        Parâmetros:
            - arg_intTotalItens (int): Total de itens processados no total.
            - arg_intTotalItensSucesso (int): Total de itens processados com sucesso.
            - arg_intTotalItensBusinessEx (int): Total de itens processados com business exception.
            - arg_intTotalItensAppEx (int): Total de itens processados com app exceptions.
            - arg_dateFimExecucao (datetime): Datetime representando o final da execução.
        """
        self.var_clssMaestro.write_log("Atualizando linha na tabela tbl_apter_resumo_execucao_robos no SQL Server com os dados finais")

        #Conectando ao SQL Server
        self.connect()

        try:
            #Fazendo update com base no id salvo anteriormente
            self.var_csrCursor.execute("UPDATE tbl_apter_resumo_execucao_robos SET " +
                                    "fim_exec = ?" +                                           #FimExecucao
                                    ",qtd_itens_fila = " + str(arg_intTotalItens) + 
                                    ",qtd_itens_sucesso = " + str(arg_intTotalItensSucesso) + 
                                    ",qtd_itens_business = " + str(arg_intTotalItensBusinessEx) + 
                                    ",qtd_itens_app = " + str(arg_intTotalItensAppEx)+ 
                                    "WHERE  id_execucao = " + self.var_strIdentitySintetico, arg_dateFimExecucao)
            self.var_csrCursor.commit()

        except Exception as exception:
            self.var_clssMaestro.write_log("Erro ao atualizar linha no SQL Server (sintético) \n" + exception.__str__(), 
                                           arg_enumLogLevel= LogLevel.ERROR) 

        finally:
            #Desconectando do SQL Server
            self.disconnect()


    def insert_linha_analitico(self, 
                               arg_dctItemFila:dict,
                               arg_strNomeProcesso:str,
                               arg_strNomeFila:str,  
                               arg_strStatusItem:str,
                               arg_dateInicioItem:datetime, 
                               arg_dateFimItem:datetime, 
                               arg_strTipoExcecao:str="", 
                               arg_strDescricaoExcecao:str=""
                               ):
        """
        Insere uma nova linha na tabela tbl_dados_analiticos. Esse método é executado no final de cada item.

        Parâmetros:
            arg_dctItemFila (dict): Dicionário referente ao item da fila por completo.
            arg_strNomeProcesso (str): Nome do Processo em execução
            arg_strNomeFila (str): Nome da fila (no caso do BotCity, nome da tabela de fila).
            arg_strStatusItem (str): Status final do item.
            arg_strTipoExcecao (str): Tipo de exceção (se alguma) do item. (default="")
            arg_strDescricaoExcecao (str): Descrição da exceção. (default="")
            arg_dateInicioItem (datetime): Datetime do início da execução do item.
            arg_dateFimItem (datetime): Datetime do final da execução do item.
        """
        self.var_clssMaestro.write_log("Incluindo linha na tabela tbl_apter_dados_analiticos_robos no SQL Server")

        #Realizando assigns e tratando variáveis
        var_strReferencia = arg_dctItemFila["referencia"].__str__()
        var_strItemFila = arg_dctItemFila.__str__().replace('"', "*").replace("'", "*")
        var_strDescricaoExcecao = arg_strDescricaoExcecao.replace('"', "*").replace("'", "*")

        #Conectando ao SQL Server
        self.connect()

        try:
            #Comando Insert
            self.var_csrCursor.execute("INSERT INTO tbl_apter_dados_analiticos_robos ("
                                       "id_execucao"
                                       ",nome_processo"
                                       ",data_hora_inicio"
                                       ",data_hora_fim"
                                       ",nome_fila"
                                       ",referencia"
                                       ",item_fila"
                                       ",status"
                                       ",tipo_excecao"
                                       ",descricao_excecao) VALUES(" +
                                       self.var_strIdentitySintetico +                          #IdSintetico
                                       ",'" + arg_strNomeProcesso + "'"+                        #NomeProcesso
                                       ",?"+                                                    #Inicio
                                       ",?"+                                                    #Fim
                                       ",'" + arg_strNomeFila + "'" +                           #NomeFila
                                       ",'" + var_strReferencia + "'" +                         #Referencia
                                       ",?"                                                     #ItemFila
                                       ",'" + arg_strStatusItem + "'" +                         #StatusItem
                                       ",'" + arg_strTipoExcecao + "'" +                        #TipoExcecao
                                       ",?"+                                                    #DescExcecao
                                       ")", arg_dateInicioItem, arg_dateFimItem, var_strItemFila, var_strDescricaoExcecao
                                       ) #Parâmetros
            
            self.var_csrCursor.commit()

        except Exception as exception:
            self.var_clssMaestro.write_log("Erro ao incluir linha no SQL Server (analítico) \n" + exception.__str__(), 
                                           arg_enumLogLevel= LogLevel.ERROR)

        finally:
            #Desconectando do SQL Server
            self.disconnect()


    def populate_process_queue(self, arg_clssQueueItem:T2CSqliteQueue) -> None:
        """
        Procura na tabela 'controle_arquivos_baixados' por itens que estejam com o 'status_projeto' nulo. 
        Os itens encontrados são inseridos na Fila de Itens.

        Parâmetros:
        - arg_clssQueueItem (T2CSqliteQueue): Classe responsável por inserir os itens na fila.
        """
        self.connect()

        var_lstRow = self.var_csrCursor.execute("SELECT TOP 5"+
                                                "id, "+
                                                "projeto, "+
                                                "email_gerente, "+
                                                "tipo_projeto, "+
                                                "cnpj_matriz, "+
                                                "procuracao_eletronica, "+
                                                "data_inicio, "+
                                                "data_fim, "+
                                                "efd_contribuicoes, "+
                                                "ecd, "+
                                                "ecf, "+
                                                "efd_fiscal, "+
                                                "dctf, "+
                                                "darf, "+
                                                "tarefa_infos "+ 
                                                "FROM controle_arquivos_baixados "+ 
                                                "WHERE status_projeto IS NULL "+
                                                "ORDER BY id ASC").fetchall()

        for row in var_lstRow:            
            try:
                var_strReferencia = datetime.datetime.now().strftime("%d%m%Y%H%M%S")

                arg_clssQueueItem.insert_new_queue_item(arg_strReferencia= var_strReferencia, arg_listInfAdicional=row)

            except Exception as exception:
                self.var_clssMaestro.write_log("Erro ao inserir o item id -"+ str(row[0])+"- na fila de item. \n" + str(exception),
                                               arg_enumLogLevel=LogLevel.ERROR)


        self.disconnect()


    def insert_controle_arquivos(self, arg_strInfosProjeto:str) -> None:

        self.connect()

        try:

            self.var_csrCursor.execute("INSERT INTO "+
                                    "controle_arquivos_baixados ("+
                                    "tipo_projeto "+
                                    ",projeto "+
                                    ",email_gerente "+
                                    ",cnpj_matriz "+
                                    ",procuracao_eletronica "+
                                    ",data_inicio "+
                                    ",data_fim "+
                                    ",efd_contribuicoes "+
                                    ",efd_fiscal "+
                                    ",ecd "+
                                    ",ecf "+
                                    ",dctf "+
                                    ",darf "+
                                    ",tarefa_infos) VALUES("+
                                    arg_strInfosProjeto +
                                    ")")
            
            self.var_csrCursor.commit()

        except Exception as exception:
            self.var_clssMaestro.write_log("Erro ao incluir linha no SQL Server (controle_arquivos) \n" + exception.__str__(), 
                                           arg_enumLogLevel= LogLevel.ERROR)
            
            raise Exception("Erro ao incluir linha no SQL Server (controle_arquivos) \n" + exception.__str__())
            
        finally:
            #Desconectando do SQL Server
            self.disconnect()


    def update_status_arquivo(self, arg_strTipo:str, arg_strIdDownload:str, arg_strStatus:str) -> None:
        

        self.var_clssMaestro.write_log("Atualizando status da consulta -"+ arg_strTipo +"- para "+arg_strStatus+".")

        #Conectando ao SQL Server
        self.connect()

        try:
            # Atualizando status do arquivo processado
            self.var_csrCursor.execute("UPDATE controle_arquivos_baixados "+ 
                                       "SET "+ arg_strTipo +" = '"+arg_strStatus+"' "+ 
                                       "WHERE id = "+ arg_strIdDownload
                                       )

            self.var_csrCursor.commit()

        except Exception as exception:
            self.var_clssMaestro.write_log("Erro ao atualizar linha no SQL Server (controle_arquivos_baixados) \n" + exception.__str__(), 
                                           arg_enumLogLevel= LogLevel.ERROR) 

        finally:
            #Desconectando do SQL Server
            self.disconnect()


    def get_lines_to_remove_dctf(self, var_lstTable:list[dict], arg_dateDtInicio:datetime.datetime, arg_dateDtFim:datetime.datetime) -> list[int]:

        self.connect()

        """
        var_lstTable.Columns("CNPJ").ColumnName = "cnpj"
        var_lstTable.Columns("Período").ColumnName = "periodo"
        var_lstTable.Columns("Data de Recepção").ColumnName = "data_de_recepcao"
        var_lstTable.Columns("Período Inicial").ColumnName = "data_inicial"
        var_lstTable.Columns("Período Final").ColumnName = "data_final

        """
        var_lstLinesToSearch:list[int] = []
        var_strDtInicio = arg_dateDtInicio.strftime("%Y-%m-%d")
        var_strDtFim = arg_dateDtFim.strftime("%Y-%m-%d")


        try:
            self.var_csrCursor.execute("EXEC pr_Apter_LimpaTableArquivos")

            for row in var_lstTable:
                var_strCommandInst = "INSERT INTO AUX_Arquivos_DCTF ("+ list(row.keys()).__str__() +") VALUES("+ list(row.values()).__str__() +")"
                var_strCommandInst = var_strCommandInst.replace("[", "").replace("]", "")

                self.var_csrCursor.execute(var_strCommandInst)
                self.var_csrCursor.commit()

            var_strCommandPrd = f"EXEC pr_Apter_Arquivos_DCTF @periodo_inicial = {var_strDtInicio}, @periodo_final = {var_strDtFim}"
            var_lstRow = self.var_csrCursor.execute(var_strCommandPrd).fetchall()

            for row in var_lstRow:
                var_lstLinesToSearch += int( str(row[0]) )

        except:
            pass

        return var_lstLinesToSearch