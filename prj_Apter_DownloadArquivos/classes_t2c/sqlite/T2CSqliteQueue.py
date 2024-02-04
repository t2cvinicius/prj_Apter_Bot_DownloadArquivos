from prj_Apter_DownloadArquivos.classes_t2c.utils.T2CMaestro import T2CMaestro, LogLevel, ErrorType
from pathlib import Path

import sqlite3
import datetime

ROOT_DIR = Path(__file__).parent.parent.parent.__str__()

"""
ESTRUTURA DA TABELA DO PROJETO:

CREATE TABLE tbl_Fila_DownloadArquivos (
    id                 INTEGER        PRIMARY KEY,
    referencia         VARCHAR (200),
    datahora_criado    VARCHAR (50),
    ultima_atualizacao VARCHAR (50),
    nome_maquina       VARCHAR (200),
    status             VARCHAR (100),
    obs                VARCHAR (500),
    retry              VARCHAR (2),
    id_download        VARCHAR (10),
    tipo_projeto       VARCHAR (50),
    projeto            VARCHAR (50),
    email              VARCHAR (75),
    cnpj               VARCHAR (20),
    procuracao         VARCHAR (50),
    data_inicio        VARCHAR (10),
    data_fim           VARCHAR (10),
    efd_contribuicoes  VARCHAR (4),
    efd_fiscal         VARCHAR (4),
    ecd                VARCHAR (4),
    ecf                VARCHAR (4),
    dctf               VARCHAR (4),
    darf               VARCHAR (4),
    tarefa_infos       VARCHAR (2000) 
);


COLUNAS PODEM SER ADICIONADAS, MAS NUNCA EXCLUÍDAS OU MOVIDAS
O ARQUIVO DO BANCO LOCALIZADO EM resources/sqlite/banco_dados.db JÁ POSSUI ESSA TABELA CRIADA E VAZIA
"""

class T2CSqliteQueue:
    """
    #Classe responsável para manipulação do sqlite e controle de fila.
    """

    def __init__(self, arg_clssMaestro:T2CMaestro, arg_strTabelaFila:str, arg_strCaminhoBd:str=None, arg_strMaxRetry:str=None, arg_strNomeMaquina:str=None):
        """
        Inicializa a classe T2CSqliteQueue.
        - Cria a conexão com o banco
        - Nome da máquina precisa ser algum identificador único por execução
        - CaminhoBD e TabelaFila não precisam ser informados por padrão
        
        Parâmetros:
        - arg_clssMaestro (T2CMaestro): instância de T2CMaestro.
        - arg_strCaminhoBd (str): caminho do banco de dados (opcional, default=None).
        - arg_strTabelaFila (str): nome da tabela da fila (opcional, default=None).
        - arg_strMaxRetry (str): Número Máximo de repetições permitidos pelo processo.
        - arg_strNomeMaquina (str): nome da máquina (opcional, default=None).
        """
        self.var_clssMaestro = arg_clssMaestro
        self.var_strTabelaFila = arg_strTabelaFila
        self.var_intMaxRetryNumber = int( arg_strMaxRetry )
        self.var_strNomeMaquina = arg_strNomeMaquina
        self.var_strPathToDb = ROOT_DIR + "\\resources\\sqlite\\banco_dados.db" if(arg_strCaminhoBd is None) else arg_strCaminhoBd
        var_csrCursor = sqlite3.connect(self.var_strPathToDb).execute("SELECT * FROM " + self.var_strTabelaFila + " WHERE status = 'NEW'")

        self.var_intItemsQueue = len(var_csrCursor.fetchall())
        var_csrCursor.close()


    def update(self):
        """
        Atualiza a própria classe, usado em vários pontos do projeto para atualizar os itens na fila
        """
        var_csrCursor = sqlite3.connect(self.var_strPathToDb).execute("SELECT * FROM " + self.var_strTabelaFila + " WHERE status = 'NEW'")
        self.var_intItemsQueue = len(var_csrCursor.fetchall())
        var_csrCursor.close()


    def insert_new_queue_item(self, arg_strReferencia:str, arg_listInfAdicional:list) -> None:
        """
        Insere um item na tabela especificada, com a referência e com os valores adicionais
        
        Parâmetros:
        - arg_strReferencia (str): referência do item.
        - arg_listInfAdicional (list): informações adicionais.
        """

        #Aqui eu insere com o nome da máquina vazio, para que qualquer máquina possa processar em paralelo mais a frente
        var_strNow = datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        var_listValues = [arg_strReferencia, var_strNow, var_strNow, "", "NEW", "", "0"]
        var_listColumns = []
        
        var_csrCursor = sqlite3.connect(self.var_strPathToDb).execute("SELECT * FROM " + self.var_strTabelaFila + " WHERE id = 0")
        for description in var_csrCursor.description:
            if(description[0] != "id"): var_listColumns.append(description[0])

        var_listValues.extend(arg_listInfAdicional)

        self.update()

        #Construindo o comando insert
        var_strInsert = "INSERT INTO " + self.var_strTabelaFila + " (" + var_listColumns.__str__() + ") VALUES (" + var_listValues.__str__() + ")"
        var_strInsert = var_strInsert.replace("[", "").replace("]", "")
        
        #Executando o comando insert
        try:
            var_csrCursor.execute(var_strInsert)
            var_csrCursor.connection.commit()
        except Exception:
            self.var_clssMaestro.write_log(arg_strMensagemLog="Erro ao inserir linhas: " + str(Exception), arg_enumLogLevel=LogLevel.ERROR, arg_enumErrorType=ErrorType.APP_ERROR)
            raise

        var_csrCursor.close()
        self.update()


    def get_specific_queue_item(self, arg_intIndex:int) -> dict:
        """
        Retorna um item específico da fila.

        Parâmetros:
        - arg_intIndex (int): índice do item.

        Retorna:
        - dict: dicionário com as informações do item.
        """
        var_csrCursor = sqlite3.connect(self.var_strPathToDb).execute("SELECT * FROM " + self.var_strTabelaFila + " WHERE id = " + arg_intIndex.__str__())
        var_tplRow:tuple = var_csrCursor.fetchone()

        var_dctItem = dict()

        if( var_tplRow is not None):
            var_intCountRows = 0

            for item in var_csrCursor.description:
                var_dctItem[item[0]] = var_tplRow[var_intCountRows]
                var_intCountRows += 1

        else: var_dctItem = None

        var_csrCursor.close()
        self.update()

        return var_dctItem
   

    def get_next_queue_item(self) -> dict:
        """
        Retorna o próximo item da fila que não foi processado e não possui máquina alocada, None se não existe nenhum item assim.

        Retorna:
        - dict: Dicionário com as informações do item da fila.
        """
        var_dctQueueItem = dict()
        var_strNow = datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        
        var_csrCursor = sqlite3.connect(self.var_strPathToDb).execute("UPDATE " + self.var_strTabelaFila + " SET ultima_atualizacao = '" + var_strNow + "', nome_maquina = '" + self.var_strNomeMaquina + "', status = 'ON QUEUE' WHERE id = (SELECT MIN(id) FROM " + self.var_strTabelaFila + " WHERE status = 'NEW')").connection.commit()
        var_csrCursor = sqlite3.connect(self.var_strPathToDb).execute("SELECT * FROM " + self.var_strTabelaFila + " WHERE nome_maquina = '" + self.var_strNomeMaquina + "' and status = 'ON QUEUE' ORDER BY id LIMIT 1")
        var_tplRow:tuple = var_csrCursor.fetchone()

        if(var_tplRow is not None):

            var_lstFields = [field[0] for field in var_csrCursor.description]
            
            var_dctQueueItem = dict( zip(var_lstFields, var_tplRow) )

            var_csrCursor.execute("UPDATE " + self.var_strTabelaFila + " SET status = 'RUNNING' WHERE id = " + str(var_dctQueueItem['id']) ).connection.commit()
            
            self.var_clssMaestro.var_strReferencia = str(var_dctQueueItem['referencia'])

            var_csrCursor.close()

            self.update()
        
        return var_dctQueueItem



    def update_status_item(self, arg_intIndex:int, arg_strNovoStatus:str, arg_strObs:str=""):
        """
        Atualiza o status de um item com um determinado índice.

        Parâmetros:
        - arg_intIndex (int): índice do item.
        - arg_strNovoStatus (str): novo status do item.
        - arg_strObs (str): observação (opcional, default= "").
        """

        #Tratando os casos onde obs vem com quotes, trocando por *
        arg_strObs = arg_strObs.replace('"', '*').replace("'", '*')

        var_strNow = datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        var_csrCursor = sqlite3.connect(self.var_strPathToDb).execute("UPDATE " + self.var_strTabelaFila + " SET ultima_atualizacao = '" + var_strNow + "', status = '" + arg_strNovoStatus + "', obs = '" + arg_strObs + "' WHERE id = " + arg_intIndex.__str__())
        var_csrCursor.connection.commit()
   
        self.var_clssMaestro.var_strReferencia = '-'
        
        var_csrCursor.close()
        self.update()


    def abandon_queue(self):
        """
        Marca todos os itens com status NEW, ON QUEUE ou RUNNING como ABANDONED.
        """
        var_csrCursor = sqlite3.connect(self.var_strPathToDb).execute("UPDATE " + self.var_strTabelaFila + " SET status = 'ABANDONED' WHERE status in ('NEW', 'ON QUEUE', 'RUNNING')")
        var_csrCursor.connection.commit()

        var_csrCursor.close()
        self.update()
 
   
    def retry_queue_item(self, arg_dctQueueItem:dict,  arg_strStatus:str, arg_strObs:str):
        """
        Atualiza o Status do Item atual para 'RETRIED' e, caso não tenha excedido o número de tentativas 'MaxRetryNumber',
        adiciona um novo item na fila para executar novamente

        Parâmetros: 
        - arg_dctQueueItem (dict): NamedTuple com as informações do item da fila.
        - arg_strStatus (str): Status que será salvo caso atinja o máximo de tentativas.
        - arg_strObs (str): Informações sobre o Erro
        """

        if( int( str( arg_dctQueueItem['retry'] ) ) < self.var_intMaxRetryNumber ):
            self.update_status_item(arg_intIndex=int( str(arg_dctQueueItem['id']) ), arg_strNovoStatus="RETRIED", arg_strObs=arg_strObs)

            self.var_clssMaestro.write_log(arg_strMensagemLog= "Inserindo o item id -"+ str( arg_dctQueueItem['id'] ) +"- de volta na fila pra Reprocessamento." )

            var_listColumns = []
            var_listValues = []

            # Recuperando Nome das Colunas da Tabela
            var_csrCursor = sqlite3.connect(self.var_strPathToDb).execute("SELECT * FROM " + self.var_strTabelaFila + " WHERE id = "+ str( arg_dctQueueItem['id'] )  )
            var_listColumns = [description[0] for description in var_csrCursor.description if description[0] != "id" ]

            # Adicionando Valores às colunas Fixas do novo Item 
            var_strNow = datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")
            var_intRetry = int( str(arg_dctQueueItem['retry']) ) + 1
            var_listValues = [arg_dctQueueItem['referencia'], var_strNow, var_strNow, "", "NEW", "", var_intRetry ]

            # Pula as primeiras linhas 7 Colunas de informações Fixas, settadas acima, e preenche as colunas restantes com as informações específicas do item
            i = 8
            while ( i < len( arg_dctQueueItem ) ):
                var_listValues.append( list( arg_dctQueueItem )[i] )
                i += 1
            
            #Construindo o comando insert
            var_strInsert = "INSERT INTO " + self.var_strTabelaFila + " (" + var_listColumns.__str__() + ") VALUES (" + var_listValues.__str__() + ")"
            var_strInsert = var_strInsert.replace("[", "").replace("]", "").replace("None","NULL")
            
            #Executando o comando insert
            var_csrCursor.execute(var_strInsert)
            var_csrCursor.connection.commit()
            var_csrCursor.close()
        
        else:
            self.var_clssMaestro.write_log(arg_strMensagemLog= "Máximo de Tentativas Atingido. O Item id -" + str(arg_dctQueueItem['id']) +"- será salvo com status -" + arg_strStatus + "-.")
            self.update_status_item( arg_intIndex=int( str(arg_dctQueueItem['id']) ), arg_strNovoStatus=arg_strStatus, arg_strObs=arg_strObs )

            
        self.update()


