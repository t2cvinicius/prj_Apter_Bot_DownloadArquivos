from prj_Apter_DownloadArquivos.classes_t2c.utils.FilesTreatment import Files
from prj_Apter_DownloadArquivos.classes_t2c.utils.T2CMaestro import T2CMaestro, LogLevel, ErrorType

import os
import fitz
import numpy
import regex
import pandas
import datetime
import time
import pdfplumber
from openpyxl import load_workbook


class Planilhamento():
    
    def __init__(self, arg_clssMaestro:T2CMaestro,arg_strPathExcel:str, arg_strPathProject:str, arg_strPathDownload:str, arg_dctQueue:dict) -> None:
        self.var_strPathExcel = arg_strPathExcel
        self.var_strPathProject = arg_strPathProject
        self.var_strPathDownload = arg_strPathDownload
        self.var_clssMaestro = arg_clssMaestro

        self.var_datInicio = datetime.datetime.strptime(arg_dctQueue['data_inicio'], "%d-%m-%Y")
        self.var_datFim = datetime.datetime.strptime(arg_dctQueue['data_fim'], "%d-%m-%Y")


    def read_cnpj(self, arg_strPathPdf:str) -> None:
        """
        
        """
        
        var_dctData = dict()
        
        with fitz.open(arg_strPathPdf) as pdf:
            txt = pdf[0].get_text()

        # Leitura dos Campos do Arquivo Cartão CNPJ, por meio de Regex
        try:
            var_dctData['cnpj']             = regex.search(r"(?<=NÚMERO DE INSCRIÇÃO)\s+(\d{2}\.\d{3}\.\d{3}\/\d{4}\-\d{2})", txt).group().strip()

            var_dctData['dt_abertura']      = regex.search(r"(?<=DATA DE ABERTURA)\s+(\d{2}\/\d{2}\/\d{4})\s+(?=NOME EMPRESARIAL)", txt).group().strip()

            var_dctData['nm_empresarial']   = regex.search(r"(?<=NOME EMPRESARIAL)\s+((?:.+\s)+)\s*(?=TÍTULO DO ESTABELECIMENTO)", txt).group().strip().replace(r"\n", "")

            var_dctData['nm_fantasia']      = regex.search(r"(?<=NOME DE FANTASIA\))\s+((?:.+\s)+)\s*(?=PORTE)", txt).group().strip().replace(r"\n", "")

            var_dctData['porte']            = regex.search(r"(?<=PORTE)\s+((?:.+\s)+)\s*(?=CÓDIGO E DESCRIÇÃO DA ATIVIDADE ECONÔMICA PRINCIPAL)", txt).group().strip().replace(r"\n", "")

            var_dctData['cnae_principal']   = regex.search(r"(?<=ECONÔMICA PRINCIPAL)\s+((?:.+\s)+)\s*(?=CÓDIGO E DESCRIÇÃO DAS ATIVIDADES ECONÔMICAS SECUNDÁRIAS)", txt).group().strip().replace(r"\n", "")

            var_dctData['cnae_secundario']  = regex.search(r"(?<=ECONÔMICAS SECUNDÁRIAS)\s+((?:.+\s)+)\s*(?=CÓDIGO E DESCRIÇÃO DA NATUREZA JURÍDICA)", txt).group().strip().replace(r"\n", "")

            var_dctData['natureza']         = regex.search(r"(?<=NATUREZA JURÍDICA)\s+((?:.+\s)+)\s*(?=LOGRADOURO)", txt).group().strip().replace(r"\n", "")

            var_dctData['logradouro']       = regex.search(r"(?<=LOGRADOURO)\s+((?:.+\s)+)\s*(?=NÚMERO)", txt).group().strip().replace(r"\n", "")

            var_dctData['numero']           = regex.search(r"(?<=NÚMERO)\n+((?:.+\s)+)\s*(?=COMPLEMENTO)", txt).group().strip().replace(r"\n", "")

            var_dctData['complemento']      = regex.search(r"(?<=COMPLEMENTO)\s+((?:.+\s)+)\s*(?=CEP)", txt).group().strip().replace(r"\n", "")

            var_dctData['cep']              = regex.search(r"(?<=CEP)\s+((?:.+\s)+)\s*(?=BAIRRO)", txt).group().strip()

            var_dctData['bairro']           = regex.search(r"(?<=DISTRITO)\s+((?:.+\s)+)\s*(?=MUNICÍPIO)", txt).group().strip().replace(r"\n", "")

            var_dctData['municipio']        = regex.search(r"(?<=MUNICÍPIO)\s+((?:.+\s)+)\s*(?=UF)", txt).group().strip().replace(r"\n", "")

            var_dctData['uf']               = regex.search(r"(?<=UF)\s+(\w{2})", txt).group().strip()

            var_dctData['end_eletronico']   = regex.search(r"(?<=ENDEREÇO ELETRÔNICO)\s+((?:.+\s)*)\s*(?=TELEFONE)", txt).group().strip().replace(r"\n", "")

            var_dctData['telefone']         = regex.search(r"(?<=TELEFONE)\s+((?:.+\s)*)\s*(?=ENTE FEDERATIVO RESPONSÁVEL)", txt).group().strip()

            var_dctData['efr']              = regex.search(r"(?<=RESPONSÁVEL \(EFR\))\s+((?:.+\s)*?)\s*(?=SITUAÇÃO CADASTRAL)", txt).group().strip().replace(r"\n", "")

            var_dctData['cad_situacao']     = regex.search(r"(?<=SITUAÇÃO CADASTRAL)\s+((?:.+\s)*)\s*(?=DATA DA SITUAÇÃO CADASTRAL)", txt).group().strip()

            var_dctData['cad_data']         = regex.search(r"(?<=DATA DA SITUAÇÃO CADASTRAL)\s+(\d{2}\/\d{2}\/\d{4})", txt).group().strip()

            var_dctData['cad_motivo']       = regex.search(r"(?<=MOTIVO DE SITUAÇÃO CADASTRAL)\s+((?:.+\s)*?)\s*(?=SITUAÇÃO ESPECIAL)", txt).group().strip().replace(r"\n", "")

            var_dctData['esp_situacao']     = regex.search(r"(?<=SITUAÇÃO ESPECIAL)\s+((?:.+\s)*)\s*(?=DATA DA SITUAÇÃO ESPECIAL)", txt).group().strip()

            var_dctData['esp_data']         = regex.search(r"(?<=DATA DA SITUAÇÃO ESPECIAL)\s+((?:\*)*|(?:\d{2}\/\d{2}\/\d{4}))\s", txt).group().strip()

        except Exception as exception:
            raise Exception("Erro durante a leitura do 'Cartão CNPJ'."+ arg_strPathPdf +" \n" + str(exception) )


        var_dfData = pandas.DataFrame.from_dict(data= var_dctData, orient="index", columns=['informacoes'])
        
        self.__append_line(arg_dtData= var_dfData, 
                         arg_strSheetName= 'SITUAÇÃO CADASTRAL', 
                         arg_strStartRow= 8, 
                         arg_intStartCol= 3)
                

    def read_darf(self, arg_strPathPdf:str) -> None:

        var_dctData = dict()
        
        #pdf = 
        with pdfplumber.open(arg_strPathPdf) as pdf:
        #for number, pageText in enumerate(pdf.pages):
            txt = pdf.pages[0].extract_text() 

        # Leitura dos Campos do Arquivo DARF, por meio de Regex
        try:
            var_dctData['ref']              = os.path.basename(arg_strPathPdf)

            var_dctData['cnpj']             = regex.search(r"(\d{2}\.\d{3}\.\d{3}\/\d{4}\-\d{2})", txt).group().strip()

            var_dctData['rz_social']        = regex.search(r"(?<=\d{2}\.\d{3}\.\d{3}\/\d{4}\-\d{2})\s((.*\s)+?)\s*(?=Período Apuração)", txt).group().strip().replace(r"\n", "")

            var_dctData['dt_apuracao']      = regex.findall(r"(\d{2}\/\d{2}\/\d{4})", txt)[0]

            var_dctData['dt_vencimento']    = regex.findall(r"(\d{2}\/\d{2}\/\d{4})", txt)[1]

            var_dctData['n_documento']      = regex.search(r"(?<=\d{2}\/\d{2}\/\d{4})\s(\d*)\s+(?=Composição do Documento)", txt).group().strip()

            var_dctData['codigo']           = regex.search(r"(?<=Multa Juros Total)\s+(\d{4})\s", txt).group().strip()

            var_dctData['descricao']        = regex.search(r"(?<=Multa Juros Total\s\d{4})\s+(.*?\s)(?=(?:\d*\.)?\d*\,\d{2})", txt).group().strip()

            var_dctData['principal']        = regex.findall(r"((?:\d{1,3}\.)*\d{1,3}\,\d{2})", txt)[2]

            var_dctData['multa']            = regex.findall(r"((?:\d{1,3}\.)*\d{1,3}\,\d{2})", txt)[3]

            var_dctData['juros']            = regex.findall(r"((?:\d{1,3}\.)*\d{1,3}\,\d{2})", txt)[4]

            var_dctData['total']            = regex.findall(r"((?:\d{1,3}\.)*\d{1,3}\,\d{2})", txt)[5]

            var_dctData['obs']              = ''
            
            var_dctData['banco']            = regex.search(r"(?<=Banco Data de Arrecadação)\s+(.*?\s)(?=\d{2}\/\d{2}\/\d{4})", txt).group().strip()

            var_dctData['dt_arrecadacao']   = regex.findall(r"(\d{2}\/\d{2}\/\d{4})", txt)[2]

            var_dctData['agencia']          = regex.search(r"(?<=Referência)\s+(.*?\s)*?(?=\d{4})", txt).group().strip()

            var_dctData['estabelecimento']  = regex.search(r"(\d{4})\s(?=(?:\d{1,3}\.)*\d{1,3}\,\d{2})", txt).group().strip()

            var_dctData['resituicao']       = regex.findall(r"((?:\d{1,3}\.)*\d{1,3}\,\d{2})", txt)[6]

            var_dctData['referencia']       = regex.search(r"(?<=\,\d{2})\s(.*?\s)?(?=Comprovante)", txt).group().strip()

            var_dctData['cod_controle']     = regex.search(r"(?<=sob o código de controle)\s(.*?\s)(?=A autenticidade deste comprovante)", txt).group().strip()

            self.__append_line(arg_dtData= var_dctData, arg_strSheetName= 'DARF')


        except Exception as exception:
            raise Exception("Erro durante a leitura do 'DARF'."+ arg_strPathPdf +" \n" + str(exception) )
    

    def read_dctf(self, arg_strPathPdf:str) -> None:

        var_dctData = dict()
        with fitz.open(arg_strPathPdf) as pdf:
            txt = pdf[0].get_text()

            try:                            
                var_dctData['ref']              = os.path.basename(arg_strPathPdf)

                var_dctData['tipo']             = regex.search(r"(CNPJ\:)\s(\d{2}\.\d{3}\.\d{3}\/\d{4}\-\d{2})\s(.*\/\d{2,4})", txt).group().strip().replace("\n", "")

                var_dctData['cnpj']             = regex.search(r"(\d{2}\.\d{3}\.\d{3}\/\d{4}\-\d{2})", txt).group().strip()

                var_dctData['n_declaracao']     = regex.search(r"(?<=Número da Declaração\:)\s(\d{3}\.\d{4}\.\d{4}\.\d{10})", txt).group().strip()

                var_dctData['n_recibo']         = regex.search(r"(?<=Número do Recibo\:)\s(\d{2}\.\d{2}\.\d{2}\.\d{2}\.\d{2}\-\d{2})", txt).group().strip()

                var_dctData['dt_recepcao']      = regex.search(r"(?<=Data de Recepção\:)\s(\d{2}\/\d{2}\/\d{4})", txt).group().strip()

                var_dctData['dt_processamento'] = regex.search(r"(?<=Data de Processamento\:)\s(\d{2}\/\d{2}\/\d{4})", txt).group().strip()


                for page in pdf:
                    txt:str = page.get_text()

                    if( not txt.upper().__contains__("DÉBITO APURADO E CRÉDITOS VINCULADOS") ):
                        continue

                    var_dctData['cnpj_1']           = var_dctData['cnpj']
                    
                    var_dctData['gp_tributo']       = regex.search(r"(?<=GRUPO DO TRIBUTO\s\:)\s((?:.*\s)*?)\s*(?=CÓDIGO RECEITA)", txt).group().strip()

                    if( txt.upper().__contains__("CNPJ DO ESTABELECIMENTO") ): 
                        var_dctData['cnpj_estab']       = regex.search(r"(?<=CNPJ DO ESTABELECIMENTO\s\:)\s(\d{2}\.\d{3}\.\d{3}\/\d{4}\-\d{2})", txt).group().strip()
                    else: var_dctData['cnpj_estab'] = ''
                    
                    var_dctData['cod_receita']      = regex.search(r"(?<=CÓDIGO RECEITA\s\:)\s(\d{4}\-\d{2})", txt).group().strip()

                    var_dctData['periodicidade']    = regex.search(r"(?<=PERIODICIDADE\:)\s(.*\s)", txt).group().strip()

                    var_dctData['prd_apuracao']     = regex.search(r"(?<=PERÍODO DE APURAÇÃO\:)\s(.*\s)", txt).group().strip()

                    var_dctData['deb_apurado']      = regex.search(r"(?<=DÉBITO APURADO)\s((?:\d{1,3}\.)*\d{1,3}\,\d{2})", txt).group().strip()

                    var_dctData['credv_pgto']       = regex.search(r"(?<=\- PAGAMENTO)\s((?:\d{1,3}\.)*\d{1,3}\,\d{2})", txt).group().strip()
                    
                    var_dctData['credv_compens']    = regex.search(r"(?<=\- COMPENSAÇÕES)\s((?:\d{1,3}\.)*\d{1,3}\,\d{2})", txt).group().strip()

                    if( txt.upper().__contains__("- DEDUÇÃO COM DARF") ): 
                        var_dctData['credv_darf']       = regex.search(r"(?<=\- DEDUÇÃO COM DARF)\s((?:\d{1,3}\.)*\d{1,3}\,\d{2})", txt).group().strip()
                    else: var_dctData['credv_darf'] = ''
                    
                    var_dctData['credv_pclto']      = regex.search(r"(?<=\- PARCELAMENTO)\s((?:\d{1,3}\.)*\d{1,3}\,\d{2})", txt).group().strip()

                    var_dctData['credv_spnsao']     = regex.search(r"(?<=\- SUSPENSÃO)\s((?:\d{1,3}\.)*\d{1,3}\,\d{2})", txt).group().strip()

                    var_dctData['credv_soma']       = regex.search(r"(?<=SOMA DOS CRÉDITOS VINCULADOS\:)\s*((?:\d{1,3}\.)*\d{1,3}\,\d{2})", txt).group().strip()

                    var_dctData['deb_saldopg']      = regex.search(r"(?<=SALDO A PAGAR DO DÉBITO\:)\s((?:\d{1,3}\.)*\d{1,3}\,\d{2})", txt).group().strip()

                    var_dctData['deb_valor']        = regex.search(r"(?<=Total\:)\s*((?:\d{1,3}\.)*\d{1,3}\,\d{2})\s*(?=Total)", txt).group().strip()

                    var_dctData['total']            = regex.search(r"(?<=efetuadas as compensações\:)\s((?:\d{1,3}\.)*\d{1,3}\,\d{2})", txt).group().strip()


                    self.__append_line(arg_dtData= var_dctData, arg_strSheetName='DCTF')

            except Exception as exception:
                raise Exception("Erro durante a leitura do 'DCTF'. "+ arg_strPathPdf +" \n" + str(exception) )
            

    def read_table_receita(self, arg_strPathTable:str, arg_strTipo:str) -> pandas.DataFrame:
        
        var_dfTable = pandas.read_excel(io= arg_strPathTable, sheet_name='Sheet', engine='openpyxl', date_format="mm/dd/yyyy")

        var_dfTable.index = var_dfTable.index + 1

        if( arg_strTipo == "efd_contribuicoes" ):
            var_dtiDateTimeIndex = pandas.date_range(self.var_datInicio, self.var_datFim, freq='M')

            var_dfTable['Data Fim'] = pandas.to_datetime(var_dfTable['Data Fim'], format='%Y-%m-%dT%H:%M:%S', errors='coerce')

            var_dfTable['Data Início'] = pandas.to_datetime(var_dfTable['Data Início'], format='%Y-%m-%dT%H:%M:%S', errors='coerce')

            var_dfTable['Transmissão'] = pandas.to_datetime(var_dfTable['Transmissão'], format='%Y-%m-%dT%H:%M:%S', errors='coerce')
            
            var_dfTable = var_dfTable.rename(columns={"Transmissão" : "column_filter"})

            var_dfTable['cod'] = var_dfTable['Recibo'].apply(self.__apply_filter_efd)

            var_dfTable = var_dfTable.rename(columns={"Contribuinte" : "CNPJ"})


        elif( arg_strTipo == "efd_fiscal" ):

            var_dtiDateTimeIndex = pandas.date_range(self.var_datInicio, self.var_datFim, freq='M')

            var_dfTable['Data Fim'] = pandas.to_datetime(var_dfTable['Data Fim'], format='%Y-%m-%d', errors='coerce')

            var_dfTable['Data Início'] = pandas.to_datetime(var_dfTable['Data Início'], format='%Y-%m-%d', errors='coerce')

            var_dfTable = var_dfTable.rename(columns={"Data Envio SPED" : "column_filter"})

            var_dfTable['cod'] = var_dfTable['Hash'].apply(self.__apply_filter_efd)

        elif( arg_strTipo == "ecd" ):
            var_dtiDateTimeIndex = pandas.date_range(self.var_datInicio, self.var_datFim, freq='Y')

            var_dfTable['Data Fim'] = pandas.to_datetime(var_dfTable['Data Fim'], format='%Y-%m-%d', errors='coerce')

            var_dfTable['Data Início'] = pandas.to_datetime(var_dfTable['Data Início'], format='%Y-%m-%d', errors='coerce')
          
            var_dfTable['Data Recepção'] = pandas.to_datetime(var_dfTable['Data Recepção'], format='%Y-%m-%dT%H:%M:%S', errors='coerce')

            var_dfTable = var_dfTable.rename(columns={"Data Recepção" : "column_filter"})

            var_dfTable['cod'] = var_dfTable['Hash']
            

        elif( arg_strTipo == "ecf" ):
            var_dtiDateTimeIndex = pandas.date_range(self.var_datInicio, self.var_datFim, freq='Y')

            var_dfTable['Data Fim'] = pandas.to_datetime(var_dfTable['Data Fim'], format='%Y-%m-%dT%H:%M:%S', errors='coerce')

            var_dfTable['Data Início'] = pandas.to_datetime(var_dfTable['Data Início'], format='%Y-%m-%dT%H:%M:%S', errors='coerce')
            
            var_dfTable['Transmissão'] = pandas.to_datetime(var_dfTable['Transmissão'], format='%Y-%m-%dT%H:%M:%S', errors='coerce')
            
            var_dfTable = var_dfTable.rename(columns={"Transmissão" : "column_filter"})
    
            var_dfTable['cod'] = var_dfTable['column_filter'].dt.strftime("%Y%m%d%H%M")

            var_dfTable = var_dfTable.rename(columns={"Contribuinte" : "CNPJ"})
    
        # Cria uma coluna 'index'
        var_dfTable = var_dfTable.reset_index()

        var_dfPeriodo = var_dtiDateTimeIndex.to_frame(0, 'Data Fim')

        # Mescla a tabela de informações com a tabela de Período (full join), unindo as duas tabelas pela coluna 'Data Fim'.
        var_dfMergedTable = pandas.merge(left=var_dfTable, right=var_dfPeriodo, on='Data Fim', how='outer')

        # Atualiza a coluna filtro para o tipo datetime64.
        var_dfMergedTable['column_filter'] = pandas.to_datetime(var_dfMergedTable['column_filter'])

        # Ordena as colunas de 'Data Fim' e 'column_filter' (asc).
        var_dfMergedTable = var_dfMergedTable.sort_values(by=['Data Fim', 'column_filter', 'CNPJ'], ascending=True).reset_index(drop=True)

        # Adiciona a coluna 'remove_dpt', inserindo 'True' nas linhas 'Data Fim' repetidas, apenas deixando como 'False' as linhas mais recentes.
        var_dfMergedTable['remove_dpt'] = var_dfMergedTable.duplicated(subset=['Data Fim', 'CNPJ'], keep='last' ).values

        # # Adiciona a coluna 'remove_out', inserindo 'True' nas linhas 'Data Fim' que estão fora do range de do Período.
        # var_dfMergedTable['remove_out'] = var_dfMergedTable['Data Fim'].isin(var_dtiDateTimeIndex).__invert__()

        # Preenche com '-1', as linhas em branco da coluna 'index' 
        var_dfMergedTable['index'] = pandas.DataFrame.astype(var_dfMergedTable['index'], dtype='Int8', errors='ignore').fillna(-1)

        # Adiciona uma coluna de status    
        var_dfMergedTable['status_download'] = numpy.where(var_dfMergedTable['remove_dpt'], 
                                                                    "NÃO BAIXADO. ENCONTRADO ARQUIVO MAIS RECENTE PARA O MESMO PERÍODO", 
                                                                    numpy.where(var_dfMergedTable['index'] == -1, 
                                                                                "NÃO BAIXADO. ARQUIVO NÃO ENCONTRADO", 
                                                                                pandas.NA))
        
        #Exlui linha da dataframe cao CNPJ caso ela seja NaN
        var_dfMergedTable.dropna(subset=['CNPJ'], inplace=True)           
            

        var_dfMergedTable = var_dfMergedTable.fillna('')

        return var_dfMergedTable


    def planilhamento_receitanet(self, arg_strTipo:str, arg_dfTable:pandas.DataFrame, arg_strPathTipo:str) -> None:

        var_LogMaestro = self.var_clssMaestro.write_log(f"Iniciando planilhamento do tipo {arg_strTipo}.",arg_enumLogLevel=LogLevel.INFO)
        
        arg_dfTable = arg_dfTable.sort_values(by=['Data Fim','column_filter'],ascending=False)
            
        if( arg_strTipo == 'efd_contribuicoes'): 
            var_strSheetName = 'EFD CONTRIBUIÇÕES'

        elif( arg_strTipo == 'efd_fiscal'): 
            var_strSheetName = 'EFD ICMS IPI'
        
        else: 
            var_strSheetName = arg_strTipo.upper()

        var_lstTable = self.__format_columns_receitanet(arg_strTipo=arg_strTipo, arg_dfTable= arg_dfTable)

        for row in var_lstTable:
            var_strCod = row.pop('cod')

            if(row['status_download'] == ''):
                var_lstFiles = Files().get_Files_By_ID(arg_strPath= self.var_strPathDownload, arg_strIdentificador= var_strCod)

                if ( len(var_lstFiles) > 0):
                    var_strPathFile = [arquivos for arquivos in var_lstFiles if arquivos.endswith('.txt')].pop()

                else:
                    row['status_download'] = "ARQUIVO NÃO ENCONTRADO NA PASTA DOWNLOAD"
                    continue

                row['status_download'] = "BAIXADO"

                row['nome_arquivo'] = os.path.basename(var_strPathFile)

                row['tamanho'] = str( round( os.path.getsize(var_strPathFile) / 1024 ) ) 

                for arquivos in var_lstFiles:
                    Files().move_to( arg_strPathOrigin= arquivos, arg_strPathDest= arg_strPathTipo)


            self.__append_line(arg_dtData= row, arg_strSheetName= var_strSheetName)
        
        var_LogMaestro = self.var_clssMaestro.write_log(f"Realizado a movimentação dos arquivos baixados.",arg_enumLogLevel=LogLevel.INFO)
       
        var_lstFilesRemanescentes = Files().get_Files(arg_strPath= self.var_strPathDownload)
        
        if( len(var_lstFilesRemanescentes) > 0): 
            var_strPathArquivosExcessao = os.path.join(arg_strPathTipo, "arquivos_excessao")

            Files().create_Folder(arg_strPath= var_strPathArquivosExcessao)
        
            for files in var_lstFilesRemanescentes:                
                Files().move_to(arg_strPathOrigin= files, arg_strPathDest=var_strPathArquivosExcessao)

    def lines_to_remove_dctf(self, arg_dtDataTable:list[dict]|pandas.DataFrame ) -> list[int]:
        
        if( type(arg_dtDataTable) == type(list()) ): var_dfTable = pandas.DataFrame.from_records(arg_dtDataTable)
        else: var_dfTable:pandas.DataFrame = arg_dtDataTable

                
        # Indices dos arquivos à serem removidos da lista de Download (Arquivos Obsoletos e Arquivos fora do período).
        return list( var_dfTable.loc[ var_dfTable['lines_to_remove'] ]['index'] )

    def __append_line(self, 
                    arg_dtData:dict|pandas.DataFrame, 
                    arg_strSheetName:str, 
                    arg_strStartRow:int = None, 
                    arg_intStartCol:int = 2) -> None:
        
        if( type(arg_dtData) == type(dict()) ): var_dfData = pandas.DataFrame(data=arg_dtData,index=[0])
        else: var_dfData:pandas.DataFrame = arg_dtData

        with pandas.ExcelWriter(self.var_strPathExcel, mode='a', engine="openpyxl", if_sheet_exists='overlay') as writer:
            if(arg_strStartRow is None): arg_strStartRow = writer.book.get_sheet_by_name( arg_strSheetName ).max_row

            var_dfData.to_excel(excel_writer= writer,
                                    sheet_name= arg_strSheetName,
                                    header= False,
                                    startrow= arg_strStartRow,
                                    startcol= arg_intStartCol,
                                    index= False)

            sheet = writer.sheets[arg_strSheetName]


    def __format_columns_receitanet(self,arg_strTipo:str, arg_dfTable:pandas.DataFrame) -> list[dict]:
        
        arg_dfTable['Data Fim'] = pandas.to_datetime(arg_dfTable['Data Fim'],errors='raise', format="%d/%m/%Y",dayfirst=True,utc=False )
        arg_dfTable['Data Início'] = pandas.to_datetime(arg_dfTable['Data Início'],errors='raise', format="%d/%m/%Y",dayfirst=True,utc=False )
        arg_dfTable["CNPJ"] = pandas.DataFrame.astype(arg_dfTable["CNPJ"], dtype='str',errors='ignore')

        if( arg_strTipo == "efd_contribuicoes" ):                    
            var_lstColToRemove = ['Column-0', 'SCP', 'index', 'remove_dpt']

        elif( arg_strTipo == "efd_fiscal" ):
            var_lstColToRemove = ['Column-0','index', 'remove_dpt']
            arg_dfTable["Inscrição Estadual"] = pandas.DataFrame.astype(arg_dfTable["Inscrição Estadual"], dtype='str',errors='ignore')
            
            
        elif( arg_strTipo == "ecd" ):
            var_lstColToRemove = ['Column-0', 'Forma', 'SCP', 'index', 'remove_dpt']                   


        elif( arg_strTipo == "ecf" ):
            var_lstColToRemove = ['Column-0', 'SCP', 'Retificadora', 'index', 'remove_dpt']


        mask = arg_dfTable['CNPJ'].astype(str).str.contains('\.')
        if mask.any():
            arg_dfTable['CNPJ'] = arg_dfTable['CNPJ'].astype(str).apply(lambda x:x[:-2])


        arg_dfTable = arg_dfTable.drop(columns=var_lstColToRemove)

        if arg_strTipo != "efd_fiscal" :
        
            var_serColumn = arg_dfTable.pop('Data Fim')
            arg_dfTable.insert(2, 'Data Fim', var_serColumn)

        arg_dfTable = arg_dfTable.fillna('')
        
        return arg_dfTable.to_dict('records')


    def __apply_filter_efd(self, x:str) -> str:
        return x.split('-')[0]

