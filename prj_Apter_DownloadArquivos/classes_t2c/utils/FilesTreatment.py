from prj_Apter_DownloadArquivos.classes_t2c.utils.T2CExceptions import BusinessRuleException
from pathlib import Path

import os
import pyunpack
import shutil
import pyautogui
import datetime
import pandas

ROOT_DIR = Path(__file__).parent.parent.parent.__str__()



class Files:


    def create_Folder(self, arg_strPath:str) -> None:
        """
        Cria uma Pasta no caminho indicado. Caso a Pasta não exista.

        Parâmetros:
        - arg_strPath (str): Caminho que será criado.
        """
        if( not (self.exists_Path(arg_strPath= arg_strPath)) ):
            os.mkdir(arg_strPath)
    

    def unzip_File(self, arg_strPathFile:str, arg_strPathDest:str) -> str:
        """
        Descompacta arquivo.

        Parâmetros:
        - arg_strPathFile (str): Caminho onde se encontra o Arquivo Compactado.
        - arg_strPathDest (str): Caminho onde o arquivo será descompactado.
        """        
        var_strFolderName = os.path.splitext(arg_strPathFile)[0]
        var_strFolderName = os.path.basename(var_strFolderName)

        var_strPathFolder = os.path.join( arg_strPathDest, var_strFolderName)

        pyunpack.Archive(arg_strPathFile).extractall(directory=var_strPathFolder, auto_create_dir= True)

        # os.remove(arg_strPathFile)

        var_strRootPath = self.__check_Folder_Zip(var_strPathFolder)

        return var_strRootPath


    def zip_Folder(sefl, arg_strPathFolder) -> str:
        """
        Compacta uma Pasta no formato .zip.

        Parâmetros:
        - arg_strPathFolder (str): Caminho da Pasta à ser Compactada.

        Retorno:
        - str: Caminho do Arquivo Compactado
        """          
        shutil.make_archive(base_name= arg_strPathFolder,
                            format= "zip",
                            root_dir= arg_strPathFolder)
        
        
        return shutil.move(arg_strPathFolder + '.zip', arg_strPathFolder)


    def get_Files(self, arg_strPath) -> list[str]:
        """
        Lista os arquivos no Caminho.

        Parâmetros:
        - arg_strPath (str): Caminho da Pasta à ser consultada.

        Retorno:
        - list[str]: Lista com os caminhos dos arquivos encontrados.
        """  
        var_lstPathFiles = os.listdir(path= arg_strPath)

        return [ archive for archive in var_lstPathFiles if os.path.isfile(archive)]


    def get_Files_By_ID(self, arg_strPath:str, arg_strIdentificador:str) -> list[str]:
        """
        Lista os arquivos no Caminho que contenham a string informada como parte do nome.

        Parâmetros:
        - arg_strPath (str): Caminho da Pasta à ser consultada.
        - arg_strIdentificador (str): String utilizada para identificar os arquivos a serem listados.

        Retorno:
        - list[str]: Lista com os caminhos dos arquivos encontrados.
        """  
        var_lstPathFiles = os.listdir(path= arg_strPath)
        
        return [os.path.join(arg_strPath, archive) for archive in var_lstPathFiles if arg_strIdentificador in archive ]


    def get_Folders(self, arg_strPath:str) -> list[str]: 
        """
        Lista as Pastas encontradas no Diretório.

        Parâmetros:
        - arg_strPath (str): Caminho da Pasta à ser consultada.

        Retorno:
        - list[str]: Lista com os caminhos das Pastas encontradas.
        """  
        var_lstArchive = os.listdir(arg_strPath)

        return [ os.path.join(arg_strPath, archive) for archive in var_lstArchive if os.path.isdir( os.path.join(arg_strPath, archive) )]


    def __check_Folder_Zip(self, arg_strPath:str) -> str:
        """
        Verifica se existem Pastas dentro do diretório Descompactado. Caso exista, altera o Caminho Raiz 
        do Projeto para a Pasta encontrada. Caso contrário, mantém o Diretório Descompactado como Raiz.

        Parâmetros:
        - arg_strPath (str): Caminho da Pasta à ser consultada.

        Retorno:
        - str: Caminho do Diretório Raiz.

        Raises:
        - BusinessRuleException: Mais de uma pasta encontrada no Arquivo Descompactado
        """
        var_strPathFolder:str

        var_lstFolder = self.get_Folders(arg_strPath= arg_strPath)

        if( len(var_lstFolder) == 1 ): var_strPathFolder = var_lstFolder.pop()
        elif( len(var_lstFolder) < 1 ): var_strPathFolder = arg_strPath 
        elif( len(var_lstFolder) > 1 ): raise BusinessRuleException("ERRO PASTA. Mensagem: Mais de uma pasta encontrada no Arquivo Descompactado")

        return var_strPathFolder


    def exists_Path(self, arg_strPath:str) -> bool:
        """
        Verifica se o Caminho Existe.

        Parâmetros:
        - arg_strPath (str): Caminho a ser consultado.

        Retorno:
        - bool: Booleana informando se o caminho existe.
        """       
        return os.path.exists(path= arg_strPath)


    def take_screenshot(self, arg_strPath:str) -> str:
        """
        Captura um Screenshot e salva no Caminho informado.

        Parâmetros:
        - arg_strPath (str): Caminho Pasta que será salvo o Screenshot.

        Retorno:
        - str: Caminho do Arquivo Screenshot.
        """
        var_strFileName = "ExceptionScreenshot_" + datetime.datetime.now().strftime("%d%m%Y_%H%M%S%f") + ".png"

        var_strPathFile = os.path.join( arg_strPath, var_strFileName )

        pyautogui.screenshot(imageFilename= var_strPathFile)

        return var_strPathFile


    def move_to(self, arg_strPathOrigin:str, arg_strPathDest) -> str:
        """
        Move um arquivo ou diretório para uma pasta de destino.

        Parâmetros:
        - arg_strPathOrigin (str): Caminho do arquivo/diretório que será movido.
        - arg_strPathDest (str): Caminho de destino que receberá o arquivo/diretório.

        Retorno:
        - str: Novo caminho do arquivo/diretório.
        """
        return shutil.move(arg_strPathOrigin, arg_strPathDest)


    def clear_Folder(self, arg_strPath:str) -> None:
        """
        Limpa os arquivos do diretório informado.

        Parâmetros:
        - arg_strPath (str): Caminho da pasta que será limpa.
        """
        for achive in os.listdir(arg_strPath):
            var_strAuxPath = os.path.join(arg_strPath, achive)

            if( os.path.isfile(var_strAuxPath) ):
                os.remove(var_strAuxPath)


    def read_excel_cod_darf(self) -> list[dict]:
        var_dfCodDARF = pandas.read_excel(io=ROOT_DIR + "\\resources\\templates\\COD DARF.xlsx", sheet_name='COD DARF ÚNICO', usecols="A")

        var_dfCodDARF['retry'] = '0'

        return var_dfCodDARF.to_dict('records')
