from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive as Gd
from prj_Apter_DownloadArquivos.classes_t2c.utils.T2CMaestro import T2CMaestro, LogLevel, ErrorType

class GoogleDrive:

    def __init__(self,arg_clssMaestro:T2CMaestro,arg_dictConfig:dict) -> None:
        self.var_clssMaestro = arg_clssMaestro
        self.var_dictConfig = arg_dictConfig


    def get_link(self,arg_strPathName:str):
        gauth = GoogleAuth(self.var_dictConfig['GoogleAuth']) # Caminho padrão, do arquivo com as configurações para não pedir toda vez autorização para acessar o drive.
        gauth.LoadClientConfigFile(self.var_dictConfig['GoogleSecret']) 

        drive = Gd(gauth)

        var_strPasta = (self.var_dictConfig['LinkPastaDrive'])  #Caminho da pasta Onda - 2

        # var_listfiles = drive.ListFile({'q': f"title = '{'35829290_Lesaffre do Brasil Produtos Alimenticios LTDA (copy)_04122023_1930'}' and mimeType = 'application/vnd.google-apps.folder'"}).GetList()
        var_listfiles = drive.ListFile({'q': f"title = '{arg_strPathName}' and mimeType = 'application/vnd.google-apps.folder'"}).GetList()
        
        primeiro_elemento = var_listfiles[0]

        if len(primeiro_elemento) > 0:
            link_compartilhamento = primeiro_elemento['alternateLink']
            var_LogMaestro = self.var_clssMaestro.write_log(f"Link do Googledrive {link_compartilhamento}",arg_enumLogLevel=LogLevel.INFO)

        else:
            link_compartilhamento = "Não foi gerado link do drive"

        return (link_compartilhamento)
    
        