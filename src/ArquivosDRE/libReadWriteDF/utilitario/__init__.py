import os
import sys
import pandas as pd

def dicTrimestre(anoAnterior, nomeDaPasta):
    """
    Retorna um dicionário com os trimestres do ano, baseado no nome da pasta.

    Args:
        anoAnterior (str): Nome do ano anterior (exemplo: '2023').
        nomeDaPasta (str): Nome da pasta a ser verificado.

    Returns:
        dict: Dicionário com os trimestres e seus respectivos códigos (exemplo: {'1º trimestre': '01T2024', ...}).
    """
    if nomeDaPasta == anoAnterior:
        dic_trimestre = {'4º trimestre': ''}
    else:
        dic_trimestre = {'1º trimestre':'01T2024',
                        '2º trimestre': '02T2024',
                        '3º trimestre': '03T2024',
                        '4º trimestre': '04T2024',}
    return dic_trimestre


def verificarPasta(arquivosPasta, caminhoPasta):
    """
    Verifica a existência da pasta no caminho especificado e cria-a se não existir.

    Args:
        arquivosPasta (list): Lista de arquivos na pasta.
        caminhoPasta (str): Caminho da pasta a ser verificada.
    """
    if not os.path.exists(caminhoPasta):
        os.makedirs(caminhoPasta)
    elif len(arquivosPasta) == 0:
        sys.exit()


def NomeAba(i, caminhoPasta, nomeAba):
    """
    Retorna o nome da aba da planilha.

    Args:
        i (str): Nome do arquivo.
        caminhoPasta (str): Caminho da pasta.

    Returns:
        str: Nome da aba da planilha.
    """
    if nomeAba == None:
        nomeAbaTabela = pd.ExcelFile(os.path.join(caminhoPasta,i)).sheet_names
    return nomeAbaTabela[0]


def NomeEmpresa(arquivosPasta, nomeEmpresaString):
    """
    Retorna o nome da empresa.

    Args:
        arquivosPasta (list): Lista de arquivos na pasta.
        nomeEmpresaString (str): Nome da empresa.

    Returns:
        str: Nome da empresa.
    """
    for num in range(len(arquivosPasta)):
        if nomeEmpresaString == None:
            if '-' in arquivosPasta[num]:
                nomeEmpresaString = arquivosPasta[num].split('-')[0]
            else:
                nomeEmpresaString = arquivosPasta[num].split('.')[0]
    return nomeEmpresaString


def VerificarAba(nomeEmpresaString, *kwargs):
    """
    Verifica se a aba é a empresa correta.

    Args:
        nomeEmpresaString (str): Nome da empresa.

    Exemplo:
        nomeEmpresaString = 'JChagas - 1ºT 2024'
        nomepular = [
            'JChagas',
            'Fogo Atacado',
            'VRA'
        ]
    """
    for k, v in enumerate(kwargs):
        if v[k] in nomeEmpresaString:
            return sys.exit()


def tamanhoPrimeiraCelula(celula, df):
    """
    Retorna o tamanho da primeira célula do dataframe.

    Args:
        df (DataFrame): Dataframe a ser verificado.

    Returns:
        int: Tamanho da primeira célula.

    Exemplo:
        df = pd.read_excel(r'C:\Users\User\Desktop\Nova pasta (2)\Fogo Atacado - 1ºT 2024.xlsx')
        primeiraCelula = df.keys()[0]
        tamanhoPrimeiraCelula(primeiraCelula, df)
    """
    tipoCelula = type(celula)
    if tipoCelula == float or tipoCelula == str:
        df.iloc[0] = df.columns
        df.columns = [f"Unnamed: {i}" for i in range(len(df.columns))]

def verificarLinhasCabecalho(df):
    """
    Verifica e remove linhas extras do início do dataframe.

    Args:
        df (DataFrame): O dataframe a ser verificado.

    Returns:
        DataFrame: O dataframe com as linhas extras removidas.
    """
    #função para remover linhas a mais
    def contador(valor,dataFrame):
        """
        
        """
        match valor:
            case 0:
                print('opa deu um erro nessa condição')
            case 1:
                dataFrame.drop(range(0, 1), inplace=True)
                dataFrame.reset_index(drop=True)
                return dataFrame
            case 2:
                dataFrame.drop(range(0, 2), inplace=True)
                dataFrame.reset_index(drop=True)
                return dataFrame  
            case 3:
                dataFrame.drop(range(0, 3), inplace=True)
                dataFrame.reset_index(drop=True)
                return dataFrame
            case 4:
                dataFrame.drop(range(0, 4), inplace=True)
                dataFrame.reset_index(drop=True)
                return dataFrame
            case 5:
                dataFrame.drop(range(0, 5), inplace=True)
                dataFrame.reset_index(drop=True)
                return dataFrame
            case 6:
                dataFrame.drop(range(0, 6), inplace=True)
                dataFrame.reset_index(drop=True)
                return dataFrame
            case 7:
                dataFrame.drop(range(0, 7), inplace=True)
                dataFrame.reset_index(drop=True)
                return dataFrame
            case 8:
                dataFrame.drop(range(0, 8), inplace=True)
                dataFrame.reset_index(drop=True)
                return dataFrame
            case 9:
                dataFrame.drop(range(0, 9), inplace=True)
                dataFrame.reset_index(drop=True)
                return dataFrame
            case 10:
                dataFrame.drop(range(0, 10), inplace=True)
                dataFrame.reset_index(drop=True)
                return dataFrame