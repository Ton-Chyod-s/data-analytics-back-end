import os
import sys
import time
import locale
import datetime
import pandas as pd
import smtplib
import email.message
from time import sleep

inicioContador = time.time()

def enviar_email(assunto, de, para, senha,corpo_mensagem):  
    corpo_email = corpo_mensagem
    
    msg = email.message.Message()
    msg['Subject'] = assunto
    msg['From'] = de
    msg['To'] = para
    password = senha
    msg.add_header('Content-Type', 'text/html')
    msg.set_payload(corpo_email )

    s = smtplib.SMTP('smtp.gmail.com: 587')
    s.starttls()
    # Login Credentials for sending the mail
    s.login(msg['From'], password)
    s.sendmail(msg['From'], [msg['To']], msg.as_string().encode('utf-8'))
    print('Email enviado')

nomeTabelaTratada = 'tabelaTratada'
nomeDaPasta = r'\2024'
anoAnterior = r'\2023'

if nomeDaPasta == anoAnterior:
    dic_trimestre = {'4º trimestre': ''}
else:
    dic_trimestre = {'1º trimestre':'01T2024',
                    '2º trimestre': '02T2024',
                    '3º trimestre': '03T2024',
                    '4º trimestre': '04T2024',}

novaTabela = []
arquivos = ""

for key, value in enumerate(dic_trimestre):
    trimestre = f'\\{dic_trimestre[value]}'
    locale.setlocale(locale.LC_ALL, 'pt_BR')

    data_atual = datetime.date.today()

    ano = data_atual.year 

    caminhoPasta = r'C:\Users\User\Desktop\Nova pasta (2)'
    #r'\\192.168.1.2\dados\SUPERMERCADO CONTABIL\Planilhas resultados - Power Bi' + nomeDaPasta + trimestre 
    arquivosPasta = os.listdir(caminhoPasta)
    caminhoTabelaTratada = f'{os.getcwd()}\\{nomeTabelaTratada}.xlsx'

    #verificando se a pasta existe e se tem arquivos dentro da mesma
    if not os.path.exists(caminhoPasta):
        os.makedirs(caminhoPasta)
    elif len(arquivosPasta) == 0:
        sys.exit()

    #loop para interagir com arquivos dentro da pasta
    for num,i in enumerate(arquivosPasta):
        variavel = None
        nomeEmpresaString = None
        nomeAba = None
        
        arquivos += f'{i},<br>'
        #pegando nome da aba da planilha
        if nomeAba == None:
            nomeAbaTabela = pd.ExcelFile(os.path.join(caminhoPasta,i)).sheet_names
            nomeAba = nomeAbaTabela[0]
        
        #tratamento das abas da planilha localizada
        for k, v in enumerate(nomeAbaTabela):
            sleep(.1)
            df = pd.read_excel(os.path.join(caminhoPasta,i), sheet_name=v)
            if nomeEmpresaString == None:
                if '-' in arquivosPasta[num]:
                    nomeEmpresaString = arquivosPasta[num].split('-')[0]
                else:
                    nomeEmpresaString = arquivosPasta[num].split('.')[0]
            
            #nome da unidade
            nomeUnidade = f'Loja {v}'

            #verificando se a aba é a empresa correta
            if 'Pires e Cia' in nomeEmpresaString:    
                break

            elif 'Fam Atacado 600 1° T 2024' in nomeEmpresaString:    
                break
            
            elif 'Fogo Atacado - 1ºT 2024' in nomeEmpresaString:    
                break
            
            elif 'Grupo Gmais - 1ºT 2024' in nomeEmpresaString:    
                break
            
            elif 'JChagas - 1ºT 2024' in nomeEmpresaString:    
                break
            
            #valores das celulas
            primeiraCelula = df.keys()[0]
            terceiraCelula = df.keys()[2]
    
            tanhoPrimeiraCelula = len(primeiraCelula)
            #se o dataframe não tiver a primeira linha em branco, vai acrescentar a primeira linha duplicada
            if tanhoPrimeiraCelula > 10:
                df.iloc[0] = df.columns
                df.columns = [f"Unnamed: {i}" for i in range(len(df.columns))]

            data_emissao = df.iloc[0,2]

            #função para remover linhas a mais
            def contador(valor,dataFrame):
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

            #verifica se classificação esta na celula correta para a continuação
            for cont in range(11):
                celula = df.iloc[cont, 0]
                if celula == 'Classificação':
                    df = contador(cont,df)
                    df = df.reset_index(drop=True)
                    break
            
            # Concatenar as duas primeiras linhas para formar o novo cabeçalho
            primeira_linha = df.iloc[0].astype(str)
            segunda_linha = df.iloc[1].astype(str)
            
            #verificando se o ano atual é o mesmo na planilha DRE do servidor
            dataDRE = data_emissao.split(' ')[1].split('/')[2]
            for key, value in enumerate(segunda_linha):
                if dataDRE in value:
                    continue
                else:
                    break
                
            # Concatenar as duas primeiras linhas para formar o novo cabeçalho
            novo_cabecalho = primeira_linha + '-' + segunda_linha
            
            #define a primeira linha como cabeçado e remove os espaço dela
            novo_cabecalho = novo_cabecalho.str.replace(' ', '').str.replace('-', ' ').str.replace('nan', '').str.replace('ç', 'c').str.replace('ã', 'a')
            df.columns = novo_cabecalho

            '''--Lista de meses--'''
            # Extrair apenas os meses usando uma expressão regular
            segunda_linha_meses = segunda_linha.str.extract(r'(\w{3})/\d{4}').squeeze()
            # Remover valores duplicados e valores 'NaN' da série e transformar em uma lista
            meses = segunda_linha_meses.drop_duplicates().dropna().tolist()

            # Filtrar colunas que contêm a substring 'Saldo', 'Débitos', 'Créditos', 'Metas/Orçam.' e '%Mt/Or'
            colunas_a_remover = df.columns[df.columns.str.contains('Saldo|Débitos|Créditos|Metas/Orcam.|%Mt/Or')]
            df = df.drop(columns=colunas_a_remover)
            
            # Redefinir o índice
            df = df.reset_index(drop=True)
            #apagando as duas primeiras linhas
            df = df.drop(range(0, 2))

            #procurando valores CR,DB e removendo de cada celula
            def clean_value(x):
                if isinstance(x, str) and 'DB' in x:
                    x = x.replace('DB', '').strip().replace(".", "").replace(",", ".")
                    x = "-" + x
                    #x = float(x) * -1
                    return x
                elif isinstance(x, str) and 'CR' in x:
                    x = x.replace('CR', '').strip().replace(".", "").replace(",", ".")
                    #x = float(x)
                    return x
                return x
            df.iloc[:, 4:] = df.iloc[:, 4:].map(clean_value)

            #econtrar valor, coluna e linha 
            def valor_da_linha(linha_descricao, coluna_a_encontrar,contem=True):
                coluna_filtro = 'DescricaoConta '
                try:
                    # Tratar valores ausentes substituindo por uma string vazia
                    df[coluna_filtro] = df[coluna_filtro].fillna('')
                    if contem:
                        #encontrar linha a partir de uma coluna como filtro e contendo os valores
                        linha_contas_resultado = df[df[coluna_filtro].str.contains(linha_descricao, case=False)]
                    else:
                        # Localizar a célula que contém o texto inserido
                        linha_contas_resultado = df[df[coluna_filtro].str.strip() == linha_descricao]
                    try:
                        # Obter o valor da coluna inserido
                        valor_linha = linha_contas_resultado[coluna_a_encontrar].values[0]
                        indice_linha = linha_contas_resultado.index[linha_contas_resultado[coluna_a_encontrar] == valor_linha].tolist()[0]
                        num_coluna = df.columns.get_loc(coluna_a_encontrar)
                        valor_linha = linha_contas_resultado[coluna_a_encontrar].str.strip().values[0]
                    except:
                        # Obter o valor da coluna inserido
                        valor_linha = linha_contas_resultado[coluna_a_encontrar].values[0]
                        indice_linha = linha_contas_resultado.index[0]
                        num_coluna = df.columns.get_loc(coluna_a_encontrar)
                        
                    if valor_linha == '':
                        valor_linha = 0
                    else:
                        valor_linha = float(linha_contas_resultado[coluna_a_encontrar].values[0])
                    return valor_linha, indice_linha , num_coluna
                except Exception as e:
                    return 

            #atualizando o dicionario
            def att(dataset,resultado):
                for item in dataset.items():
                    indice1_dicionario = item[:2]
                    if indice1_dicionario[1] == dataset[coluna_a_encontrar]: 
                        dataset.update({ item[0]: resultado })
                        break
            
            #adicionando novas linhas para preenchimento
            linhas = [f'{data_emissao}','VENDA','LUCRO CONTABIL','RECOLHIDO','RECEITAS NAO OPERACIONAIS','LUCRO',' ','IR','ADD' ,'CSLL','TOTAL',' ', 'TOTAL IR-ADD',' ', 'MARGEM CONTABIL','RESULTADO LIQUIDO',' ']

            for lin, linha in enumerate(linhas):
                lin_ha = lin + (len(df) + 2)
                df.loc[lin_ha, 'DescricaoConta '] = linha
                df.loc[lin_ha, 'Classificacao '] = '1.0'
                df.loc[lin_ha, 'Conta '] = '10'

            # Redefinir o índice
            df = df.reset_index(drop=True)

            '''--preencher dicionario--'''
            for mes in meses:
                #definindo o ano e mes para procurar em relação ao função mês
                coluna_a_encontrar = f'MvtoLíquido {mes}/{ano}'
                
                def preencher_dicionario(linha, dicionario, dre, contem_linha=True):
                    try:
                        #encontrar valor, coluna e linha
                        valor_linha, indice_linha, num_coluna = valor_da_linha(linha, coluna_a_encontrar, contem_linha)
                        #verificar se a chave existe no dicionario
                        if dicionario in dre:
                            pass
                        else:
                            dre[dicionario] = {}
                        novoDicionario = dre[dicionario]
                        #adicionando valores no dicionario interno
                        novoDicionario[f'MvtoLíquido {mes}/{ano}'] = valor_linha
                        novoDicionario[f'Index_linha {mes}/{ano}'] = indice_linha
                        novoDicionario[f'Index_coluna {mes}/{ano}'] = num_coluna
                    except:
                        print(f'\nError when filling the {linha} out in the: ',nomeEmpresaString)

                # exemplo para preencher dicionario
                # preencher_dicionario('CONTAS DE RESULTADO','contasResultado', drePadrao)