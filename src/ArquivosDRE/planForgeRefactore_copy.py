import os
import re
import sys
import time
import locale
import pandas as pd
from time import sleep
from datetime import datetime
# from dotenv import load_dotenv

# load_dotenv()

data_atual = datetime.now()
anoAtual = data_atual.year  

def forgePlan(caminho, ondeSalvar, ano_analisado, semestre_analisado):
    try:
        inicioContador = time.time()
        tempoExecucao = 0
        nomeTabelaTratada = 'tabelaTratada'
        nomeDaPasta = f'\\{ano_analisado}'
        continuar = False
        calculoTrimestre = None
     
        dic_trimestre = {'1º trimestre':f'01T{ano_analisado}',
                        '2º trimestre': f'02T{ano_analisado}',
                        '3º trimestre': f'03T{ano_analisado}',
                        '4º trimestre': f'04T{ano_analisado}'}
        
        novaTabela = []
        arquivos = ""
        anoAlisado = nomeDaPasta.replace('\\','')
        semestreAnalisado = f'0{semestre_analisado}T{anoAlisado}'
        
        #loop para interagir com os trimestres
        for key, value in enumerate(dic_trimestre):
            semestre = dic_trimestre[value]
            if semestre != semestreAnalisado:
                semestre = semestreAnalisado
            trimestre = f'\\{semestre}'
            locale.setlocale(locale.LC_ALL, 'pt_BR')

            data_atual = datetime.now()

            ano = data_atual.year 

            if str(ano) not in nomeDaPasta:
                ano = int(nomeDaPasta.replace('\\',''))

            caminhoPasta = caminho + nomeDaPasta + trimestre 
            arquivosPasta = os.listdir(caminhoPasta)

            #verificando se a pasta existe e se tem arquivos dentro da mesma
            if not os.path.exists(caminhoPasta):
                os.makedirs(caminhoPasta)
            elif len(arquivosPasta) == 0:
                sys.exit()

            #loop para interagir com arquivos dentro da pasta
            for num,i in enumerate(arquivosPasta):
                tipoDRE = 0
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
                    #verificando se a aba da empresa esta vazia
                    if df.empty:
                        break

                    if nomeEmpresaString == None:
                        if '-' in arquivosPasta[num]:
                            nomeEmpresaString = arquivosPasta[num].split('-')[0]
                        else:
                            nomeEmpresaString = arquivosPasta[num].split('.')[0]
                    
                    #nome da unidade
                    nomeUnidade = f'Loja {v}'
                    
                    #verificando se a aba é a empresa correta
                    if 'Pires e Cia' in nomeEmpresaString:
                        tipoDRE = 2    
                    elif 'Fam Atacado' in nomeEmpresaString:
                        tipoDRE = 2   
                    elif 'VRA ' in nomeEmpresaString:
                        tipoDRE = 1    
                    elif 'Fogo Atacado' in nomeEmpresaString:
                        tipoDRE = 1    
                    elif 'JChagas' in nomeEmpresaString:
                        tipoDRE = 1   
                    elif 'Bonanca' in nomeEmpresaString:
                        tipoDRE = 3   
                    elif 'S Pires' in nomeEmpresaString:
                        tipoDRE = 3   
                        
                    #valores das celulas
                    primeiraCelula = df.keys()[0]
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
                    
                    def formatar_data(data):
                        try:
                            data_obj = datetime.strptime(data, '%Y-%m-%d %H:%M:%S')
                            return data_obj.strftime('%b/%Y').capitalize()  # Formato 'Mês/Ano'
                        except ValueError:
                            return data  # Retorna a data original se não for no formato esperado

                    # Concatenar as duas primeiras linhas para formar o novo cabeçalho
                    primeira_linha = df.iloc[0].astype(str)
                    segunda_linha = df.iloc[1].astype(str)

                    # Expressão regular para corresponder ao formato de data e hora '%Y-%m-%d %H:%M:%S'
                    pattern = re.compile(r'\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}')

                    # Verifica se algum elemento de segunda_linha corresponde ao padrão
                    temNaLinha = any(pattern.match(i) for i in segunda_linha)

                    if temNaLinha:
                        # Filtra apenas os elementos no formato 'Mês/Ano'
                        segunda_linha = [formatar_data(item) for item in segunda_linha]

                    #verificando se o ano atual é o mesmo na planilha DRE do servidor
                    try:
                        dataDRE = data_emissao.split(' ')[1].split('/')[2]
                    except:
                        dataDRE = ano.split('0')[1]
                    
                    for key, value in enumerate(segunda_linha):
                        if dataDRE in value:
                            continuar = True
                            break

                    if continuar:
                        pass
                    else:
                        break
                        
                    # Concatenar as duas primeiras linhas para formar o novo cabeçalho
                    novo_cabecalho = primeira_linha + '-' + segunda_linha
                    
                    #define a primeira linha como cabeçado e remove os espaço dela
                    novo_cabecalho = novo_cabecalho.str.replace(' ', '').str.replace('-', ' ').str.replace('nan', '').str.replace('ç', 'c').str.replace('ã', 'a')
                    df.columns = novo_cabecalho

                    '''--Lista de meses--'''
                    # Extrair apenas os meses usando uma expressão regular
                    if type(segunda_linha) == list:
                        segunda_linha_sem_duplicatas = list(set(segunda_linha))
                        meses = [x.split('/')[0].capitalize() for x in segunda_linha_sem_duplicatas if x != 'nan' and '/' in x]

                    else:
                        segunda_linha_meses = segunda_linha.str.extract(r'(\w{3})/\d{4}').squeeze()
                        # Remover valores duplicados e valores 'NaN' da série e transformar em uma lista
                        meses = segunda_linha_meses.drop_duplicates().dropna().tolist()

                    meses.sort(key=lambda x: datetime.strptime(x, '%b'))
                    
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

                    '''Inicio da procura e calculo necessario ----------------------------------------------------------------------------------'''
                    # Inicialização de dicionários vazios
                    #Planilha DRE
                    drePadrao = {}
                    dreAtipico = {}
                    dreTipo2 = {}	
                    dreTipo3 = {}

                    #Trimestre
                    dreTrimestre = {}
                    
                    def limparDicionarios():
                        drePadrao.clear()
                        dreAtipico.clear()
                        dreTrimestre.clear()

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
                            return e

                    #atualizando o dicionario
                    def att(dataset,resultado):
                        for item in dataset.items():
                            indice1_dicionario = item[:2]
                            if indice1_dicionario[1] == dataset[coluna_a_encontrar] or tipoDRE == 2: 
                                dataset.update({ item[0]: resultado })
                                break
                    
                    #adicionando novas linhas para preenchimento
                    linhas = [f'{data_emissao}','VENDA','LUCRO CONTABIL','RECOLHIDO','RECEITAS NAO OPERACIONAIS','LUCRO',' ','IR','ADD' ,'CSLL','TOTAL',' ', 'TOTAL IR-ADD',' ', 'MARGEM CONTABIL','RESULTADO LIQUIDO','EBITDA',' ']

                    for lin, linha in enumerate(linhas):
                        lin_ha = lin + (len(df) + 2)
                        df.loc[lin_ha, 'DescricaoConta '] = linha
                        df.loc[lin_ha, 'Classificacao '] = '1.0'
                        df.loc[lin_ha, 'Conta '] = '10'

                    # Redefinir o índice
                    df = df.reset_index(drop=True)

                    '''--preencher dicionario--'''
                    for conta, mes in enumerate(meses):
                        #definindo o ano e mes para procurar em relação ao função mês
                        coluna_a_encontrar = f'MvtoLíquido {mes}/{ano}'
                        
                        def preencher_dicionario(linha, dicionario, dre, contem_linha=True):
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
                    
                        if tipoDRE == 0:
                            #atualização conta de resultado (contas de resultado - provisão de imposto)
                            try:      
                                preencher_dicionario('CONTAS DE RESULTADO','contas_resultado', drePadrao)
                            except Exception as e:
                                if 'cannot unpack non-iterable IndexError object' in str(e):
                                    try:
                                        preencher_dicionario('Contas de Resultado','contas_resultado', drePadrao)
                                    except Exception as e:
                                        if 'cannot unpack non-iterable IndexError object' in str(e):
                                            try:
                                                preencher_dicionario('R E S U L T A D O','contas_resultado', drePadrao)
                                            except Exception as e:
                                                print("An unexpected error occurred:", e)
                            try:
                                preencher_dicionario('PROVISAO DE IMPOSTO S/L', 'provisao_de_imposto', drePadrao)
                            except Exception as e:
                                if 'cannot unpack non-iterable IndexError object' in str(e):
                                    try:
                                        preencher_dicionario('PROVISAO DE IMPOSTO S/LUCRO', 'provisao_de_imposto', drePadrao)
                                    except Exception as e:
                                        try:
                                            preencher_dicionario('Provisão de Imposto S/L', 'provisao_de_imposto', drePadrao)
                                        except Exception as e:
                                            print("An unexpected error occurred:", e)
                            try:
                                preencher_dicionario('RECEITAS OPERACIONAL LIQUIDA', 'receitas_operacaional', drePadrao)
                            except Exception as e:
                                print(e)

                            #contas de resultado calculo e atualização na planilha                 
                            resultado_analisado = drePadrao['contas_resultado'][coluna_a_encontrar] + drePadrao['provisao_de_imposto'][coluna_a_encontrar] * - 1
                            att(drePadrao['contas_resultado'],resultado_analisado)
                            
                            #calculo resultado liquido (contas de resultado - Recuperacao De Despesas Exerc Anterior)
                            preencher_dicionario('Recuperacao De Despesas Exerc Anterior','recuperacao_despesas', drePadrao)
                            preencher_dicionario('RESULTADO LIQUIDO','resultado_liquido', drePadrao)

                            calc_res_liquido = drePadrao['contas_resultado'][f'MvtoLíquido {mes}/{ano}'] - drePadrao['recuperacao_despesas'][f'MvtoLíquido {mes}/{ano}']
                            df.at[drePadrao['resultado_liquido'][f'Index_linha {mes}/{ano}'], f'MvtoLíquido {mes}/{ano}'] = calc_res_liquido

                            #calculo margem contabil (receitas op liquida - CVM custo de mercadorias)
                            try:
                                preencher_dicionario('CUSTO DAS MERCADORIAS VENDIDOS - CMV', 'custo_mercadorias', drePadrao)
                            except Exception as e:
                                if 'cannot unpack non-iterable IndexError object' in str(e):
                                    try:
                                        preencher_dicionario('CUSTOS DAS MERCADORIAS VENDIDOS', 'custo_mercadorias', drePadrao, False)
                                    except Exception as e:
                                        if 'cannot unpack non-iterable IndexError object' in str(e):
                                            try:
                                                preencher_dicionario('Custos Das Mercadorias e Serviços Vendidos', 'custo_mercadorias', drePadrao, False)
                                            except Exception as e:
                                                print("An unexpected error occurred:", e)

                            preencher_dicionario('MARGEM CONTABIL','margem_contabil', drePadrao)
                            calc_margem_cont = drePadrao['receitas_operacaional'][f'MvtoLíquido {mes}/{ano}'] + drePadrao['custo_mercadorias'][f'MvtoLíquido {mes}/{ano}']
                            df.at[drePadrao['margem_contabil'][f'Index_linha {mes}/{ano}'], f'MvtoLíquido {mes}/{ano}'] = calc_margem_cont

                            '''--dict trimestre--'''
                            #calculo venda /trimestre
                            preencher_dicionario('VENDA DE MERCADORIAS','venda_mercadorias', dreTrimestre)
                            preencher_dicionario('VENDA','venda', dreTrimestre, False)
                            #adicionando na planilha
                            df.at[dreTrimestre['venda'][f'Index_linha {mes}/{ano}'], f'MvtoLíquido {mes}/{ano}'] = dreTrimestre['venda_mercadorias'][f'MvtoLíquido {mes}/{ano}']
                            
                            #lucro contabil == soma de conta de resultados
                            preencher_dicionario('LUCRO CONTABIL','lucro_contabil', dreTrimestre, False)
                            lucroContabil = drePadrao['contas_resultado'][f'MvtoLíquido {mes}/{ano}']
                            df.at[dreTrimestre['lucro_contabil'][f'Index_linha {mes}/{ano}'], f'MvtoLíquido {mes}/{ano}'] = lucroContabil

                            #calculo recolhido /trimestre
                            preencher_dicionario('RECOLHIDO','recolhido',dreTrimestre, False)
                            
                            try:
                                preencher_dicionario('Contribuicao Social', 'contribuicao_social', dreTrimestre, False)
                            except Exception as e:
                                if 'cannot unpack non-iterable IndexError object' in str(e):
                                    try:
                                        preencher_dicionario('Contribuição Social', 'contribuicao_social', dreTrimestre, False)
                                    except Exception as e:
                                        if 'cannot unpack non-iterable IndexError object' in str(e):
                                            try:
                                                preencher_dicionario('Contribuicao Sindical', 'contribuicao_social', dreTrimestre, False)
                                            except Exception as e:
                                                print("An unexpected error occurred:", e)

                            comp = round(dreTrimestre['contribuicao_social'][f'MvtoLíquido {mes}/{ano}'] / 0.9  * 10, 2)
                            df.at[dreTrimestre['recolhido'][f'Index_linha {mes}/{ano}'], f'MvtoLíquido {mes}/{ano}'] = comp
                            
                            #calculo Receitas Não Operacionais 
                            preencher_dicionario('RECEITAS NAO OPERACIONAIS','receita_nao_operacional', dreTrimestre, False)
                            df.at[dreTrimestre['receita_nao_operacional'][f'Index_linha {mes}/{ano}'], f'MvtoLíquido {mes}/{ano}'] = drePadrao['recuperacao_despesas'][f'MvtoLíquido {mes}/{ano}']

                            #lucro
                            preencher_dicionario('LUCRO','lucro', dreTrimestre, False)
                            lucro_res = drePadrao['contas_resultado'][coluna_a_encontrar] - drePadrao['recuperacao_despesas'][coluna_a_encontrar]
                            dreTrimestre['lucro'][f'MvtoLíquido {mes}/{ano}'] = lucro_res
                            #adicionando na planilha
                            df.at[dreTrimestre['lucro'][f'Index_linha {mes}/{ano}'], f'MvtoLíquido {mes}/{ano}'] = dreTrimestre['lucro'][f'MvtoLíquido {mes}/{ano}']
                            
                            #ebitda
                            preencher_dicionario('DEPRECIACOES E PROVISOES','drepreciacaoProvisoes',drePadrao)
                            preencher_dicionario('DESPESAS FINANCEIRAS','despesasFinaceiras',drePadrao)
                            preencher_dicionario('Imposto De Renda','impostoRenda',drePadrao)

                            preencher_dicionario('IR','imposto', dreTrimestre,False)
                            preencher_dicionario('ADD','add', dreTrimestre)
                            preencher_dicionario('CSLL','csll',dreTrimestre)
                            preencher_dicionario('TOTAL','total', dreTrimestre)
                            preencher_dicionario('TOTAL IR-ADD','totalIrAdd',dreTrimestre)
                            preencher_dicionario('EBITDA','ebitda',dreTrimestre)

                            lucro = dreTrimestre['lucro'][f'MvtoLíquido {mes}/{ano}']
                            contribuicaoSocial = dreTrimestre['contribuicao_social'][f'MvtoLíquido {mes}/{ano}']
                            depreciacaoProvisoes = drePadrao['drepreciacaoProvisoes'][f'MvtoLíquido {mes}/{ano}']
                            despesasFinaceiras = drePadrao['despesasFinaceiras'][f'MvtoLíquido {mes}/{ano}']
                            impostoRenda = drePadrao['impostoRenda'][f'MvtoLíquido {mes}/{ano}']
                            
                            df.at[dreTrimestre['ebitda'][f'Index_linha {mes}/{ano}'], f'MvtoLíquido {mes}/{ano}'] = lucro + contribuicaoSocial + depreciacaoProvisoes + despesasFinaceiras + impostoRenda

                            calculoTrimestre = drePadrao

                            if conta + 1 == len(meses):
                                break

                        elif tipoDRE == 1:
                            #contas de resultado
                            preencher_dicionario('CONTAS DE RESULTADO','contas_resultado', dreAtipico)
                            try:
                                preencher_dicionario('Outras Receitas - Não Tributáveis','outras_receitas', dreAtipico)
                            except Exception as e:
                                if 'cannot unpack non-iterable IndexError object' in str(e):
                                    try:
                                        preencher_dicionario('Outras Receitas Não Operacionais','outras_receitas', dreAtipico)
                                    except Exception as e:
                                        print("An unexpected error occurred:", e)
                            
                            preencher_dicionario('PROVISAO DE IMPOSTO S/L', 'provisao_de_imposto', dreAtipico)
                            preencher_dicionario('Multas Fiscais - Não Dedutíveis', 'multas_fiscais', dreAtipico)

                            calc_contas_result = dreAtipico['contas_resultado'][f'MvtoLíquido {mes}/{ano}'] - dreAtipico['outras_receitas'][f'MvtoLíquido {mes}/{ano}'] - dreAtipico['provisao_de_imposto'][f'MvtoLíquido {mes}/{ano}'] - dreAtipico['multas_fiscais'][f'MvtoLíquido {mes}/{ano}']
                            
                            df.at[dreAtipico['contas_resultado'][f'Index_linha {mes}/{ano}'], f'MvtoLíquido {mes}/{ano}'] = calc_contas_result
                            att(dreAtipico['contas_resultado'],calc_contas_result)

                            #resultado liquido
                            preencher_dicionario('RESULTADO LIQUIDO','resultado_liquido', dreAtipico)
                            preencher_dicionario('Outras Receitas Não Operacionais','outras_receitas_nao_operacionais', dreAtipico)
                            calc_resultado_liquido = dreAtipico['contas_resultado'][f'MvtoLíquido {mes}/{ano}'] - dreAtipico['outras_receitas_nao_operacionais'][f'MvtoLíquido {mes}/{ano}']
                            df.at[dreAtipico['resultado_liquido'][f'Index_linha {mes}/{ano}'], f'MvtoLíquido {mes}/{ano}'] = calc_resultado_liquido
                            
                            #Trimestre
                            #venda
                            preencher_dicionario('VENDA','venda', dreTrimestre, False)
                            preencher_dicionario('VENDA BRUTA DE MERCADORIAS','venda_brutal_mercadorias', dreTrimestre)
                            df.at[dreTrimestre['venda'][f'Index_linha {mes}/{ano}'], f'MvtoLíquido {mes}/{ano}'] = dreTrimestre['venda_brutal_mercadorias'][f'MvtoLíquido {mes}/{ano}']

                            #lucro contabil
                            preencher_dicionario('LUCRO CONTABIL','lucro_contabil', dreTrimestre, False)
                            lucroContabil = dreAtipico['contas_resultado'][f'MvtoLíquido {mes}/{ano}']
                            df.at[dreTrimestre['lucro_contabil'][f'Index_linha {mes}/{ano}'], f'MvtoLíquido {mes}/{ano}'] = lucroContabil

                            #calculo recolhido
                            preencher_dicionario('RECOLHIDO','recolhido', dreTrimestre, False)
                            try:
                                preencher_dicionario('Contribuição Social','contribuicao_social', dreTrimestre)
                            except Exception as e:
                                if 'cannot unpack non-iterable IndexError object' in str(e):
                                    try:
                                        preencher_dicionario('Contribuicao Social','contribuicao_social', dreTrimestre)
                                    except Exception as e:
                                        print("An unexpected error occurred:", e)
                            
                            comp = round(dreTrimestre['contribuicao_social'][f'MvtoLíquido {mes}/{ano}'] / 0.9  * 10, 2)
                            df.at[dreTrimestre['recolhido'][f'Index_linha {mes}/{ano}'], f'MvtoLíquido {mes}/{ano}'] = comp

                            #calculo Receitas Não Operacionais 
                            preencher_dicionario('RECEITAS NAO OPERACIONAIS','receita_nao_operacional', dreAtipico, False)
                            df.at[dreAtipico['receita_nao_operacional'][f'Index_linha {mes}/{ano}'], f'MvtoLíquido {mes}/{ano}'] = dreAtipico['outras_receitas_nao_operacionais'][f'MvtoLíquido {mes}/{ano}']
                            
                            #lucro
                            preencher_dicionario('LUCRO','lucro', dreAtipico, False)
                            lucro_res = lucroContabil - dreAtipico['receita_nao_operacional'][f'MvtoLíquido {mes}/{ano}']
                            #adicionando na planilha
                            df.at[dreAtipico['lucro'][f'Index_linha {mes}/{ano}'], f'MvtoLíquido {mes}/{ano}'] = lucroContabil - dreAtipico['outras_receitas_nao_operacionais'][f'MvtoLíquido {mes}/{ano}']

                            preencher_dicionario('IR','imposto', dreTrimestre,False)
                            preencher_dicionario('ADD','add', dreTrimestre)
                            preencher_dicionario('CSLL','csll',dreTrimestre)
                            preencher_dicionario('TOTAL','total', dreTrimestre)
                            preencher_dicionario('TOTAL IR-ADD','totalIrAdd',dreTrimestre)

                            calculoTrimestre = dreAtipico

                        elif tipoDRE == 2:
                            preencher_dicionario('RESULTADO LIQUIDO','resultadoLiquido', dreTrimestre)
                            preencher_dicionario('MARGEM CONTABIL','margemContabil', dreTrimestre)

                            # resultado liquido
                            preencher_dicionario('R E S U L T A D O','resultadoDRE', dreTipo2, False)
                            df.at[dreTrimestre['resultadoLiquido'][f'Index_linha {mes}/{ano}'], f'MvtoLíquido {mes}/{ano}'] = dreTipo2['resultadoDRE'][coluna_a_encontrar]

                            #margin contabil
                            preencher_dicionario('RECEITAS LÍQUIDAS OPERACIONAIS','receitasOperacionais', dreTipo2)
                            preencher_dicionario('CUSTO DA MERCADORIA VENDIDA','CVM', dreTipo2)
                            calc_resultado = dreTipo2['receitasOperacionais'][coluna_a_encontrar] - dreTipo2['CVM'][coluna_a_encontrar]
                            df.at[dreTrimestre['margemContabil'][f'Index_linha {mes}/{ano}'], f'MvtoLíquido {mes}/{ano}'] = calc_resultado

                            ## trimestre ##
                            preencher_dicionario('VENDA','venda', dreTrimestre, False)
                            preencher_dicionario('LUCRO CONTABIL','lucroContabil', dreTrimestre)
                            preencher_dicionario('RECOLHIDO','recolhido', dreTrimestre, False)  
                            preencher_dicionario('RECEITAS NAO OPERACIONAIS','receitaNaoOperacional', dreTrimestre, False)
                            preencher_dicionario('LUCRO','lucro', dreTrimestre, False)
                            preencher_dicionario('IR','imposto', dreTrimestre, False)
                            preencher_dicionario('ADD','add', dreTrimestre, False)
                            preencher_dicionario('CSLL','csll', dreTrimestre, False)
                            preencher_dicionario('TOTAL','total', dreTrimestre, False)
                            preencher_dicionario('TOTAL IR-ADD','totalIrAdd', dreTrimestre, False)
                            
                            #venda
                            preencher_dicionario('RECEITAS OPERACIONAIS','receitasOperacionais', dreTipo2, False)
                            df.at[dreTrimestre['venda'][f'Index_linha {mes}/{ano}'], f'MvtoLíquido {mes}/{ano}'] = dreTipo2['receitasOperacionais'][coluna_a_encontrar]
                            #lucro contabil
                            df.at[dreTrimestre['lucroContabil'][f'Index_linha {mes}/{ano}'], f'MvtoLíquido {mes}/{ano}'] = dreTipo2['resultadoDRE'][coluna_a_encontrar]
                            #recolhido
                            preencher_dicionario('Contribuição Social','contribuicaoSocial', dreTrimestre, False)
                            df.at[dreTrimestre['recolhido'][f'Index_linha {mes}/{ano}'], f'MvtoLíquido {mes}/{ano}'] = round(dreTrimestre['contribuicaoSocial'][coluna_a_encontrar] / 0.9 * 10)
                            #receitas nao operacionais
                            try:
                                preencher_dicionario('Outras Receitas Nao Operacionais','outrasReceitas', dreTipo2, False)
                            except Exception as e:
                                if 'cannot unpack non-iterable IndexError object' in str(e):
                                    try:
                                        preencher_dicionario('Outras Receitas Não Operacionais','outrasReceitas', dreTipo2, False)
                                    except Exception as e:
                                        print("An unexpected error occurred:", e)   
                            df.at[dreTrimestre['receitaNaoOperacional'][f'Index_linha {mes}/{ano}'], f'MvtoLíquido {mes}/{ano}'] = dreTipo2['outrasReceitas'][coluna_a_encontrar]
                            #lucro
                            df.at[dreTrimestre['lucro'][f'Index_linha {mes}/{ano}'], f'MvtoLíquido {mes}/{ano}'] = dreTipo2['resultadoDRE'][coluna_a_encontrar] - dreTipo2['outrasReceitas'][coluna_a_encontrar]

                            calculoTrimestre = dreTipo2

                        elif tipoDRE == 3:
                            preencher_dicionario('RESULTADO LIQUIDO','resultadoLiquido', dreTrimestre)
                            preencher_dicionario('R E S U L T A D O','resultadoDRE', dreTipo3, False)
                            preencher_dicionario('MARGEM CONTABIL','margemContabil', dreTrimestre)
                            preencher_dicionario('PROVISÃO PARA CSLL E IRPJ','provisaoCsllIrpj', dreTipo3, False)
                            preencher_dicionario('CUSTO DA MERCADORIA VENDIDA','custoMercadoriaVendida', dreTipo3, False)

                            calcResultado = dreTipo3['resultadoDRE'][coluna_a_encontrar] + dreTipo3['provisaoCsllIrpj'][coluna_a_encontrar]
                            #resultado liquido
                            df.at[dreTrimestre['resultadoLiquido'][f'Index_linha {mes}/{ano}'], f'MvtoLíquido {mes}/{ano}'] = calcResultado

                            #margin contabil
                            preencher_dicionario('RECEITAS LÍQUIDAS OPERACIONAIS','receitasOperacionais', dreTipo3)

                            calcMargem = dreTipo3['receitasOperacionais'][coluna_a_encontrar] + dreTipo3['custoMercadoriaVendida'][coluna_a_encontrar]

                            df.at[dreTrimestre['margemContabil'][f'Index_linha {mes}/{ano}'], f'MvtoLíquido {mes}/{ano}'] = calcMargem
                            
                            ## trimestre ##
                            preencher_dicionario('VENDA','venda', dreTrimestre, False)
                            preencher_dicionario('LUCRO CONTABIL','lucroContabil', dreTrimestre)
                            preencher_dicionario('RECOLHIDO','recolhido', dreTrimestre, False)
                            preencher_dicionario('RECEITAS NAO OPERACIONAIS','receitaNaoOperacional', dreTrimestre, False)
                            preencher_dicionario('LUCRO','lucro', dreTrimestre, False)

                            #venda
                            df.at[dreTrimestre['venda'][f'Index_linha {mes}/{ano}'], f'MvtoLíquido {mes}/{ano}'] = dreTipo3['receitasOperacionais'][coluna_a_encontrar]
                            #lucro
                            df.at[dreTrimestre['lucroContabil'][f'Index_linha {mes}/{ano}'], f'MvtoLíquido {mes}/{ano}'] = dreTipo3['resultadoDRE'][coluna_a_encontrar]

                            #recollhido
                            try:
                                preencher_dicionario('Contribuição Social','contribuicaoSocial', dreTrimestre, False)
                            except Exception as e:
                                if 'cannot unpack non-iterable IndexError object' in str(e):
                                    try:
                                        preencher_dicionario('Contribuicao Social','contribuicaoSocial', dreTrimestre, False)
                                    except Exception as e:
                                        print("An unexpected error occurred:", e)
                            df.at[dreTrimestre['recolhido'][f'Index_linha {mes}/{ano}'], f'MvtoLíquido {mes}/{ano}'] = round(dreTrimestre['contribuicaoSocial'][coluna_a_encontrar] / 0.9 * 10)

                            #receitas nao operacionais
                            try:
                                preencher_dicionario('Outras Receitas Não Operacionais','outrasReceitas', dreTipo3, False)
                            except Exception as e:
                                if 'cannot unpack non-iterable IndexError object' in str(e):
                                    try:
                                        preencher_dicionario('Outras Receitas Não Operacionais','outrasReceitas', dreTipo3, False)
                                    except Exception as e:
                                        pass
                            #lucro
                            df.at[dreTrimestre['lucro'][f'Index_linha {mes}/{ano}'], f'MvtoLíquido {mes}/{ano}'] = dreTipo3['resultadoDRE'][coluna_a_encontrar] 
                            
                            preencher_dicionario('IR','imposto', dreTrimestre,False)
                            preencher_dicionario('ADD','add', dreTrimestre, False)
                            preencher_dicionario('CSLL','csll',dreTrimestre, False)
                            preencher_dicionario('TOTAL','total', dreTrimestre, False)
                            preencher_dicionario('TOTAL IR-ADD','totalIrAdd',dreTrimestre, False)
                            calculoTrimestre = dreTipo3

                    '''Fim da procura e calculo necessario ----------------------------------------------------------------------------------'''
                    for item in range(0, len(meses) , 3):
                        try:
                            colunaTrimestreJan = 'MvtoLíquido {}/{}'.format(meses[item], ano)
                            colunaTrimestreFev = 'MvtoLíquido {}/{}'.format(meses[item + 1], ano)
                            colunaTrimestreMar = 'MvtoLíquido {}/{}'.format(meses[item + 2], ano)
                            
                            #IR
                            try:
                                somaMesesLucroContabil = calculoTrimestre['resultadoDRE'][colunaTrimestreJan] + calculoTrimestre['resultadoDRE'][colunaTrimestreFev] + calculoTrimestre['resultadoDRE'][colunaTrimestreMar]
                            except:
                                somaMesesLucroContabil = calculoTrimestre['contas_resultado'][colunaTrimestreJan] + calculoTrimestre['contas_resultado'][colunaTrimestreFev] + calculoTrimestre['contas_resultado'][colunaTrimestreMar]

                            res_IR = round(somaMesesLucroContabil * 0.15, 2)

                            df.at[dreTrimestre['imposto']['Index_linha {}/{}'.format(meses[item + 2], ano)], 'MvtoLíquido {}/{}'.format(meses[item + 2], ano)] = res_IR

                            #ADD
                            if somaMesesLucroContabil < 0:
                                calc_add = round(somaMesesLucroContabil * 0.1, 2)
                            else:
                                if somaMesesLucroContabil == 0:
                                    calc_add = 0
                                else:
                                    calc_add = round((somaMesesLucroContabil - 60000) * 0.1, 2)
                            df.at[dreTrimestre['add']['Index_linha {}/{}'.format(meses[item + 2], ano)], 'MvtoLíquido {}/{}'.format(meses[item + 2], ano)] = calc_add
                            #CSLL
                            calc_csll = round(somaMesesLucroContabil * 0.09, 2)
                            df.at[dreTrimestre['csll']['Index_linha {}/{}'.format(meses[item + 2], ano)], 'MvtoLíquido {}/{}'.format(meses[item + 2], ano)] = calc_csll
                            #total
                            total_ = round(res_IR + calc_add + calc_csll,2)
                            df.at[dreTrimestre['total']['Index_linha {}/{}'.format(meses[item + 2], ano)], 'MvtoLíquido {}/{}'.format(meses[item + 2], ano)] = total_
                            #total_ir_add
                            tot_ir_add = round(res_IR + calc_add,2)
                            df.at[dreTrimestre['totalIrAdd']['Index_linha {}/{}'.format(meses[item + 2], ano)], 'MvtoLíquido {}/{}'.format(meses[item + 2], ano)] = tot_ir_add
                            '''--Fim do calculo trimestre--'''
                        except:
                            break

                    #substitui . por ,
                    def clean_value(x):
                        if isinstance(x, str):
                            x = x.replace('.', ',')
                            return x
                        return x
                    df.iloc[:, 4:] = df.iloc[:, 4:].map(clean_value)

                    try:
                        trimestre = trimestre.replace('\\', '')
                        trimestre = trimestre.split('T')
                    except:
                        pass

                    #inserir novas colunas com nome e unidade
                    df.insert(0,'Empresa',nomeEmpresaString)
                    df.insert(1,'Unidade', nomeUnidade)
                    df.insert(2,'Trimestre', trimestre[0] + 'º')
                    novaTabela.append(df)
                    
            #concatenando as tabelas
            novaTabela = pd.concat(novaTabela, ignore_index=True)
            
            # Salvar planilha Excel tratada, trocando index = True mostra o index das linhas
            try:
                if str(anoAtual) in nomeDaPasta:
                    novaTabela.to_excel(f'{ondeSalvar}{nomeTabelaTratada}.xlsx', index=False)
                else: 
                    novaTabela.to_excel(f'{ondeSalvar}{nomeTabelaTratada}-{ano}.xlsx', index=False)

            except PermissionError:
                print('Feche a planilha para salvar o arquivo')
                sys.exit()
            break

        finalContador = time.time()
        tempoExecucao = finalContador - inicioContador
        return f'Tempo de execução: <strong>{round(tempoExecucao,2)} seconds</strong>'
    except Exception as e:
        return f'Erro: {e}'

if __name__ == '__main__':
    ondeSalvarForge = r'C:\Users\User\Documents\GitHub\data-analytics-back-end\src\ArquivosDRE\\'
    #caminhoForge = r'\\192.168.1.2\dados\SUPERMERCADO CONTABIL\Planilhas resultados - Power Bi'
    caminhoForge = r'C:\Users\User\Documents\GitHub\data-analytics-back-end\src\ArquivosDRE\test\Planilhas resultados - Power Bi'
    
    tempoExecucao = forgePlan(caminhoForge, ondeSalvarForge, 2024, 2)
    print(tempoExecucao)
