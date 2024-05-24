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
dreAno = 2024

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

    caminhoPasta = r'\\192.168.1.2\dados\SUPERMERCADO CONTABIL\Planilhas resultados - Power Bi' + nomeDaPasta + trimestre 
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
        #pegando nome da empresa
        if nomeAba == None:
            nomeAbaTabela = pd.ExcelFile(os.path.join(caminhoPasta,i)).sheet_names
            nomeAba = nomeAbaTabela[0]
        
        #tratamento das abas da planilha localizada
        for k, v in enumerate(nomeAbaTabela):
            sleep(.1)
            df = pd.read_excel(os.path.join(caminhoPasta,i), sheet_name=v)
            if nomeEmpresaString == None:
                nomeEmpresaString = arquivosPasta[num].split('.')[0]
            nomeUnidade = f'Loja {v}'
            nome = nomeEmpresaString.split(' ')[0]

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
            
            elif 'VRA - 1ºT 2024' in nomeEmpresaString:    
                break

            try:
                primeiraCelula = df.keys()[0]
                terceiraCelula = df.keys()[2]
            except IndexError as e:
                print(f'Empresa {nomeEmpresaString}')
                break
            
            tanhoPrimeiraCelula = len(primeiraCelula)
            #se o dataframe não tiver a primeira linha em branco, vai acrescentar a primeira linha duplicada
            if tanhoPrimeiraCelula > 10:
                df.iloc[0] = df.columns
                df.columns = [f"Unnamed: {i}" for i in range(len(df.columns))]

            data_emissao = df.iloc[0,2]

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
            for key, value in enumerate(segunda_linha):
                if key == 10:
                    try:
                        anoDre = dreAno
                    except:
                        print(nomeEmpresaString)
                        print(nomeUnidade)
                    if anoDre != ano:
                        ano = anoDre
                        break
                    else:
                        break

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

            '''Inicio da procura e calculo necessario ----------------------------------------------------------------------------------'''
            #dicionario a ser preenchido
            #Planilha DRE
            contas_resultado = {}
            provisao_de_imposto = {}
            receitas_operacaional = {}
            custo_mercadorias = {}
            receitas_operacaional = {}
            venda_mercadorias = {}
            recuperacao_despesas = {}
            contribuicao_social = {}
            lucro = {}

            outras_receitas = {}
            outras_receitas_nao_operacionais = {}
            venda_brutal_mercadorias = {}
            multas_fiscais = {}

            #Trimestre
            venda = {}
            lucro_contabil = {}
            recolhido = {}
            receita_nao_operacional = {}
            imposto = {}
            add = {}
            csll = {}
            total = {}
            total_ir_add = {}
            margem_contabil = {}
            resultado_liquido = {}

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
                
                def preencher_dicionario(linha, dicionario, contem_linha=True):
                    try:
                        valor_linha, indice_linha, num_coluna = valor_da_linha(linha, coluna_a_encontrar, contem_linha)
                        dicionario[f'MvtoLíquido {mes}/{ano}'] = valor_linha
                        dicionario[f'Index_linha {mes}/{ano}'] = indice_linha
                        dicionario[f'Index_coluna {mes}/{ano}'] = num_coluna
                    except:
                        print('na empresa\t',nomeEmpresaString)

                if nome == 'Veratti':
                    #contas de resultado
                    preencher_dicionario('CONTAS DE RESULTADO',contas_resultado)
                    preencher_dicionario('Outras Receitas - Não Tributáveis',outras_receitas)
                    preencher_dicionario('PROVISAO DE IMPOSTO S/L', provisao_de_imposto)
                    preencher_dicionario('Multas Fiscais - Não Dedutíveis', multas_fiscais)

                    calc_contas_result = contas_resultado[f'MvtoLíquido {mes}/{ano}'] - outras_receitas[f'MvtoLíquido {mes}/{ano}'] - provisao_de_imposto[f'MvtoLíquido {mes}/{ano}'] - multas_fiscais[f'MvtoLíquido {mes}/{ano}']
                    
                    df.at[contas_resultado[f'Index_linha {mes}/{ano}'], f'MvtoLíquido {mes}/{ano}'] = calc_contas_result
                    att(contas_resultado,calc_contas_result)

                    #resultado liquido
                    preencher_dicionario('RESULTADO LIQUIDO',resultado_liquido)
                    preencher_dicionario('Outras Receitas Não Operacionais',outras_receitas_nao_operacionais)
                    calc_resultado_liquido = contas_resultado[f'MvtoLíquido {mes}/{ano}'] - outras_receitas_nao_operacionais[f'MvtoLíquido {mes}/{ano}']
                    df.at[resultado_liquido[f'Index_linha {mes}/{ano}'], f'MvtoLíquido {mes}/{ano}'] = calc_resultado_liquido
                    
                    #Trimestre
                    #venda
                    preencher_dicionario('VENDA',venda, False)
                    preencher_dicionario('VENDA BRUTA DE MERCADORIAS',venda_brutal_mercadorias)
                    df.at[venda[f'Index_linha {mes}/{ano}'], f'MvtoLíquido {mes}/{ano}'] = venda_brutal_mercadorias[f'MvtoLíquido {mes}/{ano}']

                    #lucro contabil
                    preencher_dicionario('LUCRO CONTABIL',lucro_contabil, False)
                    lucroContabil = contas_resultado[f'MvtoLíquido {mes}/{ano}']
                    df.at[lucro_contabil[f'Index_linha {mes}/{ano}'], f'MvtoLíquido {mes}/{ano}'] = lucroContabil

                    #calculo recolhido
                    preencher_dicionario('RECOLHIDO',recolhido, False)
                    if variavel != 0 :
                        preencher_dicionario('Contribuição Social',contribuicao_social)
                        comp = round(-contribuicao_social[f'MvtoLíquido {mes}/{ano}'] / 0.9  * 10, 2)
                    else:
                        preencher_dicionario('Contribuicao Social',contribuicao_social,False)
                        comp = contribuicao_social[f'MvtoLíquido {mes}/{ano}']
                    df.at[recolhido[f'Index_linha {mes}/{ano}'], f'MvtoLíquido {mes}/{ano}'] = comp

                    #calculo Receitas Não Operacionais 
                    preencher_dicionario('RECEITAS NAO OPERACIONAIS',receita_nao_operacional, False)
                    df.at[receita_nao_operacional[f'Index_linha {mes}/{ano}'], f'MvtoLíquido {mes}/{ano}'] = outras_receitas_nao_operacionais[f'MvtoLíquido {mes}/{ano}']
                    
                    #lucro
                    preencher_dicionario('LUCRO',lucro, False)
                    lucro_res = lucroContabil - receita_nao_operacional[f'MvtoLíquido {mes}/{ano}']
                    #adicionando na planilha
                    df.at[lucro[f'Index_linha {mes}/{ano}'], f'MvtoLíquido {mes}/{ano}'] = lucroContabil - outras_receitas_nao_operacionais[f'MvtoLíquido {mes}/{ano}']

                    #IR
                    preencher_dicionario('IR',imposto,False)
                    res_IR = round(lucroContabil * 0.15, 2)
                    #adicionando na planilha
                    df.at[imposto[f'Index_linha {mes}/{ano}'], f'MvtoLíquido {mes}/{ano}'] = res_IR

                    #ADD
                    preencher_dicionario('ADD',add)
                    if lucroContabil < 0:
                        calc_add = round(lucroContabil * 0.1, 2 )
                    else:
                        if lucroContabil == 0:
                            calc_add = 0
                        else:
                            calc_add = round((lucroContabil - 60000) * 0.1, 2)
                    #adicionando na planilha
                    df.at[add[f'Index_linha {mes}/{ano}'], f'MvtoLíquido {mes}/{ano}'] = calc_add

                    #CSLL
                    preencher_dicionario('CSLL',csll)
                    calc_csll = round(lucroContabil * 0.09,2 )
                    df.at[csll[f'Index_linha {mes}/{ano}'], f'MvtoLíquido {mes}/{ano}'] = calc_csll

                    #total
                    preencher_dicionario('TOTAL',total)
                    total_ = round(res_IR + calc_add + calc_csll,2)
                    df.at[total[f'Index_linha {mes}/{ano}'], f'MvtoLíquido {mes}/{ano}'] = total_

                    #total_ir_add
                    preencher_dicionario('TOTAL IR-ADD',total_ir_add)
                    tot_ir_add = round(res_IR + calc_add,2)
                    df.at[total_ir_add[f'Index_linha {mes}/{ano}'], f'MvtoLíquido {mes}/{ano}'] = tot_ir_add

                elif 'Pires' in nome:
                    break
                
                else:
                    #atualização conta de resultado (contas de resultado - provisão de imposto)      
                    preencher_dicionario('CONTAS DE RESULTADO',contas_resultado)
                    preencher_dicionario('PROVISAO DE IMPOSTO S/L', provisao_de_imposto)
                    try:
                        preencher_dicionario('RECEITAS OPERACIONAL LIQUIDA', receitas_operacaional)
                        variavel = 0
                        #contas de resultado calculo e atualização na planilha                 
                        resultado_analisado = contas_resultado[coluna_a_encontrar] + provisao_de_imposto[coluna_a_encontrar] * - 1
                        att(contas_resultado,resultado_analisado)

                        #calculo resultado liquido (contas de resultado - Recuperacao De Despesas Exerc Anterior)
                        preencher_dicionario('Recuperacao De Despesas Exerc Anterior',recuperacao_despesas)
                        preencher_dicionario('RESULTADO LIQUIDO',resultado_liquido)

                        calc_res_liquido = contas_resultado[f'MvtoLíquido {mes}/{ano}'] - recuperacao_despesas[f'MvtoLíquido {mes}/{ano}']
                        df.at[resultado_liquido[f'Index_linha {mes}/{ano}'], f'MvtoLíquido {mes}/{ano}'] = calc_res_liquido

                        #calculo margem contabil (receitas op liquida - CVM custo de mercadorias)
                        try:
                            preencher_dicionario('CUSTO DAS MERCADORIAS VENDIDOS - CMV', custo_mercadorias)
                        except:
                            preencher_dicionario('Custos Das Mercadorias e Serviços Vendidos', custo_mercadorias)

                        preencher_dicionario('MARGEM CONTABIL',margem_contabil)
                        calc_margem_cont = receitas_operacaional[f'MvtoLíquido {mes}/{ano}'] + custo_mercadorias[f'MvtoLíquido {mes}/{ano}']
                        df.at[margem_contabil[f'Index_linha {mes}/{ano}'], f'MvtoLíquido {mes}/{ano}'] = calc_margem_cont

                        '''--dict trimestre--'''
                        #calculo venda /trimestre
                        preencher_dicionario('VENDA DE MERCADORIAS',venda_mercadorias)
                        preencher_dicionario('VENDA',venda, False)
                        #adicionando na planilha
                        df.at[venda[f'Index_linha {mes}/{ano}'], f'MvtoLíquido {mes}/{ano}'] = venda_mercadorias[f'MvtoLíquido {mes}/{ano}']
                    except:
                        pass

                    #lucro contabil == soma de conta de resultados
                    preencher_dicionario('LUCRO CONTABIL',lucro_contabil, False)
                    lucroContabil = contas_resultado[f'MvtoLíquido {mes}/{ano}']
                    df.at[lucro_contabil[f'Index_linha {mes}/{ano}'], f'MvtoLíquido {mes}/{ano}'] = lucroContabil

                    #calculo recolhido /trimestre
                    preencher_dicionario('RECOLHIDO',recolhido, False)
                    if variavel != 0 :
                        preencher_dicionario('Contribuição Social',contribuicao_social)
                        comp = contribuicao_social[f'MvtoLíquido {mes}/{ano}'] * 0.9
                    else:
                        try:
                            preencher_dicionario('Contribuicao Social',contribuicao_social,False)
                        except:
                            preencher_dicionario('Contribuicao Sindical',contribuicao_social,False)

                        comp = contribuicao_social[f'MvtoLíquido {mes}/{ano}']
                    df.at[recolhido[f'Index_linha {mes}/{ano}'], f'MvtoLíquido {mes}/{ano}'] = comp
                    try:
                        #calculo Receitas Não Operacionais 
                        preencher_dicionario('RECEITAS NAO OPERACIONAIS',receita_nao_operacional, False)
                        df.at[receita_nao_operacional[f'Index_linha {mes}/{ano}'], f'MvtoLíquido {mes}/{ano}'] = recuperacao_despesas[f'MvtoLíquido {mes}/{ano}']

                        #lucro
                        preencher_dicionario('LUCRO',lucro, False)
                        lucro_res = contas_resultado[coluna_a_encontrar] - recuperacao_despesas[coluna_a_encontrar]
                        lucro[f'MvtoLíquido {mes}/{ano}'] = lucro_res
                        #adicionando na planilha
                        df.at[lucro[f'Index_linha {mes}/{ano}'], f'MvtoLíquido {mes}/{ano}'] = lucro[f'MvtoLíquido {mes}/{ano}']
                
                        #IR
                        preencher_dicionario('IR',imposto,False)
                        res_IR = round(lucroContabil * 0.15, 2)
                        #adicionando na planilha
                        df.at[imposto[f'Index_linha {mes}/{ano}'], f'MvtoLíquido {mes}/{ano}'] = res_IR

                        #ADD
                        preencher_dicionario('ADD',add)
                        if lucroContabil < 0:
                            calc_add = round(lucroContabil * 0.1, 2 )
                        else:
                            if lucroContabil == 0:
                                calc_add = 0
                            else:
                                calc_add = round((lucroContabil - 60000) * 0.1, 2)
                        #adicionando na planilha
                        df.at[add[f'Index_linha {mes}/{ano}'], f'MvtoLíquido {mes}/{ano}'] = calc_add

                        #CSLL
                        preencher_dicionario('CSLL',csll)
                        calc_csll = round(lucroContabil * 0.09,2 )
                        df.at[csll[f'Index_linha {mes}/{ano}'], f'MvtoLíquido {mes}/{ano}'] = calc_csll

                        #total
                        preencher_dicionario('TOTAL',total)
                        total_ = round(res_IR + calc_add + calc_csll,2)
                        df.at[total[f'Index_linha {mes}/{ano}'], f'MvtoLíquido {mes}/{ano}'] = total_

                        #total_ir_add
                        preencher_dicionario('TOTAL IR-ADD',total_ir_add)
                        tot_ir_add = round(res_IR + calc_add,2)
                        df.at[total_ir_add[f'Index_linha {mes}/{ano}'], f'MvtoLíquido {mes}/{ano}'] = tot_ir_add
                    except:
                        pass
                

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
    if nomeDaPasta == anoAnterior:
        novaTabela.to_excel(os.path.abspath(f'{nomeTabelaTratada}-2023.xlsx'), index=False)
    else:
        novaTabela.to_excel(os.path.abspath(f'{nomeTabelaTratada}.xlsx'), index=False)

    # Exemplo de uso:
    corpo_mensagem = f"""<p><strong>Atualizado com sucesso {nomeTabelaTratada}.xlsx</strong><br>Arquivos - </p>
    <p> {arquivos} </p>"""

    assunto = "Atualização DRE-servidor"
    para = "klayton.oliveira@perincontabil.com.br"
    de = "perindevboot@gmail.com"
    senha = "gxkqsyymnogquthd"

    enviar_email(assunto, de, para, senha, corpo_mensagem)

    finalContador = time.time()
    tempoExecucao = finalContador - inicioContador
    print(f"Tempo total de execução: {tempoExecucao:.2f} segundos")
    time.sleep(1.5)