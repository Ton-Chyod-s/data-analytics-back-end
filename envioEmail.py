import time
from time import sleep
from src.ArquivoResCliente.limpando_planilha_refactore import limp_plan
from src.ArquivosDRE.planForgeRefactore_copy import forgePlan
from src.enviarEmail.email import enviar_email, assunto, de, para, senha

inicioContador = time.time()
caminhoPastaLimp = r'\\192.168.1.2\dados\SUPERMERCADO ACOMPANHAMENTO\RESULTADOS CLIENTES - ALEX 2024.xlsx'
ondeSalvarLimp = r'C:\Users\User\Documents\GitHub\data-analytics-back-end\src\ArquivosDRE\\'

ondeSalvarForge = r'C:\Users\User\Documents\GitHub\data-analytics-back-end\src\ArquivosDRE\\'
caminhoForge = r'\\192.168.1.2\dados\SUPERMERCADO CONTABIL\Planilhas resultados - Power Bi'

print('Iniciando as analises...')
tempoLimp = limp_plan(caminhoPastaLimp, ondeSalvarLimp)
tempoForge = forgePlan(caminhoForge, ondeSalvarForge, 2024)

corpo_mensagem = f'''<p>Olá, segue em anexo o resultado do processo de limpeza e da planilha DRE.</p>
<p>Planilha Resultado Cliente<br>{tempoLimp}</p>
<p>Planilha DRE<br>{tempoForge}</p>
<p>Atenciosamente,<br><i>PerinDevBoot~</i></p>'''

enviar_email(assunto, de, para, senha, corpo_mensagem)

finalContador = time.time()
tempoExecucao = finalContador - inicioContador
print(f'Tempo de execução total: {round(tempoExecucao,2)} s')
sleep(1.5)