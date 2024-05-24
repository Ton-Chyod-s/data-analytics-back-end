from DRE.analise import *
from DRE.tabelas import *

"""plano_geral(PlanoGeral,
        'produtos',
        'Classificação',
        'Classificação','Descrição Conta','Conta','Tipo'
        )

saldos_con(
        Saldoscon,
        'saldos',
        'scon_pcon_conta','scon_unid_codigo','scon_date','scon_valor','scon_debcre'
    )

movcon(movimentacao_conta,
        'arquivo-movcon',
        'mcon_transacao',
        'mcon_operacao',
        'mcon_status',
        'mcon_datalcto', 
        'mcon_datamvto', 
        'mcon_pcon_conta', 
        'mcon_zeramento', 
        'mcon_unid_codigo', 
        'mcon_valor', 
        'mcon_dc', 
        'mcon_hist_codigo', 
        'mcon_complemento', 
        'mcon_numerodcto', 
        'mcon_audit'
        )"""

produtos = session.query(movimentacao_conta).all()
print(produtos)