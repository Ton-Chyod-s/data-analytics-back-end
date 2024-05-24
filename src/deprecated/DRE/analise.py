from DRE.tabelas import *
import pandas as pd


def plano_geral(tabela,
          arquivo,
          classificacao,
          coluna1,coluna2,coluna3,coluna4,
          salvar=True
          ):
    # Consultar produtos na tabela
    produtos = session.query(tabela).all()
    # Convert the SQLAlchemy result to a Pandas DataFrame
    df = pd.DataFrame([(p.pger_classificacao, p.pger_descricao, p.pger_conta, p.pger_tipo) for p in produtos],
                    columns=[coluna1,coluna2,coluna3,coluna4])
    
    #ordenar data
    df_reordenado = df.sort_values(by=f'{classificacao}')
    if salvar:
        # Save the DataFrame to a CSV file
        df_reordenado.to_excel(f'{arquivo}.xlsx', index=False)
    else:
        return df_reordenado

def saldos_con(tabela,
          arquivo,
          coluna1,coluna2,coluna3,coluna4,coluna5,
          salvar=True
          ):
    # Consultar produtos na tabela
    produtos = session.query(tabela).all()
    # Convert the SQLAlchemy result to a Pandas DataFrame
    df = pd.DataFrame([(p.scon_pcon_conta, p.scon_unid_codigo, p.scon_data, p.scon_valor,p.scon_debcre) for p in produtos],columns=[coluna1,coluna2,coluna3,coluna4,coluna5])
    if salvar:
        # Save the DataFrame to a CSV file
        df.to_excel(f'{arquivo}.xlsx', index=False)
    else:
        return df

def movcon(tabela,
          arquivo,coluna0,
          coluna1,coluna2,coluna3,coluna4,coluna5,coluna6,coluna7,coluna8,coluna9,coluna10,coluna11,coluna12,coluna13,
          salvar=True
          ):
    # Consultar produtos na tabela
    produtos = session.query(tabela).all()
    # Convert the SQLAlchemy result to a Pandas DataFrame
    df = pd.DataFrame([(p.mcon_transacao, 
                        p.mcon_operacao, 
                        p.mcon_status, 
                        p.mcon_datalcto,
                        p.mcon_datamvto,
                        p.mcon_pcon_conta,
                        p.mcon_zeramento,
                        p.mcon_unid_codigo,
                        p.mcon_valor,
                        p.mcon_dc,
                        p.mcon_hist_codigo,
                        p.mcon_complemento,
                        p.mcon_numerodcto,
                        p.mcon_audit) for p in produtos],columns=[coluna0,coluna1,coluna2,coluna3,coluna4,coluna5,coluna6,coluna7,coluna8,coluna9,coluna10,coluna11,coluna12,coluna13])

    if salvar:
        # Save the DataFrame to a CSV file
        df.to_excel(f'{arquivo}.xlsx', index=False)
    else:
        return df
        
    
    
    