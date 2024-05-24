from sqlalchemy import Column, String, Integer, Date
from DRE.conexão import *


class Planodecontas(Base):
    __tablename__ = 'planocon'

    pcon_classificacao = Column(Integer, primary_key=True)
    pcon_classref = Column(Integer, nullable=False)
    pcon_descricao = Column(String, nullable=False)
    pcon_conta = Column(Integer, nullable=False)
    pcon_tipo = Column(String, nullable=False)

    def __repr__(self):
        return f'{self.pcon_classref} {self.pcon_descricao} {self.pcon_conta} {self.pcon_tipo}'


class PlanoGeral(Base):
    __tablename__ = 'planoger'

    pger_classificacao = Column(Integer, primary_key=True)
    pger_descricao = Column(String,nullable=False)
    pger_conta = Column(Integer,nullable=False)
    pger_tipo = Column(String,nullable=False)

    def __repr__(self):
        return f'{self.pger_classificacao} {self.pger_descricao} {self.pger_conta} {self.pger_tipo}'
    

class Saldoscon(Base):
    __tablename__ = 'saldoscon'

    scon_pcon_conta = Column(Integer,primary_key=True)
    scon_unid_codigo = Column(String,nullable=False)
    scon_data = Column(Date, nullable=False)
    scon_valor = Column(Integer,nullable=False)
    scon_debcre = Column(String,nullable=False)
    
    def __repr__(self):
        return f'{self.scon_pcon_conta} {self.scon_unid_codigo} {self.scon_date} {self.scon_valor} {self.scon_debcre}'


class movimentacao_conta(Base):
    __tablename__ = 'movcon'

    mcon_transacao = Column(String,primary_key=True)
    mcon_operacao = Column(String,nullable=False)
    mcon_status = Column(String,nullable=False)
    mcon_datalcto = Column(String,nullable=False)
    mcon_datamvto = Column(String,nullable=False)
    mcon_pcon_conta = Column(Integer,nullable=False)
    mcon_zeramento = Column(Integer,nullable=False)
    mcon_unid_codigo = Column(String,nullable=False)
    mcon_valor = Column(Integer,nullable=False)
    mcon_dc = Column(String,nullable=False)
    mcon_hist_codigo = Column(String,nullable=False)
    mcon_complemento = Column(String,nullable=False)
    mcon_numerodcto = Column(String,nullable=False)
    mcon_audit = Column(String,nullable=False)

    def __repr__(self):
        return f'{self.mcon_transacao} {self.mcon_operacao} {self.mcon_status} {self.mcon_datalcto} {self.mcon_datamvto} {self.mcon_pcon_conta} {self.mcon_zeramento} {self.mcon_unid_codigo} {self.mcon_valor} {self.mcon_dc} {self.mcon_hist_codigo} {self.mcon_complemento} {self.mcon_numerodcto} {self.mcon_audit}'


