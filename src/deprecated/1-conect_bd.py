from sqlalchemy import create_engine, Column, String, Integer
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker
import pandas as pd

engine = create_engine('postgresql://postgres:gm560max2005@localhost:5432/erp?client_encoding=LATIN-1')
Base = declarative_base()
Session = sessionmaker(bind=engine)
session = Session()

class Produto(Base):
    __tablename__ = 'produtos'

    prod_codigo = Column(Integer, primary_key=True)
    prod_descricao = Column(String, nullable=False)
    prod_codbarras = Column(Integer, nullable=False)
    prod_datacad = Column(Integer, nullable=False)
    prod_descrpdvs = Column(String, nullable=False)

    def __repr__(self):
        return f'Chave Primaria: {self.prod_codigo} Código de Barras: {self.prod_codbarras} Data cadastrada: {self.prod_datacad}\nDescrição: {self.prod_descricao}\nDescrição produto: {self.prod_descrpdvs}'

class Planodecontas(Base):
    __tablename__ = 'planocon'

    pcon_classificacao = Column(Integer, primary_key=True)
    pcon_classref = Column(Integer, nullable=False)
    pcon_descricao = Column(String, nullable=False)
    pcon_conta = Column(Integer, nullable=False)
    pcon_tipo = Column(String, nullable=False)

    def __repr__(self):
        return f'{self.pcon_classref} {self.pcon_descricao} {self.pcon_conta} {self.pcon_tipo}'
    
# Consultar produtos na tabela
produtos = session.query(Planodecontas).all()

"""# Imprimir os produtos encontrados
for produto in produtos:
    print(produto)"""

# Convert the SQLAlchemy result to a Pandas DataFrame
df = pd.DataFrame([(p.pcon_classref, p.pcon_descricao, p.pcon_conta, p.pcon_tipo) for p in produtos],
                  columns=['Classificação', 'Descrição Conta', 'Conta', 'Tipo'])
#ordenar data
df_reordenado = df.sort_values(by='Classificação')

# Save the DataFrame to a CSV file
df_reordenado.to_excel('produtos.xlsx', index=False)