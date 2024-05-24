from sqlalchemy import create_engine
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker

engine = create_engine('postgresql://postgres:gm560max2005@localhost:5432/erp?client_encoding=LATIN-1')
Base = declarative_base()
Session = sessionmaker(bind=engine)
session = Session()