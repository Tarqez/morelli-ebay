# -*- coding: utf-8 -*-
import sys, csv, os, xlrd, zipfile, datetime

from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy import or_
from sqlalchemy import Column, Integer, Float, Unicode, Boolean, DateTime, PickleType


# DB def with Sqlalchemy
# ----------------------

db_file = os.path.join('db', 'db.sqlite')
engine = create_engine('sqlite:///'+db_file, echo=False)
Session = sessionmaker(bind=engine)
Base = declarative_base()

class Art(Base):
    __tablename__ = 'articles'

    id = Column(Integer, primary_key=True)
    mo_code = Column(Unicode, unique=True, index=True, nullable=False)
    descrizione = Column(Unicode, default=u'')
    categoria = Column(Unicode, default=u'')
    iva = Column(Integer, default=0)
    
    qty = Column(Integer, default=0)
    prc = Column(Float, default=0.0)

Base.metadata.create_all(engine)


# A csv.DictWriter specialized with Fx csv
# ----------------------------------------

class EbayFx(csv.DictWriter):
    '''Subclass csv.DictWriter, define delimiter and quotechar and write headers'''
    def __init__(self, filename, fieldnames):
        self.fobj = open(filename, 'wb')
        csv.DictWriter.__init__(self, self.fobj, fieldnames, delimiter=';', quotechar='"')
        self.writeheader()
    def close(self,):
        self.fobj.close()
       
    def __enter__(self):
        return self
    def __exit__(self, type, value, traceback):
        self.close()


# Constants
# ---------

DATA_PATH = os.path.join('data')
ACTION = '*Action(SiteID=Italy|Country=IT|Currency=EUR|Version=745|CC=UTF-8)' # smartheaders CONST

# Fruitful functions
# ------------------

def get_fname_in(folder):
    'Return the filename inside folder'

    t = os.listdir(folder)
    if len(t) == 0: 
        raise Exception('No file found')
    elif len(t) == 1:
        el = os.path.join(folder, t[0])
        if os.path.isfile(el):
            return el
        else:
            raise Exception('No file found')
    else:
        raise Exception('More files or folders found')


# loaders
# -------

def generic_data_loader(session):
    "Load data file into DB"
    folder = os.path.join(DATA_PATH)

    for fname in get_fname_in(folder):
        with open(fname, 'rb') as f:
            dsource_rows = csv.reader(f, delimiter=';', quotechar='"')
            dsource_rows.next()
            dstore_row = dict()
            for dsource_row in dsource_rows:
                try:
                    dstore_row['ga_code']=dsource_row[0]
                    dstore_row['brand']=' '.join(dsource_row[4].split()[1:]).title()
                    dstore_row['mnf_code']=dsource_row[6]
                    dstore_row['description']=dsource_row[2].decode('iso.8859-1')
                    dstore_row['category']=' '.join(dsource_row[5].split()[1:])
                    dstore_row['units_of_sale']=dsource_row[3].split().pop(0)
                    dstore_row['min_of_sale']=int(float((dsource_row[7].strip() or '0,0').replace('.', '').replace(',', '.')))
                    
                    art = session.query(Art).filter(Art.ga_code == dstore_row['ga_code']).first()

                    if not art: art = Art()
                    for attr, value in dstore_row.items():
                        setattr(art, attr, value)
                    session.add(art)
                except ValueError:
                    print 'rejected line:'
                    print dsource_row
                    print sys.exc_info()[0]
                    print sys.exc_info()[1]
                    print sys.exc_info()[2]
        os.remove(fname)
        session.commit()

def add(session):
    'Fx add action'
    smartheaders = (ACTION,
                    '*Category=50584',
                    '*Title',
                    'Description',
                    PICURL,
                    '*Quantity',
                    '*StartPrice',
                    'StoreCategory=1',
                    'CustomLabel',
                    '*ConditionID=1000',
                    '*Format=StoresFixedPrice',
                    '*Duration=GTC',
                    'OutOfStockControl=true',
                    '*Location=Matera',
                    'VATPercent=22',
                    '*ReturnsAcceptedOption=ReturnsAccepted',
                    'ReturnsWithinOption=Days_30',
                    'ShippingCostPaidByOption=Buyer',                    
                    # Regole di vendita
                    'PaymentProfileName=PayPal-Bonifico',
                    'ReturnProfileName=Reso1',
                    'ShippingProfileName=GLS_paid',
                    # specifiche oggetto
                    'C:Marca',
                    'C:Modello',
                    'C:Genere',
                    'Counter=BasicStyle',)
    
    arts = session.query(Art).filter(Art.ebay_itemid == u'',
                                     Art.units_of_sale == 'PZ',
                                     Art.min_of_sale <= 1, #0 ia as 1
                                     Art.ebay_price > 1,
                                     Art.sale_control >= 0,
                                     Art.ebay_qty > 0)

    fout_name = os.path.join(DATA_PATH, 'fx_output', fx_fname('add', session))
    gacodes_of_images = items_with_img()
    with EbayFx(fout_name, smartheaders) as wrt:
        for art in arts:
            title = ebay_title(art.brand, art.description, art.mnf_code)
            context = {'ga_code':art.ga_code,
                       'title':title,
                       'description':'',
                       'email':EMAIL,
                       'phone':PHONE,
                       'invoice_form_url':INVOICE_FORM_URL,}
            ebay_description = ebay_template('garofoli', context)
            fx_add_row = {ACTION:'Add',
                          '*Title':title.encode('iso-8859-1'),
                          'Description':ebay_description,
                          '*Quantity':art.ebay_qty,
                          '*StartPrice':art.ebay_price,
                          'ShippingProfileName=GLS_paid':'',
                          'CustomLabel':art.ga_code,
                          PICURL:'http://'+FTPURL+'/'+art.ga_code+'.jpg' if art.ga_code in gacodes_of_images else '',
                          'StoreCategory=1':store_cat_n(art.category, session),
                          '*Category=50584':ebay_cat_n(art.category, session),
                          'C:Marca':art.brand,
                          'C:Modello':art.mnf_code,
                          'C:Genere':art.category}
            wrt.writerow(fx_add_row)        