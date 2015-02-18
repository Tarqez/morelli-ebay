# -*- coding: utf-8 -*-
import sys, csv, os, xlrd, zipfile, datetime, jinja2

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
FTPURL = 'morelli.ftpimg.eu'
PICURL = 'PicURL=http://'+FTPURL+'/nopic.jpg' # smartheaders CONST
EMAIL = 'parafarmaciamorelli@gmail.com'
PHONE = '0835/680388'

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

def ebay_template(tpl_name, context):
    'Return one line html ebay templated for description field'
    try:
        template_loader = jinja2.FileSystemLoader(
            searchpath=os.path.join(os.getcwd(), 'templates', tpl_name))
        template_env = jinja2.Environment(loader=template_loader)

        template = template_env.get_template('index.htm')
        output = template.render(context)
        res = ' '.join(output.split()).encode('iso-8859-1')
        return res
    except:
        print sys.exc_info()[0]
        print sys.exc_info()[1]        


# datasources
# -----------

def datasource(fcsv):
    'Yield a dict of values from estrazione.xls'

    data_line = dict()

    with open(fcsv, 'rb') as f:
        dsource_rows = csv.reader(f, delimiter=';', quotechar='"')
        dsource_rows.next()
        for row in dsource_rows:
            try:
                data_line['mo_code'] = row[0].strip()
                data_line['descrizione'] = row[1].strip()
                data_line['prc'] = float(row[2].strip())/1000
                data_line['categoria'] = row[3].strip()
                data_line['qty'] = int(row[4].strip())                            
                data_line['iva'] = row[5].strip().lower()
                
                yield data_line
                
            except ValueError:
                print 'rejected line:'
                print row
                print sys.exc_info()[0]
                print sys.exc_info()[1]
                print sys.exc_info()[2]        


# loaders
# -------

def data_loader():
    "Load data file into DB"
    fname = os.path.join(DATA_PATH, 'estrazione.csv')

    for data_line in datasource(fname):
        try:
            art = s.query(Art).filter(Art.mo_code == data_line['mo_code']).first()
            if not art: # not exsists, create
                art = Art()

            art.mo_code = data_line['mo_code']
            art.descrizione = data_line['descrizione']
            art.prc = data_line['prc']
            art.categoria = data_line['categoria']
            art.qty = data_line['qty']
            art.iva = data_line['iva']

            s.add(art)
        except ValueError:
            print 'rejected line:'
            print data_line
            print sys.exc_info()[0]
            print sys.exc_info()[1]
            print sys.exc_info()[2]
    s.commit()

def add():
    'Fx add action'
    smartheaders = (ACTION,
                    '*Category=67589',
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
                    # Regole di vendita
                    'PaymentProfileName=PayPal',
                    'ReturnProfileName=Reso30gg',
                    'ShippingProfileName=Raccomandata-Paccocelere',
                    'Counter=BasicStyle',)

store_cat_n = {
    'Accessori agli art. sanitari':'9115802014'
    'Accessori ai pmc':'9115803014'
    'Acque minerali':'9115797014'
    'Alimenti per la prima infanzia':'9115812014'
    'Ausili sanitari':'9115801014'
    'Diagnostici in vitro':'9115794014'
    'Edulcoranti sintetici':'9115818014'
    'Erboristeria salut. preconf.':'9115791014'
    'Integratori alimentari':'9115789014'
    'Materiale per protesi':'9115805014'
    'Parafarmaci':'9115788014'
    'Presidi medico-chirurgici':'9115800014'
    'Prod.per igiene int.uso inter.':'9115793014'
    'Prodotti capelli e cuoio cap.':'9115799014'
    'Prodotti dietetici':'9115804014'
    'Prodotti igiene del bambino':'9115815014'
    'Prodotti igiene del corpo':'9115807014'
    'Prodotti igiene dentale':'9115806014'
    'Prodotti omeopatici':'9115811014'
    'Prodotti per il corpo':'9115795014'
    'Prodotti per le mani':'9115808014'
    'Prodotti per uomo':'9115816014'
    'Prodotti sanitari':'9115790014'
    'Prodotti solari':'9115809014'
    'Prodotti viso, trattamento':'9115792014'
    'Prodotti viso, trucco':'9115817014'
    'Prodotti viso/deterg./struc.':'9115813014'
    'Prodotti zootecnici':'9115814014'
    'Sostanze materie prime uso lab':'9115796014'
    'Sostanze preconf. per vendita':'9115798014'
    'Strumenti sanitari':'9115810014'
}
    
    arts = s.query(Art).all()

    fout_name = os.path.join(DATA_PATH, 'add.csv')
    with EbayFx(fout_name, smartheaders) as wrt:
        for art in arts:
            context = {'mo_code':art.mo_code,
                       'title':art.descrizione,
                       'description':'',
                       'email':EMAIL,
                       'phone':PHONE,
                       'invoice_form_url':INVOICE_FORM_URL,}
            ebay_description = ebay_template('garofoli', context)
            fx_add_row = {ACTION:'VerifyAdd',
                          '*Title':art.descrizione.encode('iso-8859-1'),
                          'Description':ebay_description,
                          '*Quantity':art.qty,
                          '*StartPrice':art.prc,
                          'CustomLabel':art.mo_code,
                          'StoreCategory=1':store_cat_n(art.categoria),
                          }
            wrt.writerow(fx_add_row)        