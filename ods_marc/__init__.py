"""Converts an Excel export of ODS data to MARC for import into Horizon"""

import sys, os, re, logging
from datetime import datetime
from argparse import ArgumentParser
from dlx import DB
from dlx.util import Table
from dlx.marc import Bib, BibSet, Auth, Query, Condition, Or
#from bson import Regex
from pymongo.collation import Collation

logging.basicConfig(level='INFO')

###

def args(): 
    ap = ArgumentParser(
        description='Converts an Excel export of ODS data to MARC for import into Horizon',
        usage='\n$ ods-marc --connect=mongodb://dummy --input_file=in.xlsx --output_format=mrc --output_file=out.mrc'
    )
    ap.add_argument('--connect', required=True, help='MDB connection string')
    ap.add_argument('--input_file', required=True, help='Path to ODS Excel file')
    ap.add_argument('--output_format', required=True, choices =['mrc', 'mrk'])
    ap.add_argument('--output_file', help='File path to write output to. Default: use input filename + new extension')

    return ap.parse_args()

###

class Tcode():
    cache = {}
    
    @classmethod
    def lookup(cls, tcode):
        if tcode in cls.cache:
            return cls.cache[tcode] 
            
        q = Query(Condition('035', {'a': tcode}))
        auth = Auth.find_one(q.compile(), {'_id': 1})
        
        if auth:
            cls.cache[tcode] = auth.id
            return auth.id
        else:
            cls.cache[tcode] = None
            return

### 

def _symbol(bib, value):
    bib.set('191', 'a', value)
    
    # double symbols?
     
    return bib
    
def _title(bib, value):
    title_val = value.title() if value else '[Missing title]'      
    bib.set('245', 'a', title_val)
    bib.get_field('245').ind1 = '1'
    
    return bib
    
def _date(bib, value):
    bib.set('269', 'a', value)
    
    dt = datetime.strptime(value, '%Y%m%d')
    value = dt.strftime('%d %b. %Y')
    
    if value[0] == '0':
        value = value[1:]
    
    for old, new in {'May.': 'May', 'Jun.': 'June', 'Jul.': 'July', 'Sep': 'Sept'}.items():
        value = value.replace(old, new)
        
    bib.set('260', 'c', value)
    
    return bib
    
def _langs(bib, value):
    langtext = ''
    langs = set([x[0:1] for x in value.split(' ')])
    
    for lang in sorted(langs):        
        langtext += {'A': 'ara', 'C': 'chi', 'E': 'eng', 'F': 'fre', 'R': 'rus', 'S': 'spa', 'O': 'ger'}.get(lang, '')
        
    bib.set('041', 'a', langtext)
    
    return bib
    
def _job(bib, value):
    lang = {'ara': 'A', 'chi': 'C', 'eng': 'E', 'fre': 'F', 'rus': 'R', 'spa': 'S', 'ger': 'O'}
    langs = []
    
    for l in sorted(lang.keys()):
        if l in bib.get_value('041', 'a'):
            langs.append(lang[l])
    
    place = 0
    
    for job in value.split(';'):
        langcode = ' ' + langs.pop(0) if len(langs) > 0 else ''
        bib.set('029', 'a', 'JN', address=['+'])
        bib.set('029', 'b', job + langcode, address=[place])
        
        place += 1
        
    return bib
    
def _tcodes(bib, value):
    for tcode in re.split(r'[,;]', value):
        if not tcode:
            continue
        
        auth_id = Tcode.lookup(tcode)
        
        if auth_id:
            bib.set('650', 'a', auth_id, address=['+'])
        else:
            logging.warning('Auth record for Tcode "{}" not found'.format(tcode))
            
        for field in bib.get_fields('650'):
            field.ind1 = '1'
            field.ind2 = '7'
    
    return bib

###

def run(args=args()):
    DB.connect(args.connect)
    output_path = args.output_file or os.path.expanduser(args.input_file.replace('xlsx', args.output_format))
    output_handle = open(output_path, 'w')
    table = Table.from_excel(args.input_file, date_format='%Y%m%d')
    bibs = BibSet()
    
    dispatch = {
        'doc symbol': _symbol,
        'title': _title,
        'publication date': _date,
        'lang available': _langs,
        'job number': _job,
        'subjects': _tcodes,
    }

    for row in table.index.keys():
        bib = Bib()
        exists = False
        seen = []

        for field_name in reversed(sorted(table.index[row].keys())):
            todo = dispatch.get(field_name.lower())
            
            if not todo:
                #raise Exception(f'Field "{field_name}" not found. Recognized fields are {list(dispatch.keys())}')
                continue

            if todo:
                value = table.index[row][field_name].strip()
                bib = todo(bib, value)
                seen.append(field_name)
                
            if field_name == 'Doc Symbol':
                symbols = bib.get_values('191', 'a')
                symbols = list(set(symbols))
                
                q = Query(
                    Or(
                        Condition('191', {'a': {'$in': symbols}}),
                        Condition('191', {'z': {'$in': symbols}}),
                    )
                )
                
                existing = next(DB.bibs.find(q.compile(), {'_id': 1}).collation(Collation(locale='en', strength=2)), None)
                
                if existing:
                    logging.warning('{} is already in the system as id {}'.format(symbols, existing['_id']))
                    exists = True
                    break
                    
        for req in list(dispatch.keys()):
            if req not in map(lambda x: x.lower(), seen):
                raise Exception(f'Field "{req.title()}" not found in row {row}')
        
        if exists:
            continue
        
        leader = ['|' for x in range(0,24)]
        leader[5] = 'n'
        leader[6] = 'a'
        leader[7] = 'm'
        leader[17] = '#'
        leader[18] = 'a'
        bib.set('000', None, ''.join(leader))
        bib.set_008()
        bibs.records.append(bib)
    
    if args.output_format == 'mrk':
        output_handle.write(bibs.to_mrk())
    else:
        output_handle.write(bibs.to_mrc())

    logging.info('Done: ' + output_path)

###

if __name__ == '__main__':
    run(args())
