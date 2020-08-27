"""Converts an Excel export of ODS data to MARC for import into Horizon"""

import sys, os, logging, re
from datetime import datetime
from argparse import ArgumentParser
from dlx import DB
from dlx.util import Table
from dlx.marc import Bib, BibSet, Auth, QueryDocument, Condition

parser = ArgumentParser()
parser.add_argument('--connect', required=True)
parser.add_argument('--input_file', required=True)
parser.add_argument('--output_file', required=True)
parser.add_argument('--output_format', choices =['mrc', 'mrk'])
args = parser.parse_args()

DB.connect(args.connect)
output_handle = open(args.output_file, 'w')

def main():
    tbl = Table.from_excel(args.input_file, date_format='%Y%m%d')
    bibs = BibSet()
    tcode_index = {}
    
    for temp_id in tbl.index.keys():
        bib = Bib()

        for field_name in tbl.index[temp_id].keys():
            value = tbl.index[temp_id][field_name]
            
            if field_name == 'Doc Symbol':
                bib.set('191', 'a', value)
            elif field_name == 'Title':
                title_val = value if value else '[Missing title]'      
                bib.set('245', 'a', title_val.title())
            elif field_name == 'publicaion date' or field_name == 'publication date':
                bib.set('269', 'a', value)
                _260c = datetime.strptime(value, '%Y%m%d').strftime('%-d %b. %Y')
                
                for old, new in {'May.': 'May', 'Jun.': 'June', 'Jul.': 'July', 'Sep': 'Sept'}.items():
                    _260c = _260c.replace(old, new)
                    
                bib.set('260', 'c', _260c)
            elif field_name == 'Job Number':
                for job in value.split(';'):
                    bib.set('029', 'a', job, address=['+'])
            elif field_name == 'Lang available':
                langtext = ''
                langs = set([x[0:1] for x in value.split(' ')])
                
                for lang in sorted(langs):        
                    langtext += {'A': 'ara', 'C': 'chi', 'E': 'eng', 'F': 'fre', 'R': 'rus', 'S': 'spa', 'O': 'ger'}.get(lang, '')
                    
                bib.set('041', 'a', langtext)    
            elif field_name =='Tcodes' or field_name == 'tcode':
                for tcode in re.split(r'[,;]', value):
                    if tcode in tcode_index:
                        auth_id = tcode_index[tcode]
                        if auth_id == None: continue
                    else:
                        q = QueryDocument(Condition('035', {'a': tcode}))
                        auth = Auth.find_one(q.compile())
                        
                        if auth:
                            auth_id = auth.id
                            tcode_index[tcode] = auth_id
                        else:
                            logging.warning('Auth record for Tcode "{}" not found'.format(tcode))
                            tcode_index[tcode] = None
                            continue
                    
                    bib.set('650', 'a', auth_id, address=['+'])

        bib.set_008()
        bibs.records.append(bib)
    
    if args.output_format == 'mrk':
        output_handle.write(bibs.to_mrk())
    else:
        output_handle.write(bibs.to_mrc())
    
###

main()
