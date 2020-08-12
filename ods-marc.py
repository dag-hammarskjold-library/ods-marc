"""Converts an Excel export of ODS data to MARC for import into Horizon"""

import sys, os, logging
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
    
    for temp_id in tbl.index.keys():
        bib = Bib()

        for field_name in tbl.index[temp_id].keys():
            value = tbl.index[temp_id][field_name]
            
            if field_name == 'Doc Symbol':
                bib.set('191', 'a', value)
            elif field_name == 'Title':
                if not value:
                    logging.warn('Title not found for row {}'.format(temp_id + 2))
                    
                bib.set('245', 'a', value)
            elif field_name == 'publicaion date':
                bib.set('269', 'a', value)
            elif field_name == 'Job Number':
                for job in value.split(';'):
                    bib.set('029', 'a', job, address=['+'])
            elif field_name == 'Lang available':
                langtext = ''
                
                for lang in sorted(value.split(' ')):
                    langtext += {'AP': 'ara', 'EP': 'eng'}[lang]
                    
                bib.set('041', 'a', langtext)    
            elif field_name =='Tcodes':
                for tcode in value.split(';'):
                    q = QueryDocument(Condition('035', {'a': tcode}))
                    auth = Auth.find_one(q.compile())
                    
                    if auth:
                        bib.set('650', 'a', auth.id, address=['+'])
                    else:
                        logging.warning('Auth record for Tcode "{}" not found'.format(tcode))
                        
        bib.set_008()
        bibs.records.append(bib)
    
        if args.output_format == 'mrk':
            output_handle.write(bibs.to_mrk())
        else:
            output_handle.write(bib.to_mrc())
    
###

main()
