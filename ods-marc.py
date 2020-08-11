"""Converts an Excel export of ODS data to MARC for import into Horizon"""

import sys, os
from argparse import ArgumentParser
from dlx import DB
from dlx.util import Table
from dlx.marc import Bib, BibSet, Auth, QueryDocument, Condition

parser = ArgumentParser()
parser.add_argument('--connect', required=True)
parser.add_argument('--file', required=True)
args = parser.parse_args()

DB.connect(args.connect)

def main():
    tbl = Table.from_excel(args.file, date_format='%Y%m%d')
    bibs = BibSet()
    
    for temp_id in tbl.index.keys():
        bib = Bib()

        for field_name in tbl.index[temp_id].keys():
            instance = 0
            value = tbl.index[temp_id][field_name]
            
            if field_name == 'Doc Symbol':
                bib.set('191', 'a', value)
            elif field_name == 'Title':
                bib.set('245', 'a', value)
            elif field_name == 'publicaion date':
                bib.set('269', 'a', value)
            elif field_name == 'Job Number':
                bib.set('029', 'a', value)
            elif field_name == 'Lang available':
                pass
                # todo
            elif field_name =='Tcodes':
                codes = value.split(';')
                
                for code in codes:
                    q = QueryDocument(Condition('035', {'a': code}))  
                    auth = Auth.find_one(q.compile())
                    
                    if auth:
                        bib.set('650', 'a', auth.id)
                    else:
                        raise Exception('Auth record for code "{}" not found'.format(code))
                
        bibs.records.append(bib)
        
    print(bibs.to_mrc())
    
###

main()
