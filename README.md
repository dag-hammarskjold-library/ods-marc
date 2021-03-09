#### Installation
From the command line:
```bash
pip install -r requirements.txt
pip install git+https://github.com/dag-hammarskjold-library/ods-marc
```

#### Scripts
> #### ods-marc
Converts an Excel export of ODS data into MARC for import into Horizon

Usage (command line):
```bash
ods-marc --help
```
```bash
ods-marc --connect=<MDB connection string> --input_file=<path to ODS Excel file>, --output_format=mrc
```
```bash
ods-marc --connect=<MDB connection string> --input_file=<path to ODS Excel file> --output_format=mrk --output_file=my_file.mrk
```
