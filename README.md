xls2csv-py
---

Converts file to and from .CSV/.XLS(X) extension formats.

Requires **xlrd** and **xlwt** libraries (available from PyPI).

```
usage: xls2csv [-h] [-o OUTPUT] [-q {0,1,2,3}] [-e ENCODING]
               input

positional arguments:
  input                 input file name

optional arguments:
  -h, --help            show this help message and exit
  -o OUTPUT, --output OUTPUT
                        output file or folder name
  -d DELIMITER, --delimiter DELIMITER
                        column field delimiter
  -q {0,1,2,3}, --quoting {0,1,2,3}
                        text quoting {0: 'minimal', 1: 'all',
                        2: 'non-numeric', 3: 'none'}
  -e ENCODING, --encoding ENCODING
                        file encoding (default: utf-8)
```