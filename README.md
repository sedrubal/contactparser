.contact parser
===============

This is a parser for Microsofts `.contact` format to bring contacts to a proper standardized format.
I think many programs do the same job already but I didn't find a comfortable solution.

Output Formats
--------------

 - csv
 - json

Usage
-----

### brief:

`./contactparser.py -o contacts.csv foo.contact bar.contact`

### detailed:

```
usage: contactparser.py [-h] [-v] [-o file] [--json] [--pretty] [--csv]
                        [--csv-dialect {excel,excel-tab,unix}]
                        file [file ...]

a command line tool to convert microsofts .contact files into a csv or json

positional arguments:
  file                  .contact files, - for stdin

optional arguments:
  -h, --help            show this help message and exit
  -v, --verbose         prints debug messages to stderr, -vv for more)
  -o file, --output file
                        the output file, - for stdout

output: json:
  --json                output format is json
  --pretty              Make json pretty and not compact

output: csv:
  --csv                 output format is csv
  --csv-dialect {excel,excel-tab,unix}
                        the csv dialect
```

Install
-------

```sh
./setup.py install
```

Contribute
----------

 - **Currently only Python3 is supported**. If you need python2, please contribute. I think there is not much to change.
 - Currently there are not many fields supported. Please contribute to make this tool even mightier.
 - Currently only csv and json export is supported. One could also implement vcard export.

License
-------

[GPLv3](LICENSE)
