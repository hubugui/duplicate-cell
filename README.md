duplicate-cell
==============

find duplicate cell from different excel files.

currently only support the first sheet.

step
==============

1. unzip xlrd-0.6.1.zip
2. cd xlrd-0.6.1
3. python setup.py install
4. cd ..
5. python duplicate.py -f "a.xls b.xls c.xls" -c 1,4 -p 1
6. python duplicate.py -d /home/your/work -p 2