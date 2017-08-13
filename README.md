# pyxmlspreadsheet
Very simplified Office Xml Spreadsheet to Xls converter. 
Written in Python.
It reads the spreadsheet with traditional "xml.etree.ElementTree".
And writes an output file using xlwt. 

I made this script because of this error when trying to open an xml spreadsheet with xlrd (v1.0.0): 

```
    Traceback (most recent call last):
      File "<stdin>", line 1, in <module>
      File "C:\Python27\lib\site-packages\xlrd\__init__.py", line 441, in open_workbook
        ragged_rows=ragged_rows,
      File "C:\Python27\lib\site-packages\xlrd\book.py", line 91, in open_workbook_xls
        biff_version = bk.getbof(XL_WORKBOOK_GLOBALS)
      File "C:\Python27\lib\site-packages\xlrd\book.py", line 1230, in getbof
        bof_error('Expected BOF record; found %r' % self.mem[savpos:savpos+8])
      File "C:\Python27\lib\site-packages\xlrd\book.py", line 1224, in bof_error
        raise XLRDError('Unsupported format, or corrupt file: ' + msg)
    xlrd.biffh.XLRDError: Unsupported format, or corrupt file: Expected BOF record; found '<?xml ve'
```
