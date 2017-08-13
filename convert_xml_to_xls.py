# 
#  - Very simple code to convert from office spreadsheet to xls
#	
#	Sneldev.com
#

import sys
import xml.etree.ElementTree as ET
import xlwt
from dateutil.parser import parse

class xml_workbook():
	#hard code a tag here
	_tag_prefix = "{urn:schemas-microsoft-com:office:spreadsheet}"
	
	def _wt(self, key):
		return self._tag_prefix + key
	
	def __init__(self, path):
		tree = ET.parse(path)
		self.root = tree.getroot()
	
	def get_worksheets(self):
		for ws in self.root.findall(self._wt('Worksheet')):
			yield ws.findall(self._wt('Table'))[0]
	
	def get_rows(self, ws):
		return ws.findall(self._wt('Row'))
	
	def get_cells(self, row):
		def create_cell(c):
			data = c.findall(self._wt('Data'))[0]
			text = data.text
			data_type = data.get(self._wt('Type'),"String")
			return {'text' : text, 'type' : data_type}

		cells=row.findall(self._wt('Cell'))
		return [create_cell(c) for c in cells]

#This Small Class convert a xml cell to a text and a style
#as can be put in an xlwt worksheet
class cell_converter():

	def __init__(self):
		self.cell_convert = {
			"String" : lambda value : value,
			"Number" : lambda value : float(value),
			"DateTime" : lambda value : value and parse(value),
		}
		number_style = xlwt.XFStyle()
		number_style.num_format_str = "0.00"
		date_style = xlwt.XFStyle()
		date_style.num_format_str = 'dd/mm/yyyy'
		self.cell_styles = {
			"DateTime" : date_style,
			"Number" : number_style,
		}
		self.default_format = xlwt.XFStyle()
	
	def get_text(self, cell):
		return self.cell_convert[cell['type']](cell['text'])

	def get_style(self, cell):
		return self.cell_styles.get(cell['type'],self.default_format)
		

def convert_xml_spreadsheet_to_xls(in_path, out_path):
	wb_o = xlwt.Workbook(encoding="UTF-8")
	wb = xml_workbook(in_path)
	conv = cell_converter()
	for ws_idx, ws in enumerate(wb.get_worksheets()):
		ws_o = wb_o.add_sheet("Converted Data%d" % ws_idx)
		for row_idx, row in enumerate(wb.get_rows(ws)):
			for cell_idx, cell in enumerate(wb.get_cells(row)):
				ws_o.write(row_idx, cell_idx, conv.get_text(cell), conv.get_style(cell))
	wb_o.save(out_path)


if __name__ == '__main__':
	if (len(sys.argv) < 3):
		print "me.py input_spreadsheet.xml output.xls"
		sys.exit()
	convert_xml_spreadsheet_to_xls(sys.argv[1],sys.argv[2])



