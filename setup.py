from distutils.core import setup
import py2exe

setup(
	windows=['C:\\Python27\\AutomaticDocumentEditor\\gui_tkinter.py'],
    options={
        'py2exe': 
        {
            'includes': ['lxml.etree', 'lxml._elementpath', 'gzip', 'docx'],
        }
    }
)