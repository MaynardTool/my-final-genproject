from distutils.core import setup
import py2exe, sys, os

sys.path.append(r'C:\\Users\\Administrator\\Desktop\\Project\\autoinstall\\source\\version-1.0\\pycfg\\gcti_cfg\\lib.windows-x86-2.7')

from glob import glob
data_files = [("lib.windows-x86-2.7", glob(r'C:\\Users\\Administrator\\Desktop\\Project\\autoinstall\\source\\version-1.0\\pycfg\\gcti_cfg\\lib.windows-x86-2.7\\conflib.pyd'))]

images_and_xml = [("lib\\logos", glob(r'C:\\Users\\Administrator\\Desktop\\Project\\autoinstall\\source\\version-1.0\\lib\\logos\\*.*')), 
("lib\\icons", glob(r'C:\\Users\\Administrator\\Desktop\\Project\\autoinstall\\source\\version-1.0\\lib\icons\\*.*')),
("", glob(r'C:\\Users\\Administrator\\Desktop\\Project\\autoinstall\\source\\version-1.0\\auto-install.xml'))]

data_files.extend(images_and_xml)

setup(
	windows = [{
            "script": "final-auto-install.py",
            "icon_resources": [(1, "C:\\Users\\Administrator\\Desktop\\Project\\autoinstall\\source\\version-1.0\\lib\icons\\app_icon.ico")],
			}],
	data_files = data_files,
	zipfile = None,
	
	options={
				"py2exe":{
						"bundle_files": 1,
						"compressed" : True,
						"dll_excludes": ["tcl85.dll", "tk85.dll", "crypt32.dll", "mpr.dll", "mswsock.dll", "powrprof.dll"],
						"unbuffered": True,
						"optimize": 2,
						"excludes": ["email"]
					}
			}
)
