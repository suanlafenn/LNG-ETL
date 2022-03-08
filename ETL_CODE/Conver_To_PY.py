srcFolder = r'input_folderpath_here'
desFolder = r'output_folderpath_here'

import os
import nbformat
from nbconvert import PythonExporter

def convertNotebook(notebookPath, modulePath):
    with open(notebookPath) as fh:
        nb = nbformat.reads(fh.read(), nbformat.NO_CONVERT)
    exporter = PythonExporter()
    source, meta = exporter.from_notebook_node(nb)
    with open(modulePath, 'w+') as fh:
        fh.writelines(source)

# For folder creation if doesn't exist
if not os.path.exists(desFolder):
    os.makedirs(desFolder)

for file in os.listdir(srcFolder):
    if os.path.isdir(srcFolder + '\\' + file):
        continue
    if ".ipynb" in file:
        convertNotebook(srcFolder + '\\' + file, desFolder + '\\' + file[:-5] + "py")



#command
python -m PyInstaller -F -i download.ico cpp_west_market_app.py --hidden-import win32timezone
python -m PyInstaller -F -i download.ico --add-binary "C:/Users/lxmz/AppData/Local/Packages/PythonSoftwareFoundation.Python.3.10_qbz5n2kfra8p0/LocalCache/local-packages/Python310/site-packages/tabula/tabula-1.0.5-jar-with-dependencies.jar;./tabula/" SIMPSON_SPENCE.py --hidden-import win32timezone
