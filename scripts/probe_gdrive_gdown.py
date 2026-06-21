"""List Drive folder via gdown internals."""
import gdown
from gdown.download_folder import _get_directory_structure

url = 'https://drive.google.com/drive/folders/18uMCSYNDgipvGzHwlhyMcxrBaAcjChSZ?usp=sharing'
try:
    structure = _get_directory_structure(url, quiet=False)
    print('structure type', type(structure))
    print(structure)
except Exception as e:
    print('err', type(e).__name__, e)
