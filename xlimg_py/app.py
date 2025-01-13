from excel_paster import ExcelPaster
from dirtree import get_dirtree
import sys
from pathlib import Path    

def main(textfile_path):
    dir_tree = get_dirtree(textfile_path)
    paster = ExcelPaster()
    paster.run(dir_tree)

if __name__ == '__main__':
    textfile_path = Path.cwd() / Path(sys.argv[1])
    main(textfile_path)