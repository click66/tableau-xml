import logging
import sys
import TableauDesktopPy as tdp
from Tableau_Formatter import update_font

DATA_PATH = "/app/_data/"
OUTPUT_PATH = "/app/_output/"

# instantiate logging
file_handler = logging.FileHandler(filename='Logging.log')
file_handler.setFormatter(logging.Formatter('%(asctime)s %(levelname)s - %(message)s'))
stdout_handler = logging.StreamHandler(sys.stdout)
handlers = [file_handler, stdout_handler]
logging.basicConfig(format="%(message)s", level=logging.DEBUG, handlers=handlers)


if __name__ == "__main__":
    path = DATA_PATH + 'Spend Analytics Demo 4.0.twb'
    my_workbook = tdp.Workbook(path)
    my_workbook = update_font(my_workbook, 'Times New Roman', logging=logging)
    my_workbook.save(OUTPUT_PATH + 'Updated.twb')
