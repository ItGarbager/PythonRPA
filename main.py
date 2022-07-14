import logging
import warnings

from claess import get_xlsx_tasks
from tools import do_tasks

warnings.filterwarnings('ignore')
logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)

if __name__ == '__main__':
    do_tasks(get_xlsx_tasks())
