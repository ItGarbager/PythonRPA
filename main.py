import logging
import warnings

from claess import get_xlsx_tasks
from tools import do_tasks

warnings.filterwarnings('ignore')
logging.basicConfig(
    level=logging.INFO,  # 定义输出到文件的log级别，
    format='%(levelname)-8s[%(asctime)s] %(name)s-%(lineno)d : %(message)s',  # 定义输出log的格式
    datefmt='%y-%m-%d %H:%M:%S',
)
logger = logging.getLogger('RPA')

if __name__ == '__main__':
    logger.info('starting .....')
    do_tasks(get_xlsx_tasks())  # do_tasks 执行自动化任务列表，get_xlsx_tasks 从 excel 文件中读取任务
