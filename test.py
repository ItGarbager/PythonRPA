import logging
import warnings

from classes import Task, TaskList, Input, Action
from tools import do_tasks

warnings.filterwarnings('ignore')
logging.basicConfig(
    level=logging.INFO,  # 定义输出到文件的log级别，
    format='%(levelname)-8s[%(asctime)s] %(name)s-%(lineno)d : %(message)s',  # 定义输出log的格式
    datefmt='%y-%m-%d %H:%M:%S',
)
logger = logging.getLogger(__name__)


def test_action1():
    tasks = (
        Task('mouse_click', None, 'right', 1).get_data(),  # 单击右键
        Task(None, 'images/test/0.png').get_data(),  # 移动至图片处
        Task('mouse_click', 'images/test/-1.png', 'left', 1).get_data(),  # 左键单击图所在的位置
        Input(Action.input.mouse_move, (100, 100), True).task.get_data(),  # 向下向右移动鼠标100
        ('mouse_click', 'images/test/1.png', 'right', 1),  # 右键单击图所在的位置
        ('mouse_click', 'images/test/2.png', 'left', 1),  # 左键单击图所在的位置
        ('keyboard_type', None, '6hikgkjkh'),  # 输入文本
        ('keyboard_group', None, 'enter'),  # 敲击回车
        ('mouse_click', 'images/test/3.png', 'left', 2),  # 左键双击图所在的位置
        (None, 'images/test/4.png'),  # 无动作，表示等待匹配图片出现
        ('keyboard_group', None, 'ctrl+a'),  # 输入组合键，此处为全选，忽略大小写
        ('keyboard_type', None, 'sss修改名字哈哈哈'),  # 输入文本
        ('keyboard_group', None, 'ctrl+S'),  # 输入组合键，此处为保存
        ('sleep', None, .5),  # 休眠 0.5 s
        ('keyboard_group', None, 'ctrl+w'),  # 输入组合键，此处为关闭当前页
        [('mouse_click', 'images/test/6.png', 'left', 1), {'is_show': True}],  # 左键单击图所在位置，并展示图片
        ('sleep', None, 3),  # 休眠 3 s

    )
    do_tasks(tasks)


def test_action2():
    tasks = TaskList(
        Task(Action.task.mouse_click, 'images/test2/1.png', 'left', 2),
        Task(Action.task.mouse_click, 'images/test2/2.png', 'left', 2),
        Task(Action.task.mouse_click, 'images/test2/3.png', 'left', 1),
        Task(Action.task.mouse_click, 'images/test2/4.png', 'left', 1),
        Task(Action.task.keyboard_type, None, 'bilibili'),
        Task(Action.task.keyboard_group, None, 'enter'),
        Task(Action.task.mouse_click, 'images/test2/5.png', 'left', 1),
        Task(Action.task.mouse_click, 'images/test2/6.png', 'left', 1, threshold=.92),
        Task(Action.task.nothing, 'images/test2/7.png', 'left'),
        Task(Action.task.sleep, None, 1.2),
        Task(Action.task.mouse_scroll, None, 0, -3900),
        Task(
            Action.task.actions,
            None,
            TaskList(
                Task(
                    Action.task.mouse_click, 'images/test2/11.png', 'left', 1,
                    loop_limit=4, threshold=0.89,
                    is_require=False, is_many=True,
                    is_show=True,
                ),
                Task(Action.task.sleep, None, 1.2),
                Task(Action.task.mouse_scroll, None, 0, -12),
                Task(Action.task.sleep, None, 1.2),
            ),
            do_count=60
        ),
    )

    do_tasks(tasks)


if __name__ == '__main__':
    test_action2()
