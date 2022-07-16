import logging
import traceback
import warnings

import pandas as pd
from pynput.mouse import Button

from claess import Task, TaskList, Input
from tools import to_do, do_tasks

warnings.filterwarnings('ignore')
logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)


def test_action1():
    tasks = (
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
        Task('mouse_click', 'images/test2/1.png', 'left', 2),
        Task('mouse_click', 'images/test2/2.png', 'left', 2),
        Task('mouse_click', 'images/test2/3.png', 'left', 1),
        Task('mouse_click', 'images/test2/4.png', 'left', 1),
        Task('keyboard_type', None, 'bilibili'),
        Task('keyboard_group', None, 'enter'),
        Task('mouse_click', 'images/test2/5.png', 'left', 1),
        Task('mouse_click', 'images/test2/6.png', 'left', 1, threshold=.92),
        Task(None, 'images/test2/7.png', 'left'),
        Task('sleep', None, 1.2),
        Task('mouse_scroll', None, 0, -3900),
        Task(
            'actions',
            None,
            TaskList(
                Task(
                    'mouse_click', 'images/test2/11.png', 'left', 1,
                    loop_limit=4, threshold=0.89,
                    is_require=False, is_many=True,
                    is_show=True,
                ),
                Task('sleep', None, 1.2),
                Task('mouse_scroll', None, 0, -12),
                Task('sleep', None, 1.2),
            ),
            do_count=60
        ),
    )

    do_tasks(tasks)


def test_action3():
    df = pd.read_excel('tasks.xlsx', names=[
        'action', 'arg', 'image', 'kwargs',
        'threshold', 'do_count', 'loop_limit',
        'is_require', 'is_many'
    ])
    df.fillna(0, inplace=True)

    task_list = (
        Input(action, arg, image_path=image, **eval(kwargs.replace('TRUE', 'True').replace('FALSE', 'False'))).task
        for action, arg, image, kwargs, _, _, _, _, _ in df.values.tolist()
    )
    tasks = TaskList(
        *task_list
        # Input(Action.click_left_button, 1, image_path='test.png').task,
        # Input(Action.actions, TaskList(
        #     Task('sleep', 1.2)
        # ), is_many=True).task
    )
    do_tasks(tasks)


if __name__ == '__main__':
    test_action1()
