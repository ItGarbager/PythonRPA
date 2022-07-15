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
    try:
        tasks = (
            ['mouse_click', 'images/test/1.png', Button.right, 1],
            ['mouse_click', 'images/test/2.png', Button.left, 1],
            ['keyboard_type', None, '6hikgkjkh'],
            ['keyboard_group', None, 'enter'],
            ['mouse_click', 'images/test/3.png', Button.left, 2],
            [None, 'images/test/4.png'],
            ['keyboard_group', None, 'ctrl+a'],
            ['keyboard_type', None, 'sss修改名字哈哈哈'],
            ['keyboard_group', None, 'ctrl+S'],
            ['sleep', None, .5],
            ['keyboard_group', None, 'ctrl+w'],
            ['mouse_click', 'images/test/6.png', Button.left, 1],

        )
        for task in tasks:
            if isinstance(task, list):
                to_do(*task)
            elif isinstance(task, tuple):
                to_do(*task[0], **task[1])

    except:
        logger.error(f'{traceback.format_exc()}')


def test_action2():
    # tasks = [
    #     ('mouse_click', 'images/test2/1.png', Button.left, 2),
    #     ('mouse_click', 'images/test2/2.png', Button.left, 2),
    #     ('mouse_click', 'images/test2/3.png', Button.left, 1),
    #     ('mouse_click', 'images/test2/4.png', Button.left, 1),
    #     ('keyboard_type', None, 'bilibili'),
    #     ('keyboard_group', None, 'enter'),
    #     ('mouse_click', 'images/test2/5.png', Button.left, 1),
    #     [
    #         ('mouse_click', 'images/test2/6.png', Button.left, 1),
    #         {'threshold': .92}
    #     ],
    #     [
    #         (None, 'images/test2/7.png'),
    #         {'': True, '': True}
    #     ],
    #     ('sleep', None, 1.2),
    #     ('mouse_scroll', None, 0, -3900),
    #     [
    #         (
    #             'actions',
    #             None,
    #             [
    #                 [
    #                     ('mouse_click', 'images/test2/8.png', Button.left, 1),
    #                     {
    #                         'loop_limit': 4, 'is_require': False, 'is_many': True,
    #                         'threshold': 0.89,
    #                         # 'is_show': True,
    #                     }
    #                 ],
    #                 ('sleep', None, 1.2),
    #                 ('mouse_scroll', None, 0, -12),
    #                 ('sleep', None, 1.2),
    #             ]
    #         ),
    #         {'do_count': 60}
    #     ],
    #
    # ]
    tasks = TaskList(
        Task('mouse_click', 'images/test2/1.png', Button.left, 2),
        Task('mouse_click', 'images/test2/2.png', Button.left, 2),
        Task('mouse_click', 'images/test2/3.png', Button.left, 1),
        Task('mouse_click', 'images/test2/4.png', Button.left, 1),
        Task('keyboard_type', None, 'bilibili'),
        Task('keyboard_group', None, 'enter'),
        Task('mouse_click', 'images/test2/5.png', Button.left, 1),
        Task('mouse_click', 'images/test2/6.png', Button.left, 1, threshold=.92),
        Task(None, 'images/test2/7.png', Button.left),
        Task('sleep', None, 1.2),
        Task('mouse_scroll', None, 0, -3900),
        Task(
            'actions',
            None,
            TaskList(
                Task(
                    'mouse_click', 'images/test2/11.png', Button.left, 1,
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
    test_action3()
