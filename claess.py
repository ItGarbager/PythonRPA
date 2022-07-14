from typing import Union

import pandas as pd
from pynput.mouse import Button


class KeyDefaults:
    winname = 'sss'
    width = 1920
    height = 1080
    threshold = 0.87
    do_count = 1
    loop_limit = 30
    is_show = False
    is_label = False
    is_save = False
    is_require = True
    is_many = False


class Task:

    def __init__(self, action, filepath, *args, **kwargs):
        self.action = action
        self.filepath = filepath
        self.args = (self.action, self.filepath) + args
        self.kwargs = kwargs

    def get_data(self):

        if self.kwargs:
            return [self.args, self.kwargs]
        else:
            return self.args


class TaskList:

    def __init__(self, *task_list: Task):
        self.task_list = task_list

    def __iter__(self):
        for task in self.task_list:
            yield task.get_data()


class Action:
    click_left_button = 1
    click_right_button = 2
    press_left_button = 3
    release_left_button = 4
    scroll_button = 5
    input_str = 6
    press_keyboard = 7
    sleep = 8
    actions = 9


class Input:
    __action_dict = {
        1: ['mouse_click', None, Button.left],
        2: ['mouse_click', None, Button.right],
        3: ['mouse_press', None, Button.left],
        4: ['mouse_release', None, Button.left],
        5: ['mouse_scroll', None, 0],
        6: ['keyboard_type', None],
        7: ['keyboard_group', None],
        8: ['sleep', None],
        9: ['actions', None],
        None: [None, None]
    }

    def __init__(self, action: Union[Action, int, None] = None, *args, image_path=None, **kwargs):
        list_args = list(args)
        for index, arg in enumerate(list_args):
            if isinstance(arg, float):
                list_args[index] = int(arg)
        if action == 9:
            list_args[0] = get_xlsx_tasks(excel_path='tasks.xlsx', sheet_name=args[0])

        args = tuple(list_args)
        action = self.__action_dict.get(action)

        action[1] = image_path
        action = tuple(action)

        self.task = Task(*(action + args), **kwargs)


def get_xlsx_tasks(excel_path='tasks.xlsx', sheet_name='main'):
    df = pd.read_excel(excel_path, names=[
        'action', 'arg', 'image', 'kwargs',
        'threshold', 'do_count', 'loop_limit',
        'is_require', 'is_many'
    ], sheet_name=sheet_name)
    df.fillna('', inplace=True)
    task_list = (
        Input(
            action, arg, image_path=image,
            **eval("{" + kwargs.replace('TRUE', 'True').replace('FALSE', 'False') + "}")
        ).task
        for action, arg, image, kwargs, _, _, _, _, _ in df.values.tolist()
    )
    tasks = TaskList(
        *task_list
    )
    return tasks
