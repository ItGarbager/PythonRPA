import ast
import logging
from typing import Union, Any

import pandas as pd
import win32api
import win32con

logger = logging.getLogger(__name__)


class Task:
    """自动化任务对象"""

    def __init__(self, action, image_path, *args, width=win32api.GetSystemMetrics(win32con.SM_CXVIRTUALSCREEN),
                 height=win32api.GetSystemMetrics(win32con.SM_CYVIRTUALSCREEN),
                 winname='match', threshold=.87, do_count=1, loop_limit=30,
                 is_show=False, is_save=False,
                 is_require=True, is_many=False):
        """

        :param action: 执行动作
        :param image_path: 匹配图片路径
        :param args: 传入动作的参数
        :param width: 截取屏幕宽
        :param height: 截取屏幕高
        :param winname: 若展示匹配效果，展示时的屏幕名字
        :param threshold: 匹配度阈值
        :param do_count: 任务执行次数
        :param loop_limit: 图片匹配次数
        :param is_show: 是否展示匹配效果
        :param is_save: 是否保存匹配图片
        :param is_require: 图片匹配是否为必要条件
        :param is_many: 同张截图中是否要匹配多张图片
        """
        local_arg = locals().copy()
        local_arg.pop('self')
        self.action = local_arg.pop('action')
        self.image_path = local_arg.pop('image_path')
        args = local_arg.pop('args')
        self.args = (self.action, self.image_path) + args
        self.kwargs = local_arg

    def get_data(self):
        if self.kwargs:
            return [self.args, self.kwargs]
        else:
            return self.args


class TaskList:
    """任务生成器"""

    def __init__(self, *task_list: Task):
        """
        传入 Task 类型的列表，构建一个 TaskList 生成器
        :param task_list:
        """
        self.task_list = task_list

    def __iter__(self):
        for task in self.task_list:
            # 返回格式化 Task 数据
            yield task.get_data()

    def __str__(self):
        if hasattr(self, 'name'):
            return f'<TaskList {self.name}>'

    __repr__ = __str__


class Action:
    """事件"""
    nothing = 0
    click_left_button = 1
    click_right_button = 2
    press_left_button = 3
    release_left_button = 4
    scroll_button = 5
    input_str = 6
    press_keyboard = 7
    sleep = 8
    exit = 9
    actions = 10
    move = 11


class Input:
    """根据传入参数构建 TaskList 生成器对象"""
    # 构建一个初始化的类型字典
    __action_dict = {
        Action.nothing: ['', None],  # 未传入事件
        Action.click_left_button: ['mouse_click', None, 'left'],
        Action.click_right_button: ['mouse_click', None, 'right'],
        Action.press_left_button: ['mouse_press', None, 'left'],
        Action.release_left_button: ['mouse_release', None, 'right'],
        Action.scroll_button: ['mouse_scroll', None, 0],
        Action.input_str: ['keyboard_type', None],
        Action.press_keyboard: ['keyboard_group', None],
        Action.sleep: ['sleep', None],
        Action.exit: ['exit', None],
        Action.actions: ['actions', None],
        Action.move: ['mouse_move', None]
    }

    def __init__(
            self, action: Union[Action, int, None] = None, *args, image_path=None,
            excel_path='tasks.xlsx',
            **kwargs
    ):
        # 判断是否仅传了一个参
        if len(args) == 1:
            args = args[0]
            if isinstance(args, float):
                int_arg = int(args)
                args = (args == int_arg) and int_arg or args

            args = (args,)
        if not action:
            action = Action.nothing

        if action == Action.actions:
            # 若为事件集类型则调用 get_xlsx_tasks 函数继续获取相应事件集的 TaskList 生成器对象
            args = (get_xlsx_tasks(excel_path=excel_path, sheet_name=args[0]),)
        __action_dict = self.__action_dict.copy()
        action = __action_dict.get(action)
        # 进行参数构建
        action[1] = image_path
        action = tuple(action)
        if args:
            action += args
        # 构建任务对象
        self.task = Task(*action, **kwargs)


def get_xlsx_tasks(excel_path='tasks.xlsx', sheet_name: Any = 'main') -> TaskList:
    """
    在excel表中读取数据，构建 TaskList 生成器对象
    :param excel_path: excel 表格路径
    :param sheet_name: 选择的 sheet 名称
    :return: TaskList 生成器对象
    """
    df = pd.read_excel(excel_path, names=[
        'action', 'args', 'image_path',
        'threshold', 'do_count', 'loop_limit',
        'is_require', 'is_many', 'is_show'
    ], sheet_name=sheet_name)  # 读取 excel 表格，并给字段命名
    fill_str = '1+-_-空数据-_-!￥%$&'  # 用以填充空数据，可以复杂一些
    df.fillna(fill_str, inplace=True)  # 填充缺失值
    kwargs_list = df.to_dict(orient='records')  # 将 dataframe 转为 dict 字典类型，并去除行索引

    for index, kwargs in enumerate(kwargs_list):
        kwargs_c = kwargs.copy()  # 复制一份用来进行字段处理
        for k, v in kwargs.items():
            if v == fill_str:  # 判断为空则删除该位置参
                del kwargs_c[k]
            else:
                if k == 'args':
                    try:
                        v = ast.literal_eval(v)
                    except ValueError:
                        pass
                    if not isinstance(v, tuple):
                        v = (v,)
                    kwargs_c[k] = v

        kwargs_list[index] = kwargs_c

    # 处理完成后构建任务列表
    task_list = []
    for kwargs in kwargs_list:
        action = kwargs.pop('action')
        args = kwargs.pop('args')
        task_list.append(Input(
            action,
            *args,
            excel_path=excel_path,
            **kwargs
        ).task)

    # 传入当前任务列表构建 TaskList 生成器对象
    tasks = TaskList(
        *task_list
    )
    tasks.name = f'{sheet_name}'
    return tasks
