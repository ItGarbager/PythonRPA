from typing import Union

import pandas as pd


class Task:
    """自动化任务对象"""

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


class Input:
    """根据传入参数构建 TaskList 生成器对象"""
    # 构建一个初始化的类型字典
    __action_dict = {
        Action.nothing: ['', None, None],  # 未传入事件
        Action.click_left_button: ['mouse_click', None, 'left', None],
        Action.click_right_button: ['mouse_click', None, 'right', None],
        Action.press_left_button: ['mouse_press', None, 'left'],
        Action.release_left_button: ['mouse_release', None, 'right'],
        Action.scroll_button: ['mouse_scroll', None, 0, None],
        Action.input_str: ['keyboard_type', None, None],
        Action.press_keyboard: ['keyboard_group', None, None],
        Action.sleep: ['sleep', None, None],
        Action.exit: ['exit', None, None],
        Action.actions: ['actions', None, None],
    }

    def __init__(
            self, action: Union[Action, int, None] = None, arg=None, image_path=None,
            excel_path='tasks.xlsx',
            **kwargs
    ):
        if not action:
            action = Action.nothing
        if isinstance(arg, float):
            int_arg = int(arg)
            arg = (arg == int_arg) and int_arg or arg

        if action == Action.actions:
            # 若为事件集类型则调用 get_xlsx_tasks 函数继续获取相应事件集的 TaskList 生成器对象
            arg = get_xlsx_tasks(excel_path=excel_path, sheet_name=arg)

        action = self.__action_dict.get(action)
        # 进行参数构建
        if action[-1] is None:
            action[-1] = arg
        action[1] = image_path
        action = tuple(action)
        # 构建任务对象
        self.task = Task(*action, **kwargs)


def get_xlsx_tasks(excel_path='tasks.xlsx', sheet_name='main') -> TaskList:
    """
    在excel表中读取数据，构建 TaskList 生成器对象
    :param excel_path: excel 表格路径
    :param sheet_name: 选择的 sheet 名称
    :return: TaskList 生成器对象
    """
    df = pd.read_excel(excel_path, names=[
        'action', 'arg', 'image_path',
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
        kwargs_list[index] = kwargs_c

    # 处理完成后构建任务列表
    task_list = (
        Input(
            excel_path=excel_path,
            **kwargs
        ).task
        for kwargs in kwargs_list
    )
    # 传入当前任务列表构建 TaskList 生成器对象
    tasks = TaskList(
        *task_list
    )
    tasks.name = f'{sheet_name}'
    return tasks
