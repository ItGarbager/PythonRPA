import logging
import math
import os
import traceback
from time import sleep, time
from typing import List

import clipboard
import cv2
import numpy as np
import pynput
import win32api
import win32con
import win32gui
import win32ui
from pynput.keyboard import Key

from classes import MatchImageError

logger = logging.getLogger(__name__)


def mouse_move(pos, curr=False):
    if curr:
        current_pos = win32api.GetCursorPos()
        pos = (np.array(current_pos) + np.array(pos)).tolist()
    win32api.SetCursorPos(pos)


def mouse_press(key='left'):
    if key == 'left':
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    elif key == 'right':
        win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0)


def mouse_release(key='left'):
    if key == 'left':
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
    elif key == 'right':
        win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0)


def mouse_click(key='left', count=1):  # 模拟鼠标点击
    for i in range(count):
        mouse_press(key)
        mouse_release(key)


def mouse_scroll(dx, dy):
    if dx:
        win32api.mouse_event(win32con.MOUSEEVENTF_HWHEEL, 0, 0, dx)
    if dy:
        win32api.mouse_event(win32con.MOUSEEVENTF_WHEEL, 0, 0, dy)


# -------------- 重写 type
key_input = getattr(pynput.keyboard.Controller, 'type')


def key_type(self, key, *args, **kwargs):
    """
    重写 pynput.keyboard.Controller.type 方法
    :param self: Controller 对象
    :param key: 调用 type 时传入的第一个位置参
    :param args:
    :param kwargs:
    """
    if not isinstance(key, list):  # 当 key 类型不为 list 时，代表传入的为字符串， list 为输入的具体的按键
        clipboard.copy(key)  # 调用 ctrl+c 复制内容
        controlKey = Key.ctrl if os.name == 'nt' else Key.cmd
        # 模拟 ctrl+v 输入信息
        with self.pressed(controlKey):
            self.press('v')
            self.release('v')
            # self.release(Key.enter)  # 输入完成后默认跟一个回车，可选择删掉
    else:
        key_input(self, key, *args, **kwargs)


pynput.keyboard.Controller.type = key_type  # -------------- 重写 type


# -------------- 组合键
def key_group(self, keys: str):
    """
    实现组合键的方法
    :param self: Controller 对象
    :param keys: 组合键的字符串
    :return:
    """
    keys = keys.split('+')
    len_keys = len(keys)
    if len_keys > 3:
        raise Exception('仅支持最多3个键')
    if len_keys < 0:
        raise Exception('未输入按键')
    for index, key in enumerate(keys):
        key = key.lower()  # 统一小写
        is_have = hasattr(Key, key)
        if is_have:
            keys[index] = getattr(Key, key)
        else:
            keys[index] = key
    # 组合键的实现
    if len_keys == 1:
        self.press(keys[0])
        self.release(keys[0])
    elif len_keys == 2:
        with self.pressed(keys[0]):
            self.press(keys[1])
            self.release(keys[1])
    else:
        with self.pressed(keys[0]):
            with self.pressed(keys[1]):
                self.press(keys[2])
                self.release(keys[2])


pynput.keyboard.Controller.group = key_group  # -------------- 组合键


def exit_all(*args, **kwargs):
    raise SystemExit('exit...')


def grab_screen(region=None):
    """屏幕截图"""
    hwin = win32gui.GetDesktopWindow()

    if region:
        left, top, width, height = region
    else:
        width = win32api.GetSystemMetrics(win32con.SM_CXVIRTUALSCREEN)
        height = win32api.GetSystemMetrics(win32con.SM_CYVIRTUALSCREEN)
        left = win32api.GetSystemMetrics(win32con.SM_XVIRTUALSCREEN)
        top = win32api.GetSystemMetrics(win32con.SM_YVIRTUALSCREEN)

    hwindc = win32gui.GetWindowDC(hwin)
    srcdc = win32ui.CreateDCFromHandle(hwindc)
    memdc = srcdc.CreateCompatibleDC()

    bmp = win32ui.CreateBitmap()
    bmp.CreateCompatibleBitmap(srcdc, width, height)
    memdc.SelectObject(bmp)
    memdc.BitBlt((0, 0), (width, height), srcdc, (left, top), win32con.SRCCOPY)
    signedIntsArray = bmp.GetBitmapBits(True)

    img = np.fromstring(signedIntsArray, dtype='uint8')
    img.shape = (height, width, 4)

    srcdc.DeleteDC()
    memdc.DeleteDC()
    win32gui.ReleaseDC(hwin, hwindc)
    win32gui.DeleteObject(bmp.GetHandle())
    return img  # , cv2.cvtColor(img, cv2.COLOR_BGRA2RGB)


def py_nms(dets, thresh):
    """Pure Python NMS baseline."""
    # x1、y1、x2、y2、以及score赋值
    # （x1、y1）（x2、y2）为box的左上和右下角标
    x1 = dets[:, 0]
    y1 = dets[:, 1]
    x2 = dets[:, 2]
    y2 = dets[:, 3]
    scores = dets[:, 4]
    # 每一个候选框的面积
    areas = (x2 - x1 + 1) * (y2 - y1 + 1)
    # order是按照score降序排序的
    order = scores.argsort()[::-1]
    # print("order:",order)

    keep = []
    while order.size > 0:
        i = order[0]
        keep.append(i)
        # 计算当前概率最大矩形框与其他矩形框的相交框的坐标，会用到numpy的broadcast机制，得到的是向量
        xx1 = np.maximum(x1[i], x1[order[1:]])
        yy1 = np.maximum(y1[i], y1[order[1:]])
        xx2 = np.minimum(x2[i], x2[order[1:]])
        yy2 = np.minimum(y2[i], y2[order[1:]])
        # 计算相交框的面积,注意矩形框不相交时w或h算出来会是负数，用0代替
        w = np.maximum(0.0, xx2 - xx1 + 1)
        h = np.maximum(0.0, yy2 - yy1 + 1)
        inter = w * h
        # 计算重叠度IOU：重叠面积/（面积1+面积2-重叠面积）
        ovr = inter / (areas[i] + areas[order[1:]] - inter)
        # 找到重叠度不高于阈值的矩形框索引
        inds = np.where(ovr <= thresh)[0]
        # print("inds:",inds)
        # 将order序列更新，由于前面得到的矩形框索引要比矩形框在原order序列中的索引小1，所以要把这个1加回来
        order = order[inds + 1]
    return keep


def match_template(img_gray, template_img, template_threshold):
    """
    img_gray:待检测的灰度图片格式
    template_img:模板小图，也是灰度化了
    template_threshold:模板匹配的置信度
    """

    h, w = template_img.shape[:2]
    res = cv2.matchTemplate(img_gray, template_img, cv2.TM_CCOEFF_NORMED)
    loc = np.where(res >= template_threshold)  # 大于模板阈值的目标坐标
    score = res[res >= template_threshold]  # 大于模板阈值的目标置信度
    # 将模板数据坐标进行处理成左上角、右下角的格式
    xmin = np.array(loc[1])
    ymin = np.array(loc[0])
    xmax = xmin + w
    ymax = ymin + h
    xmin = xmin.reshape(-1, 1)  # 变成n行1列维度
    xmax = xmax.reshape(-1, 1)  # 变成n行1列维度
    ymax = ymax.reshape(-1, 1)  # 变成n行1列维度
    ymin = ymin.reshape(-1, 1)  # 变成n行1列维度
    score = score.reshape(-1, 1)  # 变成n行1列维度
    data_hlist = [xmin, ymin, xmax, ymax, score]
    data_hstack = np.hstack(data_hlist)  # 将xmin、ymin、xmax、yamx、scores按照列进行拼接
    thresh = 0.3  # NMS里面的IOU交互比阈值

    keep_dets = py_nms(data_hstack, thresh)
    dets = data_hstack[keep_dets]  # 最终的nms获得的矩形框
    return dets


def match_image(image_path, img_rgb, threshold=0.95, is_show=False, is_many=False) -> List:
    """
    opencv 实现图片匹配
    :param image_path: 图片路径
    :param img_rgb: 背景图，此处默认传入全屏
    :param threshold: 匹配度
    :param is_show: 是否展示
    :param is_many: 是否匹配多个目标
    :return: 匹配到的目标坐标的列表
    """

    def fo(left_top=None, right_bottom=None):
        if right_bottom is None:
            right_bottom = (left_top[0] + w, left_top[1] + h)  # 右下角
        middles.append((left_top[0] + math.floor(w / 2), left_top[1] + math.floor(h / 2)))  # 中间的位置
        if is_show:
            cv2.rectangle(img_rgb, left_top, right_bottom, (0, 0, 255), 2)  # 画出矩形位置

    template = cv2.imread(image_path, 0)
    h, w = template.shape[:2]
    img_gray = cv2.cvtColor(img_rgb, cv2.COLOR_BGR2GRAY)
    middles = []

    if is_many:
        dets = match_template(img_gray, template, threshold)
        # np.where返回的坐标值(x,y)是(h,w)，注意h,w的顺序
        for coord in dets:
            left_top = (int(coord[0]), int(coord[1]))
            fo(left_top, right_bottom=(int(coord[2]), int(coord[3])))

    else:
        res = cv2.matchTemplate(img_gray, template, cv2.TM_CCOEFF_NORMED)
        min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(res)
        # 取匹配程度大于 threshold 的坐标
        if max_val > threshold:
            left_top = max_loc  # 左上角
            fo(left_top)
    return middles


def do_tasks(tasks):
    """
    执行自动化任务列表
    :param tasks: 任务列表
    :return:
    """
    try:
        for task in tasks:

            if isinstance(task, tuple):
                to_do(*task)
            elif isinstance(task, list):
                to_do(*task[0], **task[1])
    except SystemExit:
        exit(1)
    except KeyboardInterrupt:
        exit(1)
    except MatchImageError as e:
        logger.error(f'{e}')
    except:
        logger.error(f'{traceback.format_exc()}')


def to_do(
        action=None, image_path=None,
        *args,
        width=win32api.GetSystemMetrics(win32con.SM_CXVIRTUALSCREEN),
        height=win32api.GetSystemMetrics(win32con.SM_CYVIRTUALSCREEN),
        winname='match', threshold=.87, do_count=1, loop_limit=30,
        is_show=False, is_save=False,
        is_require=True, is_many=False,
        **kwargs
):
    """执行任务"""
    do_c = 0
    while do_count == -1 or do_c < do_count:
        do_c += 1
        logger.info(f'{action or "":<25} round {do_c}')
        # 存在图片
        middles = None
        if image_path:
            c = 0
            while True:
                c += 1
                img = grab_screen((0, 0, width, height))
                middles = match_image(image_path, img, is_show=is_show, threshold=threshold, is_many=is_many)
                if middles or c > loop_limit:
                    break
                else:
                    sleep(0.05)
                    logger.warning(f'{image_path:<25} retry {c}')
            if not middles:
                if is_require:
                    raise MatchImageError(f'Match image failed, please check image {image_path}')
                else:
                    return
            logger.info(f'matched {image_path} ==> {middles}')

            if is_show:
                cv2.namedWindow(winname, cv2.WINDOW_NORMAL)
                cv2.imshow(winname, img)
                cv2.resizeWindow(winname, width // 3, height // 3)
                hwnd = win32gui.FindWindow(None, winname)
                win32gui.ShowWindow(hwnd, win32con.SW_SHOWNORMAL)

                win32gui.SetWindowPos(
                    hwnd, win32con.HWND_TOPMOST, 0, 0, 0, 0,
                    win32con.SWP_NOMOVE | win32con.SWP_NOACTIVATE | win32con.SWP_NOOWNERZORDER |
                    win32con.SWP_SHOWWINDOW | win32con.SWP_NOSIZE
                )

                k = cv2.waitKey(1)
                if k % 256 == 27:
                    cv2.destroyWindow(winname=winname)
                    return
            if is_save:
                cv2.imwrite(f'save_{int(time())}.png', img)

        if action:
            if action.startswith('mouse_'):
                func = globals().get(action, None)

            elif action.startswith('keyboard_'):
                action_type = action[len('keyboard_'):]
                controller = pynput.keyboard.Controller()
                func = getattr(controller, action_type, None)
            elif action == 'sleep':
                sleep(*args, **kwargs)
                continue
            elif action == 'actions':
                format_arg = 'load ' + str(args[-1])
                logger.info(f'{format_arg:<25} success')
                do_tasks(*args, **kwargs)
                continue
            elif action == 'exit':
                func = exit_all
            else:
                continue
            if middles is None:
                if func:
                    func(*args, **kwargs)

            else:
                for middle in middles:
                    mouse_move(middle)
                    sleep(0.03)
                    if func:
                        func(*args, **kwargs)
        else:
            if middles is not None:
                for middle in middles:
                    mouse_move(middle)
