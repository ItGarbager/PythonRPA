import logging
import math
import os
import traceback
from time import sleep, time

import clipboard
import cv2
import numpy as np
import pynput
import win32api
import win32con
import win32gui
import win32ui
from pynput.keyboard import Key

logger = logging.getLogger(__name__)

# -------------- 重写 type
key_input = getattr(pynput.keyboard.Controller, 'type')


def key_type(self, key, *args, **kwargs):
    if isinstance(key, str):
        clipboard.copy(key)
        controlKey = Key.ctrl if os.name == 'nt' else Key.cmd
        with self.pressed(controlKey):
            self.press('v')
            self.release('v')
    else:
        key_input(self, key, *args, **kwargs)


pynput.keyboard.Controller.type = key_type  # -------------- 重写 type


# -------------- 组合键
def key_group(self, keys: str):
    keys = keys.split('+')
    len_keys = len(keys)
    if len_keys > 3 or len_keys < 0:
        return
    for index, key in enumerate(keys):
        key = key.lower()
        is_have = hasattr(Key, key)
        if is_have:
            keys[index] = getattr(Key, key)
        else:
            keys[index] = key
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
                self.release(keys[3])


pynput.keyboard.Controller.group = key_group  # -------------- 组合键


def grab_screen(region=None):
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


def match_image(image_path, img_rgb, threshold=0.95, is_label=False, is_many=False):
    def fo():
        right_bottom = (left_top[0] + w, left_top[1] + h)  # 右下角
        middles.append((left_top[0] + math.floor(w / 2), left_top[1] + math.floor(h / 2)))  # 中间的位置
        if is_label:
            cv2.rectangle(img_rgb, left_top, right_bottom, (0, 0, 255), 2)  # 画出矩形位置

    template = cv2.imread(image_path, 0)
    h, w = template.shape[:2]
    img_gray = cv2.cvtColor(img_rgb, cv2.COLOR_BGR2GRAY)
    res = cv2.matchTemplate(img_gray, template, cv2.TM_CCOEFF_NORMED)
    middles = []

    if is_many:
        loc = np.where(res >= threshold)
        # np.where返回的坐标值(x,y)是(h,w)，注意h,w的顺序
        for left_top in zip(*loc[::-1]):
            fo()

    else:
        min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(res)
        # 取匹配程度大于 threshold 的坐标
        if max_val > threshold:
            left_top = max_loc  # 左上角
            fo()
    return middles


def do_tasks(tasks):
    try:
        for task in tasks:
            if isinstance(task, tuple):
                to_do(*task)
            elif isinstance(task, list):
                to_do(*task[0], **task[1])
    except:
        logger.error(f'{traceback.format_exc()}')


def to_do(
        action=None, image_path=None,
        *args,
        width=win32api.GetSystemMetrics(win32con.SM_CXVIRTUALSCREEN),
        height=win32api.GetSystemMetrics(win32con.SM_CYVIRTUALSCREEN),
        winname='match', threshold=.87, do_count=1, loop_limit=30,
        is_show=False, is_label=False,
        is_save=False, is_require=True, is_many=False,
        **kwargs
):
    for i in range(do_count):
        logger.debug(f'{action}-{image_path} Task execution round {i + 1}')
        # 存在图片
        middles = None
        if image_path:
            c = 0
            while True:
                c += 1
                img = grab_screen((0, 0, width, height))
                middles = match_image(image_path, img, is_label=is_label, threshold=threshold, is_many=is_many)
                if middles or c > loop_limit:
                    break
                else:
                    sleep(.2)
                    logger.warning(f'{image_path} retry {c} time{"s" if c > 1 else ""}')
            if not middles:
                if is_require:
                    logger.error(f'Match image failed, please check image {image_path}')
                    exit(1)
                else:
                    return
            logger.debug(f'matched {image_path} ==> {middles}')

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

        ctr = pynput.mouse.Controller()
        if action:

            if action.startswith('mouse_'):
                action_type = action[len('mouse_'):]
                controller = ctr
            elif action.startswith('keyboard_'):
                action_type = action[len('keyboard_'):]
                controller = pynput.keyboard.Controller()
            elif action == 'sleep':
                sleep(*args, **kwargs)
                continue
            elif action == 'actions':
                do_tasks(*args, **kwargs)
                continue
            else:
                continue
            if middles is None:
                is_have = hasattr(controller, action_type)
                if is_have:
                    func = getattr(controller, action_type)

                    func(*args, **kwargs)
            else:
                for middle in middles:
                    ctr.position = middle
                    sleep(.2)
                    is_have = hasattr(controller, action_type)
                    if is_have:
                        func = getattr(controller, action_type)
                        func(*args, **kwargs)
        else:
            if middles is not None:
                for middle in middles:
                    ctr.position = middle
                    sleep(.2)
