# -*- coding: utf-8 -*-
"""
카카오톡 오픈채팅방 자동 발송기 (2번컴터용)
구글시트 대기 항목 읽기 -> 카톡 PC에서 채팅방 열기 -> 이미지 전송 -> 문구 전송 -> 시트에 결과 기록
"""
import ctypes
import datetime
import io
import json
import logging
import os
import random
import re
import sys
import time

import gspread
import win32api
import win32clipboard
import win32con
import win32event
import win32gui
import win32process
import winerror
from gspread.utils import rowcol_to_a1
from PIL import Image

KAKAO_MAIN_TITLES = ("카카오톡", "KakaoTalk")
COL_ROOM, COL_TEXT, COL_IMAGE, COL_STATUS, COL_LAST = "방이름", "문구", "이미지경로", "상태", "마지막발송일시"
REQUIRED_COLS = (COL_ROOM, COL_TEXT, COL_IMAGE, COL_STATUS, COL_LAST)
STATUS_DONE = "발송완료"

DEFAULT_CONFIG = {
    "spreadsheet_id": "",
    "worksheet": "2번컴터",
    "image_folder": "C:/Users/USER/Desktop/단지마커",
    "delay_between_rooms_min": 4,
    "delay_between_rooms_max": 9,
    "kakao_wait_sec": 300,
    "test_mode": False,
}


def base_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


BASE = base_dir()
log = logging.getLogger("sender")


# ---------------------------------------------------------------- 공통
def setup_logging():
    log_dir = os.path.join(BASE, "logs")
    os.makedirs(log_dir, exist_ok=True)
    fname = os.path.join(log_dir, datetime.date.today().strftime("%Y-%m-%d") + ".log")
    fmt = logging.Formatter("%(asctime)s  %(message)s", "%H:%M:%S")
    fh = logging.FileHandler(fname, encoding="utf-8")
    fh.setFormatter(fmt)
    log.addHandler(fh)
    if sys.stdout:
        sh = logging.StreamHandler(sys.stdout)
        sh.setFormatter(fmt)
        log.addHandler(sh)
    log.setLevel(logging.INFO)


def load_config():
    cfg = dict(DEFAULT_CONFIG)
    path = os.path.join(BASE, "config.json")
    if not os.path.exists(path):
        raise RuntimeError("config.json 파일이 exe와 같은 폴더에 없습니다")
    with open(path, encoding="utf-8-sig") as f:
        cfg.update(json.load(f))
    if not cfg["spreadsheet_id"]:
        raise RuntimeError("config.json 의 spreadsheet_id 가 비어 있습니다")
    return cfg


def sent_today(value, today):
    """'2026-09-29 10:00:05' 또는 '2026. 9. 29 ...' 형식 모두 오늘 날짜인지 판별"""
    nums = re.findall(r"\d+", value or "")
    if len(nums) < 3:
        return False
    try:
        return (int(nums[0]), int(nums[1]), int(nums[2])) == (today.year, today.month, today.day)
    except ValueError:
        return False


# ---------------------------------------------------------------- 구글시트
def open_sheet(cfg):
    key = os.path.join(BASE, "service_account.json")
    if not os.path.exists(key):
        raise RuntimeError("service_account.json 파일이 exe와 같은 폴더에 없습니다")
    gc = gspread.service_account(filename=key)
    return gc.open_by_key(cfg["spreadsheet_id"]).worksheet(cfg["worksheet"])


def load_targets(ws, today):
    values = ws.get_all_values()
    if not values:
        raise RuntimeError("시트가 비어 있습니다")
    header = [c.strip() for c in values[0]]
    missing = [c for c in REQUIRED_COLS if c not in header]
    if missing:
        raise RuntimeError("시트 1행에 열 이름이 없습니다: " + ", ".join(missing))
    idx = {c: header.index(c) + 1 for c in REQUIRED_COLS}  # 1부터 시작

    targets = []
    for r, row in enumerate(values[1:], start=2):
        def get(col):
            i = idx[col] - 1
            return row[i].strip() if i < len(row) else ""

        room = get(COL_ROOM)
        if not room:
            continue
        if get(COL_STATUS) == STATUS_DONE and sent_today(get(COL_LAST), today):
            continue  # 오늘 이미 보낸 방
        targets.append({"row": r, "room": room, "text": get(COL_TEXT), "image": get(COL_IMAGE)})
    return targets, idx


def write_result(ws, idx, row, status):
    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    for attempt in range(3):
        try:
            ws.batch_update(
                [
                    {"range": rowcol_to_a1(row, idx[COL_STATUS]), "values": [[status]]},
                    {"range": rowcol_to_a1(row, idx[COL_LAST]), "values": [[now]]},
                ],
                value_input_option="RAW",
            )
            return
        except Exception as e:
            log.info(f"  시트 기록 재시도({attempt + 1}/3): {e}")
            time.sleep(3)
    log.info("  ※ 시트 기록 실패 (발송 결과는 로그 파일 참고)")


# ---------------------------------------------------------------- 윈도우 제어
def enum_top_windows():
    result = []

    def cb(h, acc):
        acc.append(h)
        return True

    win32gui.EnumWindows(cb, result)
    return result


def all_descendants(parent):
    result = []

    def cb(h, acc):
        acc.append(h)
        return True

    try:
        win32gui.EnumChildWindows(parent, cb, result)
    except win32gui.error:
        pass
    return result


def direct_children(parent, cls):
    res, h = [], 0
    while True:
        h = win32gui.FindWindowEx(parent, h, cls, None)
        if not h:
            return res
        res.append(h)


def is_kakao(h):
    try:
        return win32gui.GetClassName(h).startswith("EVA_")
    except win32gui.error:
        return False


def find_main_window():
    for h in enum_top_windows():
        if win32gui.GetWindowText(h) in KAKAO_MAIN_TITLES and is_kakao(h):
            return h
    return 0


def find_room_window(room):
    for h in enum_top_windows():
        if win32gui.GetWindowText(h) == room and is_kakao(h) and win32gui.IsWindowVisible(h):
            return h
    return 0


def set_text(h, s):
    win32api.SendMessage(h, win32con.WM_SETTEXT, 0, s)


def text_len(h):
    return win32gui.SendMessage(h, win32con.WM_GETTEXTLENGTH, 0, 0)


def get_text(h):
    n = text_len(h)
    buf = ctypes.create_unicode_buffer(n + 2)
    ctypes.windll.user32.SendMessageW(h, win32con.WM_GETTEXT, n + 2, buf)
    return buf.value


def post_key(h, vk):
    win32api.PostMessage(h, win32con.WM_KEYDOWN, vk, 0)
    time.sleep(0.05)
    win32api.PostMessage(h, win32con.WM_KEYUP, vk, 0)


def real_key(vk, ctrl=False):
    if ctrl:
        win32api.keybd_event(win32con.VK_CONTROL, 0, 0, 0)
    win32api.keybd_event(vk, 0, 0, 0)
    time.sleep(0.05)
    win32api.keybd_event(vk, 0, win32con.KEYEVENTF_KEYUP, 0)
    if ctrl:
        win32api.keybd_event(win32con.VK_CONTROL, 0, win32con.KEYEVENTF_KEYUP, 0)


def force_foreground(h):
    if win32gui.IsIconic(h):
        win32gui.ShowWindow(h, win32con.SW_RESTORE)
    # 다른 창이 포커스를 쥐고 있어도 앞으로 가져오기 위한 처리
    win32api.keybd_event(win32con.VK_MENU, 0, 0, 0)
    win32api.keybd_event(win32con.VK_MENU, 0, win32con.KEYEVENTF_KEYUP, 0)
    try:
        win32gui.SetForegroundWindow(h)
    except win32gui.error:
        pass
    time.sleep(0.5)
    return win32gui.GetForegroundWindow() == h


def wait_for_kakao(timeout):
    end = time.time() + timeout
    while time.time() < end:
        if find_main_window():
            return True
        log.info("카카오톡 실행 대기 중...")
        time.sleep(10)
    return False


# ---------------------------------------------------------------- 카카오톡 동작
def find_search_edits(main):
    edits = []
    for child in direct_children(main, "EVA_ChildWindow"):
        panes = direct_children(child, "EVA_Window")
        # 보통 [0]=친구 목록, [1]=채팅 목록 -> 채팅 목록부터 시도
        for pane in panes[1:] + panes[:1]:
            edits += direct_children(pane, "Edit")
    if not edits:  # 카톡 버전이 바뀌어 구조가 다를 때 대비
        edits = [h for h in all_descendants(main) if win32gui.GetClassName(h) == "Edit"]
    return edits


def open_room(room):
    h = find_room_window(room)
    if h:
        return h
    main = find_main_window()
    if not main:
        raise RuntimeError("카카오톡 메인창 없음(실행/로그인 확인)")
    edits = find_search_edits(main)
    if not edits:
        raise RuntimeError("카톡 검색창을 찾지 못함(로그인 상태 확인)")
    for e in edits:
        set_text(e, room)
        time.sleep(1.2)
        post_key(e, win32con.VK_RETURN)
        for _ in range(24):
            time.sleep(0.25)
            h = find_room_window(room)
            if h:
                set_text(e, "")
                time.sleep(1.0)
                return h
        set_text(e, "")
    raise RuntimeError("채팅방을 찾지 못함(방이름 확인)")


def close_room(h):
    try:
        win32gui.PostMessage(h, win32con.WM_CLOSE, 0, 0)
    except win32gui.error:
        pass
    time.sleep(0.8)


def find_input(room_hwnd):
    for h in all_descendants(room_hwnd):
        if win32gui.GetClassName(h).startswith("RICHEDIT"):
            return h
    raise RuntimeError("채팅 입력창을 찾지 못함")


def resolve_image(cfg, value):
    path = value.strip().strip('"')
    if not os.path.isabs(path):
        path = os.path.join(cfg["image_folder"], path)
    if not os.path.isfile(path):
        raise RuntimeError(f"이미지 파일 없음: {path}")
    return path


def set_clipboard_image(path):
    img = Image.open(path)
    if img.mode in ("RGBA", "LA", "P"):
        img = img.convert("RGBA")
        bg = Image.new("RGB", img.size, (255, 255, 255))
        bg.paste(img, mask=img.split()[-1])
        img = bg
    else:
        img = img.convert("RGB")
    buf = io.BytesIO()
    img.save(buf, "BMP")
    dib = buf.getvalue()[14:]  # BMP 파일 헤더 제거
    for _ in range(10):
        try:
            win32clipboard.OpenClipboard()
            break
        except Exception:
            time.sleep(0.3)
    else:
        raise RuntimeError("클립보드를 사용할 수 없음")
    try:
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(win32clipboard.CF_DIB, dib)
    finally:
        win32clipboard.CloseClipboard()


def window_pid(h):
    try:
        return win32process.GetWindowThreadProcessId(h)[1]
    except Exception:
        return 0


def visible_windows_of(pid):
    return {h for h in enum_top_windows() if win32gui.IsWindowVisible(h) and window_pid(h) == pid}


def click_center(h):
    l, t, r, b = win32gui.GetWindowRect(h)
    x, y = (l + r) // 2, (t + b) // 2
    win32api.SetCursorPos((x, y))
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    time.sleep(0.05)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)


def describe(h):
    try:
        return f"{win32gui.GetClassName(h)}/'{win32gui.GetWindowText(h)}'"
    except Exception:
        return str(h)


def send_image(room_hwnd, path):
    set_clipboard_image(path)
    edit = find_input(room_hwnd)
    set_text(edit, "")
    if not force_foreground(room_hwnd):
        raise RuntimeError("채팅방 창을 앞으로 가져오지 못함(화면 잠금 확인)")
    click_center(edit)  # 입력창에 커서 두기
    time.sleep(0.4)

    pid = window_pid(room_hwnd)
    before = visible_windows_of(pid)
    real_key(ord("V"), ctrl=True)

    popup = 0
    for _ in range(24):  # 이미지 전송 확인창 대기 (최대 6초)
        time.sleep(0.25)
        new = visible_windows_of(pid) - before
        if new:
            popup = next(iter(new))
            break
        fg = win32gui.GetForegroundWindow()
        if fg and fg != room_hwnd and window_pid(fg) == pid:
            popup = fg
            break
        if not win32gui.IsWindowEnabled(room_hwnd):  # 모달창이 떠서 채팅방이 잠김
            popup = -1
            break

    if not popup:
        fg = win32gui.GetForegroundWindow()
        log.info(f"  [진단] 전경창={describe(fg)}, 입력창 글자수={text_len(edit)}")
        raise RuntimeError("이미지 전송창이 뜨지 않음")

    if popup > 0:
        log.info(f"  [진단] 전송창={describe(popup)}")
        force_foreground(popup)
    time.sleep(0.6)
    real_key(win32con.VK_RETURN)

    for _ in range(40):
        time.sleep(0.25)
        closed = (popup > 0 and (not win32gui.IsWindow(popup) or not win32gui.IsWindowVisible(popup))) or \
                 (popup == -1 and win32gui.IsWindowEnabled(room_hwnd))
        if closed:
            time.sleep(1.5)  # 업로드 여유
            return
    raise RuntimeError("이미지 전송창이 닫히지 않음")


def set_clipboard_text(text):
    for _ in range(10):
        try:
            win32clipboard.OpenClipboard()
            break
        except Exception:
            time.sleep(0.3)
    else:
        raise RuntimeError("클립보드를 사용할 수 없음")
    try:
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(win32clipboard.CF_UNICODETEXT, text)
    finally:
        win32clipboard.CloseClipboard()


def send_text(room_hwnd, text):
    edit = find_input(room_hwnd)
    text = text.replace("\r\n", "\n").replace("\n", "\r\n")
    # 입력창에 안내문구(예: '메시지 입력')가 잡힐 수 있어, 우리 문구 앞부분이 남아있는지로 판단
    sig = re.sub(r"\s+", "", text)[:12]

    def still_there():
        return sig and sig in re.sub(r"\s+", "", get_text(edit))

    set_text(edit, "")
    time.sleep(0.3)

    if force_foreground(room_hwnd):
        # 실제 사람처럼: 입력창 클릭 -> 붙여넣기 -> 엔터
        click_center(edit)
        time.sleep(0.3)
        set_clipboard_text(text)
        real_key(ord("V"), ctrl=True)
        time.sleep(0.8)
        if not still_there():  # 붙여넣기가 안 됐으면 직접 입력
            set_text(edit, text)
            time.sleep(0.5)
        real_key(win32con.VK_RETURN)
    else:
        set_text(edit, text)
        time.sleep(0.6)
        post_key(edit, win32con.VK_RETURN)

    for _ in range(20):
        time.sleep(0.25)
        if not still_there():
            return
    log.info(f"  [진단] 전경창={describe(win32gui.GetForegroundWindow())}, 입력창='{get_text(edit)[:20]}'")
    raise RuntimeError("문구 전송 확인 실패")


# ---------------------------------------------------------------- 메인
def run():
    cfg = load_config()
    log.info("=== 발송 시작 ===")
    ws = open_sheet(cfg)
    targets, idx = load_targets(ws, datetime.date.today())
    if cfg.get("test_mode"):
        targets = targets[:1]
        log.info("[테스트 모드] 첫 번째 방 1개만 발송합니다")
    if not targets:
        log.info("오늘 보낼 대기 항목이 없습니다")
        return

    if not wait_for_kakao(int(cfg["kakao_wait_sec"])):
        log.info("카카오톡이 실행되어 있지 않아 발송을 중단합니다")
        for t in targets:
            write_result(ws, idx, t["row"], "실패: 카카오톡 미실행")
        return

    ok = fail = 0
    for i, t in enumerate(targets, 1):
        log.info(f"[{i}/{len(targets)}] {t['room']} 발송 중")
        room_hwnd = 0
        try:
            if not t["image"] and not t["text"]:
                raise RuntimeError("문구/이미지가 모두 비어 있음")
            img_path = resolve_image(cfg, t["image"]) if t["image"] else None
            room_hwnd = open_room(t["room"])
            if img_path:
                send_image(room_hwnd, img_path)
            if t["text"]:
                send_text(room_hwnd, t["text"])
            status = STATUS_DONE
            ok += 1
            log.info("  -> 발송완료")
        except Exception as e:
            status = f"실패: {e}"
            fail += 1
            log.info(f"  -> {status}")
        finally:
            if room_hwnd:
                close_room(room_hwnd)
        write_result(ws, idx, t["row"], status)
        if i < len(targets):
            time.sleep(random.uniform(float(cfg["delay_between_rooms_min"]), float(cfg["delay_between_rooms_max"])))

    log.info(f"=== 완료: 성공 {ok} / 실패 {fail} ===")


def main():
    setup_logging()
    mutex = win32event.CreateMutex(None, False, "Local\\KakaoOpenChatSender")  # noqa: F841 중복 실행 방지
    if win32api.GetLastError() == winerror.ERROR_ALREADY_EXISTS:
        log.info("이미 실행 중이라 종료합니다")
        return
    try:
        run()
    except Exception as e:
        log.exception(f"오류로 중단: {e}")
    time.sleep(5)


if __name__ == "__main__":
    main()
