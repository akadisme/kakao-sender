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
import ntpath
import os
import random
import re
import subprocess
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
from PIL import Image, ImageOps

VERSION = "2026-10-30"

KAKAO_MAIN_TITLES = ("카카오톡", "KakaoTalk")
COL_ROOM, COL_TEXT, COL_IMAGE, COL_STATUS, COL_LAST = "방이름", "문구", "이미지경로", "상태", "마지막발송일시"
REQUIRED_COLS = (COL_ROOM, COL_TEXT, COL_IMAGE, COL_STATUS, COL_LAST)
COL_TIME = "발송시간"  # 선택 열: 11:00 / 16:00 / 09:00~12:00(이 시간 안에만) / 비우면 기본시간
STATUS_DONE = "발송완료"

DEFAULT_CONFIG = {
    "spreadsheet_id": "",
    "worksheet": "2번컴터",
    "image_folder": "C:/Users/USER/Desktop/단지마커",
    "image_folders": ["D:/단지마커", "D:/Desktop/단지마커", "D:/Users/USER/Desktop/단지마커"],
    "image_max_px": 1280,          # 이미지 가로/세로 최대 크기(px). 큰 사진은 자동으로 줄여서 보냄. 0이면 원본
    "upload_wait_sec": 4,          # 이미지 보낸 뒤 문구 보내기 전 대기
    "upload_timeout_sec": 180,     # 업로드가 끝날 때까지 채팅방을 닫지 않고 기다리는 최대 시간
    "delay_between_rooms_min": 4,
    "delay_between_rooms_max": 9,
    "kakao_wait_min": 60,          # 카카오톡 실행/로그인을 기다리는 최대 시간(분)
    "kakao_path": "",              # 비워두면 C:, D: 기본 설치 위치에서 자동으로 찾음
    "retry_rounds": 2,             # 실패한 방 재시도 횟수
    "retry_wait_sec": 120,         # 재시도 전 대기
    "auto_admin": True,            # 더블클릭 실행 시 관리자 권한으로 자동 재실행
    "default_send_time": "11:00",  # 시트 발송시간 칸이 비어 있을 때 쓰는 시간
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


def as_bool(v):
    if isinstance(v, str):
        return v.strip().lower() in ("true", "1", "yes", "y", "on", "예")
    return bool(v)


def as_num(v, default):
    try:
        return float(v)
    except (TypeError, ValueError):
        return default


def parse_config_text(txt):
    """메모장으로 고치다 생기는 흔한 실수(True 대문자, D:\\단지마커 같은 역슬래시, 마지막 쉼표)는 자동 보정"""
    try:
        return json.loads(txt)
    except json.JSONDecodeError:
        pass
    fixed = txt.replace("\u201c", '"').replace("\u201d", '"').replace("\\", "/")
    fixed = re.sub(r"\bTrue\b", "true", fixed)
    fixed = re.sub(r"\bFalse\b", "false", fixed)
    fixed = re.sub(r",\s*([}\]])", r"\1", fixed)
    try:
        return json.loads(fixed)
    except json.JSONDecodeError as e:
        line = txt.splitlines()[e.lineno - 1].strip() if 0 < e.lineno <= len(txt.splitlines()) else ""
        raise RuntimeError(f"config.json {e.lineno}번째 줄 형식 오류: {line}  (true/false는 소문자, 따옴표 확인)")


def load_config():
    cfg = dict(DEFAULT_CONFIG)
    path = os.path.join(BASE, "config.json")
    if not os.path.exists(path):
        raise RuntimeError("config.json 파일이 exe와 같은 폴더에 없습니다")
    with open(path, encoding="utf-8-sig") as f:
        cfg.update(parse_config_text(f.read()))
    if not cfg["spreadsheet_id"]:
        raise RuntimeError("config.json 의 spreadsheet_id 가 비어 있습니다")
    if isinstance(cfg.get("image_folders"), str):
        cfg["image_folders"] = [cfg["image_folders"]]
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


# ---------------------------------------------------------------- 발송시간
def parse_clock(text):
    """'16:00', '16시', '오후 4:00:00', '4:00 PM', '1600' -> 하루 중 분(0~1439). 못 읽으면 None"""
    t = str(text or "").strip().lower().replace(" ", "")
    if not t:
        return None
    pm = any(k in t for k in ("오후", "pm", "p.m"))
    am = any(k in t for k in ("오전", "am", "a.m"))
    nums = re.findall(r"\d+", t)
    if not nums:
        return None
    if len(nums) == 1 and len(nums[0]) in (3, 4) and "시" not in t:  # 1600
        h, m = int(nums[0][:-2]), int(nums[0][-2:])
    else:
        h = int(nums[0])
        m = int(nums[1]) if len(nums) > 1 else 0
    if pm and h < 12:
        h += 12
    if am and h == 12:
        h = 0
    if not (0 <= h <= 24 and 0 <= m < 60):
        return None
    return min(h * 60 + m, 24 * 60)


def parse_window(text, default):
    """'16:00' -> (960, None) / '09:00~12:00' -> (540, 720) / 빈칸 -> 기본시간"""
    raw = str(text or "").strip() or str(default or "11:00")
    parts = re.split(r"\s*[~〜]\s*|\s+-\s+|(?<=\d)-(?=\s*(?:\d|오|am|pm))", raw, maxsplit=1)
    start = parse_clock(parts[0])
    end = parse_clock(parts[1]) if len(parts) > 1 else None
    if start is None:
        raise ValueError(f"발송시간을 읽을 수 없음: {raw}")
    return start, end


def fmt_min(m):
    return f"{m // 60:02d}:{m % 60:02d}"


def is_due(window, now_min):
    start, end = window
    return now_min >= start and (end is None or now_min < end)


# ---------------------------------------------------------------- 구글시트
def open_sheet(cfg):
    key = os.path.join(BASE, "service_account.json")
    if not os.path.exists(key):
        raise RuntimeError("service_account.json 파일이 exe와 같은 폴더에 없습니다")
    gc = gspread.service_account(filename=key)
    try:
        gc.set_timeout(60)  # 인터넷이 끊겨도 무한정 멈추지 않게
    except Exception:
        pass
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
    if COL_TIME in header:
        idx[COL_TIME] = header.index(COL_TIME) + 1

    targets = []
    for r, row in enumerate(values[1:], start=2):
        def get(col):
            i = idx[col] - 1
            return row[i].strip() if i < len(row) else ""

        room = get(COL_ROOM)
        if not room:
            continue
        if get(COL_STATUS).startswith(STATUS_DONE) and sent_today(get(COL_LAST), today):
            continue  # 오늘 이미 보낸 방
        targets.append({"row": r, "room": room, "text": get(COL_TEXT), "image": get(COL_IMAGE),
                        "time": get(COL_TIME) if COL_TIME in idx else ""})
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
    for _ in range(3):
        try:
            win32gui.BringWindowToTop(h)
            win32gui.SetForegroundWindow(h)
        except win32gui.error:
            pass
        time.sleep(0.5)
        if win32gui.GetForegroundWindow() == h:
            return True
    return False


def find_kakao_exe(cfg):
    cands = [str(cfg.get("kakao_path") or "").replace("/", "\\")]
    for drive in ("C:", "D:", "E:"):
        for sub in ("Program Files (x86)", "Program Files"):
            cands.append(ntpath.join(drive + "\\", sub, "Kakao", "KakaoTalk", "KakaoTalk.exe"))
    local = os.environ.get("LOCALAPPDATA")
    if local:
        cands.append(ntpath.join(local, "Kakao", "KakaoTalk", "KakaoTalk.exe"))
    for c in cands:
        if c and os.path.isfile(c):
            return c
    return ""


def launch_kakao(exe):
    try:
        subprocess.Popen([exe], cwd=os.path.dirname(exe))
        return True
    except Exception as e:
        log.info(f"  [진단] 카카오톡 실행 실패: {e}")
        return False


def ensure_kakao(cfg):
    """카톡이 꺼져 있거나 트레이(X로 닫음)에만 있으면 직접 실행해서 메인창을 띄움"""
    exe = find_kakao_exe(cfg)
    end = time.time() + as_num(cfg.get("kakao_wait_min"), 60) * 60
    last_launch = last_log = 0
    while True:
        main = find_main_window()
        if main and win32gui.IsWindowVisible(main) and find_search_edits(main):
            return True
        if time.time() >= end:
            return False
        now = time.time()
        if exe and now - last_launch > 180:
            log.info("카카오톡 메인창이 없어 카카오톡을 실행합니다" if not main else "카카오톡이 트레이에 있어 메인창을 띄웁니다")
            launch_kakao(exe)
            last_launch = now
            time.sleep(8)
            main = find_main_window()
            if main and win32gui.IsWindowVisible(main):
                time.sleep(12)  # 자동 로그인 여유
                try:
                    win32gui.ShowWindow(main, win32con.SW_MINIMIZE)
                except win32gui.error:
                    pass
            continue
        if main and not win32gui.IsWindowVisible(main):
            try:
                win32gui.ShowWindow(main, win32con.SW_SHOWMINNOACTIVE)
            except win32gui.error:
                pass
        if now - last_log > 60:
            log.info("카카오톡 실행/로그인 대기 중..." + ("" if exe else " (카카오톡 설치 위치를 못 찾음 - config의 kakao_path 확인)"))
            last_log = now
        time.sleep(10)


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


def close_room_safely(h, timeout):
    """업로드가 끝나기 전에 닫으면 '전송 중인 파일' 창이 뜨고 사진이 안 가므로,
    그런 확인창이 뜨면 취소하고 기다렸다가 다시 닫음. 끝까지 안 되면 방을 열어둠."""
    pid = window_pid(h)
    end = time.time() + timeout
    waited = False
    while True:
        if not win32gui.IsWindow(h) or not win32gui.IsWindowVisible(h):
            return True
        before = visible_windows_of(pid)
        try:
            win32gui.PostMessage(h, win32con.WM_CLOSE, 0, 0)
        except win32gui.error:
            return True
        for _ in range(8):
            time.sleep(0.25)
            if not win32gui.IsWindow(h) or not win32gui.IsWindowVisible(h):
                break
        new = visible_windows_of(pid) - before - {h}
        if not win32gui.IsWindow(h) or not win32gui.IsWindowVisible(h):
            for d in new:
                log.info(f"  [진단] 채팅방 닫은 뒤 새 창: {describe(d)}")
            return True
        # 아직 열려 있음 = '전송 중' 확인창이 떴을 가능성 -> 취소(ESC)하고 업로드 기다림
        for d in new:
            if not waited:
                log.info(f"  [안내] 이미지 업로드 중이라 기다립니다 ({describe(d)})")
            post_key(d, win32con.VK_ESCAPE)
        waited = True
        if time.time() >= end:
            log.info("  [안내] 업로드가 오래 걸려 채팅방을 열어둡니다(나중에 다시 닫기)")
            return False
        time.sleep(5)


def find_input(room_hwnd):
    for h in all_descendants(room_hwnd):
        if win32gui.GetClassName(h).startswith("RICHEDIT"):
            return h
    raise RuntimeError("채팅 입력창을 찾지 못함")


def image_candidates(cfg, value):
    raw = value.strip().strip('"').strip().replace("/", "\\")
    drive, rest = ntpath.splitdrive(raw)
    cands = []
    if drive:
        cands.append(raw)
        for d in ("C:", "D:", "E:"):  # C:에서 D:로 옮겼거나 그 반대인 경우
            if d.upper() != drive.upper():
                cands.append(d + rest)
        name = ntpath.basename(raw)
    else:
        name = raw
    folders = [cfg.get("image_folder")] + list(cfg.get("image_folders") or [])
    folders += [ntpath.join(os.path.expanduser("~"), "Desktop", "단지마커"), ntpath.join(BASE, "images"), BASE]
    for f in folders:
        if f:
            cands.append(ntpath.join(str(f).replace("/", "\\"), name))
    seen, out = set(), []
    for c in cands:
        if c.lower() not in seen:
            seen.add(c.lower())
            out.append(c)
    return out


def resolve_image(cfg, value):
    cands = image_candidates(cfg, value)
    for i, c in enumerate(cands):
        if os.path.isfile(c):
            if i > 0:
                log.info(f"  [안내] 시트 경로에 이미지가 없어 대신 사용: {c}")
            return c
    raise RuntimeError(f"이미지 파일 없음: {cands[0]}")


def open_clipboard():
    for _ in range(60):  # 다른 프로그램이 잠깐 쥐고 있어도 최대 15초 기다림
        try:
            win32clipboard.OpenClipboard()
            return
        except Exception:
            time.sleep(0.25)
    raise RuntimeError("클립보드를 사용할 수 없음")


_DIB_CACHE = {}


def image_to_dib(path, max_px):
    key = (path, os.path.getmtime(path), max_px)
    if key in _DIB_CACHE:
        return _DIB_CACHE[key]
    img = Image.open(path)
    try:
        img = ImageOps.exif_transpose(img)
    except Exception:
        pass
    if img.mode in ("RGBA", "LA", "P"):
        img = img.convert("RGBA")
        bg = Image.new("RGB", img.size, (255, 255, 255))
        bg.paste(img, mask=img.split()[-1])
        img = bg
    else:
        img = img.convert("RGB")
    if max_px and max(img.size) > max_px:
        before = img.size
        img.thumbnail((max_px, max_px), Image.LANCZOS)
        log.info(f"  [안내] 이미지가 커서 줄여서 보냄: {before[0]}x{before[1]} -> {img.size[0]}x{img.size[1]}")
    buf = io.BytesIO()
    img.save(buf, "BMP")
    dib = buf.getvalue()[14:]  # BMP 파일 헤더 제거
    _DIB_CACHE.clear()
    _DIB_CACHE[key] = dib
    return dib


def set_clipboard_image(path, max_px=1280):
    dib = image_to_dib(path, int(max_px or 0))
    open_clipboard()
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
    try:
        win32api.SetCursorPos((x, y))
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
        time.sleep(0.05)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
        return True
    except Exception as e:
        log.info(f"  [진단] 커서 이동 실패 -> 메시지 클릭으로 대체 ({e})")
    # 마우스 커서를 못 움직이는 상황(화면 꺼짐 등) -> 창에 직접 클릭 메시지 전송
    lp = win32api.MAKELONG((r - l) // 2, (b - t) // 2)
    win32gui.PostMessage(h, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lp)
    time.sleep(0.05)
    win32gui.PostMessage(h, win32con.WM_LBUTTONUP, 0, lp)
    return False


def focus_input(room_hwnd, edit):
    """마우스 없이 입력창에 키보드 포커스를 직접 지정"""
    cur = win32api.GetCurrentThreadId()
    target = win32process.GetWindowThreadProcessId(room_hwnd)[0]
    attached = False
    try:
        attached = bool(ctypes.windll.user32.AttachThreadInput(cur, target, True))
        ctypes.windll.user32.SetFocus(edit)
    except Exception as e:
        log.info(f"  [진단] 포커스 지정 실패 ({e})")
    finally:
        if attached:
            ctypes.windll.user32.AttachThreadInput(cur, target, False)
    time.sleep(0.2)


ES_CONTINUOUS, ES_SYSTEM_REQUIRED, ES_DISPLAY_REQUIRED = 0x80000000, 0x00000001, 0x00000002


def wake_display():
    """꺼진 화면을 켜고, 발송하는 동안 화면이 꺼지지 않게 유지"""
    try:
        ctypes.windll.kernel32.SetThreadExecutionState(ES_CONTINUOUS | ES_SYSTEM_REQUIRED | ES_DISPLAY_REQUIRED)
        win32api.mouse_event(win32con.MOUSEEVENTF_MOVE, 1, 0, 0, 0)
        time.sleep(0.05)
        win32api.mouse_event(win32con.MOUSEEVENTF_MOVE, -1, 0, 0, 0)
    except Exception:
        pass
    time.sleep(3)


def describe(h):
    try:
        return f"{win32gui.GetClassName(h)}/'{win32gui.GetWindowText(h)}'"
    except Exception:
        return str(h)


def wait_popup(room_hwnd, pid, before, seconds):
    for _ in range(int(seconds * 4)):
        time.sleep(0.25)
        new = visible_windows_of(pid) - before
        if new:
            return next(iter(new))
        fg = win32gui.GetForegroundWindow()
        if fg and fg != room_hwnd and window_pid(fg) == pid:
            return fg
        if not win32gui.IsWindowEnabled(room_hwnd):  # 모달창이 떠서 채팅방이 잠김
            return -1
    return 0


def send_image(room_hwnd, path, cfg):
    max_px = int(as_num(cfg.get("image_max_px"), 1280))
    set_clipboard_image(path, max_px)
    edit = find_input(room_hwnd)
    set_text(edit, "")
    if not force_foreground(room_hwnd):
        raise RuntimeError("채팅방 창을 앞으로 가져오지 못함(화면 잠금 확인)")
    click_center(edit)
    focus_input(room_hwnd, edit)
    time.sleep(0.3)

    pid = window_pid(room_hwnd)
    before = visible_windows_of(pid)

    # 방법1: Ctrl+V 키 입력
    real_key(ord("V"), ctrl=True)
    popup = wait_popup(room_hwnd, pid, before, 3)

    # 방법2: 입력창에 붙여넣기 명령 직접 전송
    if not popup:
        log.info("  [진단] Ctrl+V 반응 없음 -> 붙여넣기 명령으로 재시도")
        set_clipboard_image(path, max_px)
        win32gui.SendMessage(edit, win32con.WM_PASTE, 0, 0)
        popup = wait_popup(room_hwnd, pid, before, 3)

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
            time.sleep(as_num(cfg.get("upload_wait_sec"), 4))  # 업로드 여유
            return
    raise RuntimeError("이미지 전송창이 닫히지 않음")


def set_clipboard_text(text):
    open_clipboard()
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
        focus_input(room_hwnd, edit)
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
def send_one(cfg, t, leftover):
    room_hwnd = 0
    try:
        if not t["image"] and not t["text"]:
            raise RuntimeError("문구/이미지가 모두 비어 있음")
        img_path = resolve_image(cfg, t["image"]) if t["image"] else None
        room_hwnd = open_room(t["room"])
        if img_path:
            send_image(room_hwnd, img_path, cfg)
        if t["text"]:
            send_text(room_hwnd, t["text"])
        status = STATUS_DONE
        time.sleep(2)
        if not close_room_safely(room_hwnd, as_num(cfg.get("upload_timeout_sec"), 180)):
            leftover.append(room_hwnd)
            status = STATUS_DONE + "(업로드 확인 필요)"
        room_hwnd = 0
        return True, status
    except Exception as e:
        return False, f"실패: {e}"
    finally:
        if room_hwnd:
            close_room_safely(room_hwnd, 30)


def pick_due(cfg, targets, now_min):
    """지금 보낼 방 / 아직 시간이 안 된 방 / 시간이 지나 오늘은 건너뛰는 방"""
    due, waiting, missed = [], [], []
    for t in targets:
        try:
            w = parse_window(t["time"], cfg.get("default_send_time"))
        except ValueError as e:
            log.info(f"[주의] {t['room']}: {e} -> 기본시간 {cfg.get('default_send_time')} 사용")
            w = parse_window("", cfg.get("default_send_time"))
        t["window"] = w
        if is_due(w, now_min):
            due.append(t)
        elif now_min < w[0]:
            waiting.append(t)
        else:
            missed.append(t)
    return due, waiting, missed


def run():
    """반환값: 사람이 직접 실행했을 때 보여줄 결과 문장"""
    cfg = load_config()
    ws = open_sheet(cfg)
    now = datetime.datetime.now()
    targets, idx = load_targets(ws, now.date())
    now_min = now.hour * 60 + now.minute
    targets, waiting, missed = pick_due(cfg, targets, now_min)
    if as_bool(cfg.get("test_mode")):
        targets = targets[:1]
    if not targets:
        # 30분마다 실행되므로 보낼 게 없을 땐 로그를 남기지 않음
        nxt = sorted({fmt_min(t["window"][0]) for t in waiting})
        return "지금 보낼 방이 없어요." + (f"\n오늘 남은 발송시간: {', '.join(nxt)}" if nxt else "")

    log.info(f"=== 발송 시작 (버전 {VERSION}, 관리자 권한: {'예' if is_admin() else '아니오'}) ===")
    if as_bool(cfg.get("test_mode")):
        log.info("[테스트 모드] 지금 보낼 방 중 첫 번째 1개만 발송합니다")
    log.info(f"지금 보낼 방 {len(targets)}개: " + ", ".join(f"{t['room']}({fmt_min(t['window'][0])})" for t in targets))
    for t in missed:
        log.info(f"[안내] {t['room']}: 발송 가능 시간({fmt_min(t['window'][0])}~{fmt_min(t['window'][1])})이 지나 오늘은 건너뜀")

    wake_display()
    if not ensure_kakao(cfg):
        log.info("카카오톡이 실행되어 있지 않아 발송을 중단합니다")
        for t in targets:
            write_result(ws, idx, t["row"], "실패: 카카오톡 미실행")
        return "카카오톡이 실행되어 있지 않아 발송을 못 했어요."

    dmin = as_num(cfg.get("delay_between_rooms_min"), 4)
    dmax = max(dmin, as_num(cfg.get("delay_between_rooms_max"), 9))
    rounds = 1 + max(0, int(as_num(cfg.get("retry_rounds"), 2)))
    pending, ok, leftover = targets, 0, []
    for rnd in range(1, rounds + 1):
        if rnd > 1:
            wait = as_num(cfg.get("retry_wait_sec"), 120)
            log.info(f"--- 실패한 {len(pending)}개 방 재시도 ({rnd - 1}/{rounds - 1}), {int(wait)}초 후 시작 ---")
            time.sleep(wait)
            wake_display()
            if not ensure_kakao(cfg):
                log.info("카카오톡이 없어 재시도를 중단합니다")
                break
        failed = []
        for i, t in enumerate(pending, 1):
            log.info(f"[{i}/{len(pending)}] {t['room']} 발송 중")
            success, status = send_one(cfg, t, leftover)
            log.info(f"  -> {status}")
            write_result(ws, idx, t["row"], status)
            if success:
                ok += 1
            else:
                failed.append(t)
            if i < len(pending):
                time.sleep(random.uniform(dmin, dmax))
        pending = failed
        if not pending:
            break

    for h in leftover:  # 업로드 때문에 열어둔 방은 마지막에 다시 닫기
        close_room_safely(h, 60)
    log.info(f"=== 완료: 성공 {ok} / 실패 {len(pending)} ===")
    return f"발송 완료: 성공 {ok} / 실패 {len(pending)}" + ("\n실패한 방은 시트 상태 칸을 확인하세요." if pending else "")


def is_admin():
    try:
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        return False


def relaunch_as_admin():
    if getattr(sys, "frozen", False):
        exe, params = sys.executable, subprocess.list2cmdline(sys.argv[1:])
    else:
        exe, params = sys.executable, subprocess.list2cmdline([os.path.abspath(__file__)] + sys.argv[1:])
    try:
        return ctypes.windll.shell32.ShellExecuteW(None, "runas", exe, params, BASE, 1) > 32
    except Exception:
        return False


def admin_wanted():
    try:
        return as_bool(load_config().get("auto_admin", True))
    except Exception:
        return True


def main():
    setup_logging()
    auto = "--auto" in sys.argv  # 작업 스케줄러 실행
    if not is_admin():
        if not auto and admin_wanted() and relaunch_as_admin():
            return
        log.info("[주의] 관리자 권한이 아닙니다. 카톡이 관리자 권한이면 입력이 막힐 수 있어요")
    mutex = win32event.CreateMutex(None, False, "Local\\KakaoOpenChatSender")  # noqa: F841 중복 실행 방지
    if win32api.GetLastError() == winerror.ERROR_ALREADY_EXISTS:
        log.info("이미 실행 중이라 종료합니다")
        return
    try:
        msg = run()
    except Exception as e:
        log.exception(f"오류로 중단: {e}")
        msg = f"오류로 중단: {e}\n(logs 폴더의 오늘 날짜 파일 확인)"
    if not auto and msg:  # 더블클릭으로 실행했을 때만 결과 창 표시
        try:
            ctypes.windll.user32.MessageBoxW(0, msg, "카톡 자동발송", 0x40 | 0x40000)
        except Exception:
            pass


if __name__ == "__main__":
    main()
