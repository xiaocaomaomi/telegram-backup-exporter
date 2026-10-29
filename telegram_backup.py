# -*- coding: utf-8 -*-
"""
电报备份 —— 一键导出你关注的所有群组、频道、机器人和联系人。

设计目标：给完全不懂技术的人用。
  * 只填一个手机号
  * 一个按钮
  * 表格自动保存到桌面并自动打开
  * 不用装 Python，不用配代理，不用管 api_id

项目主页：https://github.com/xiaocaomaomi/telegram-backup-exporter
许可：MIT
"""

import asyncio
import json
import os
import queue
import re
import socket
import subprocess
import sys
import threading
import time
import traceback
from datetime import datetime
from pathlib import Path

import tkinter as tk
from tkinter import messagebox, ttk
from tkinter.scrolledtext import ScrolledText

try:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
    from telethon import TelegramClient, errors
    from telethon.tl.types import Channel, Chat, User
    try:
        from telethon.tl.functions.messages import ExportChatInviteRequest
        _INVITE_IN_MESSAGES = True
    except ImportError:
        from telethon.tl.functions.channels import ExportChatInviteRequest
        _INVITE_IN_MESSAGES = False
except Exception as exc:  # 打包缺失时给一句人话
    print("缺少组件：" + str(exc))
    raise

APP_NAME = "电报备份"
VERSION = "1.0.0"

# 默认程序身份（公开的 Telegram Desktop 常量）。人多时可能被限流，
# 想更稳可以在「高级」里换成自己申请的。详见 README。
DEFAULT_API_ID = 2040
DEFAULT_API_HASH = "b18441a1ff607e10a989891a5462e627"

REGIONS = [
    ("中国 +86", "+86"), ("英国 +44", "+44"), ("香港 +852", "+852"),
    ("澳门 +853", "+853"), ("台湾 +886", "+886"), ("美国/加拿大 +1", "+1"),
    ("日本 +81", "+81"), ("韩国 +82", "+82"), ("新加坡 +65", "+65"),
    ("马来西亚 +60", "+60"), ("泰国 +66", "+66"), ("越南 +84", "+84"),
    ("印尼 +62", "+62"), ("菲律宾 +63", "+63"), ("印度 +91", "+91"),
    ("巴基斯坦 +92", "+92"), ("阿联酋 +971", "+971"), ("澳大利亚 +61", "+61"),
    ("新西兰 +64", "+64"), ("德国 +49", "+49"), ("法国 +33", "+33"),
    ("意大利 +39", "+39"), ("西班牙 +34", "+34"), ("荷兰 +31", "+31"),
    ("瑞士 +41", "+41"), ("瑞典 +46", "+46"), ("波兰 +48", "+48"),
    ("俄罗斯 +7", "+7"), ("乌克兰 +380", "+380"), ("土耳其 +90", "+90"),
    ("巴西 +55", "+55"), ("墨西哥 +52", "+52"), ("南非 +27", "+27"),
    ("埃及 +20", "+20"),
]
REGION_MAP = dict(REGIONS)

PROXY_PORTS = [
    (7890, "socks5"), (7891, "http"), (7897, "http"), (7898, "socks5"),
    (1080, "socks5"), (1081, "http"), (10808, "socks5"), (10809, "http"),
    (2080, "socks5"), (8889, "http"),
]

TG_DCS = [
    ("149.154.167.51", 443), ("149.154.175.53", 443),
    ("91.108.56.130", 443), ("149.154.171.5", 443),
]


# ---------------------------------------------------------------- 配置

def app_dir() -> Path:
    base = os.environ.get("APPDATA") or str(Path.home())
    p = Path(base) / "TelegramBackup"
    try:
        p.mkdir(parents=True, exist_ok=True)
    except Exception:
        p = Path.cwd()
    return p


CFG_PATH = app_dir() / "config.json"
SESSION_PATH = app_dir() / "session"


def load_cfg() -> dict:
    try:
        if CFG_PATH.exists():
            return json.loads(CFG_PATH.read_text(encoding="utf-8"))
    except Exception:
        pass
    return {}


def save_cfg(cfg: dict):
    try:
        CFG_PATH.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass


def desktop_dir() -> Path:
    home = Path.home()
    for cand in (home / "Desktop", home / "OneDrive" / "Desktop",
                 home / "OneDrive" / "桌面", home / "桌面"):
        if cand.exists():
            return cand
    return home


# ---------------------------------------------------------------- 网络

def port_open(port, host="127.0.0.1", timeout=0.35):
    try:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            s.settimeout(timeout)
            return s.connect_ex((host, port)) == 0
    except Exception:
        return False


def detect_proxy():
    for port, ptype in PROXY_PORTS:
        if port_open(port):
            return ptype, "127.0.0.1", port
    try:
        import winreg
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Internet Settings",
        )
        server, _ = winreg.QueryValueEx(key, "ProxyServer")
        winreg.CloseKey(key)
        if server:
            part = str(server).split(";")[-1]
            if "=" in part:
                part = part.split("=", 1)[1]
            host, _, port = part.partition(":")
            if port.isdigit() and port_open(int(port), host or "127.0.0.1"):
                return "http", host or "127.0.0.1", int(port)
    except Exception:
        pass
    return None


def socks5_connect(ph, pp, dh, dp, timeout=8):
    s = socket.create_connection((ph, pp), timeout)
    try:
        s.settimeout(timeout)
        s.sendall(b"\x05\x01\x00")
        if s.recv(2) != b"\x05\x00":
            raise OSError("代理不支持 SOCKS5")
        s.sendall(b"\x05\x01\x00\x01" + socket.inet_aton(dh) + dp.to_bytes(2, "big"))
        rep = s.recv(10)
        if len(rep) < 2 or rep[1] != 0:
            raise OSError("代理拒绝连接")
    finally:
        s.close()


def http_connect(ph, pp, dh, dp, timeout=8):
    s = socket.create_connection((ph, pp), timeout)
    try:
        s.settimeout(timeout)
        s.sendall(f"CONNECT {dh}:{dp} HTTP/1.1\r\nHost: {dh}:{dp}\r\n\r\n".encode())
        if b" 200" not in s.recv(200).split(b"\r\n")[0]:
            raise OSError("代理拒绝连接")
    finally:
        s.close()


def proxy_works(info) -> bool:
    if not info:
        return False
    try:
        if info.get("proxy_type") == "socks5":
            socks5_connect(info["addr"], info["port"], TG_DCS[0][0], TG_DCS[0][1])
        else:
            http_connect(info["addr"], info["port"], TG_DCS[0][0], TG_DCS[0][1])
        return True
    except Exception:
        return False


# ---------------------------------------------------------------- 导出

def chat_type(entity):
    if isinstance(entity, Channel):
        if entity.megagroup:
            return "公开群组" if entity.username else "私有群组"
        return "公开频道" if entity.username else "私有频道"
    if isinstance(entity, Chat):
        return "基础群组"
    return "会话"


def fmt_dt(dt):
    if not dt:
        return ""
    try:
        return dt.strftime("%Y-%m-%d %H:%M")
    except Exception:
        return str(dt)


async def make_invite(client, entity):
    if isinstance(entity, Chat):
        return "", "基础群无邀请链接"
    if not isinstance(entity, Channel):
        return "", "非群组"
    try:
        if _INVITE_IN_MESSAGES:
            res = await client(ExportChatInviteRequest(peer=entity))
        else:
            res = await client(ExportChatInviteRequest(entity.id))
        link = getattr(res, "link", "") or ""
        return (link, "已生成邀请链接") if link else ("", "无法生成")
    except errors.ChatAdminRequiredError:
        return "", "需管理员权限"
    except errors.ChatAdminForbiddenError:
        return "", "权限不足"
    except errors.FloodWaitError as exc:
        return "", f"限流，需等{exc.seconds}秒"
    except errors.RPCError as exc:
        return "", f"失败:{type(exc).__name__}"
    except Exception:
        return "", "未知错误"


HEAD_FILL = PatternFill("solid", fgColor="2F5597")
HEAD_FONT = Font(bold=True, color="FFFFFF", size=11)


def style_sheet(ws, headers, rows):
    ws.append(headers)
    for cell in ws[1]:
        cell.fill = HEAD_FILL
        cell.font = HEAD_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")
    for row in rows:
        ws.append(row)
    for idx, head in enumerate(headers, 1):
        longest = len(str(head))
        for row in rows:
            if idx - 1 < len(row):
                val = "" if row[idx - 1] is None else str(row[idx - 1])
                longest = max(longest, sum(2 if ord(ch) > 127 else 1 for ch in val))
        ws.column_dimensions[get_column_letter(idx)].width = min(max(10, longest + 3), 58)
    ws.freeze_panes = "A2"
    if rows:
        ws.auto_filter.ref = ws.dimensions
    ws.row_dimensions[1].height = 22


async def do_export(cfg, want_code, want_password, log, progress):
    api_id = int(cfg.get("api_id") or DEFAULT_API_ID)
    api_hash = cfg.get("api_hash") or DEFAULT_API_HASH
    phone = cfg["phone"]
    proxy = cfg.get("proxy")

    client = TelegramClient(
        str(SESSION_PATH), api_id, api_hash, proxy=proxy,
        device_model="Backup Tool", system_version="Windows 10",
        app_version=VERSION, lang_code="zh",
    )

    log("正在连上电报服务器 ...")
    await client.connect()
    if not await client.is_user_authorized():
        log("第一次使用，需要验证一下。电报会给你发一条验证码消息。")
        try:
            await client.send_code_request(phone)
        except errors.FloodWaitError as exc:
            raise RuntimeError(f"电报让你等 {exc.seconds} 秒，稍后再点一次。")
        except Exception as exc:
            raise RuntimeError(f"发不出验证码：{exc}\n（先检查代理软件有没有打开）")

        code = want_code()
        if not code:
            raise RuntimeError("没有输入验证码，已取消。")
        try:
            await client.sign_in(phone=phone, code=code.strip())
        except errors.SessionPasswordNeededError:
            pw = want_password()
            if not pw:
                raise RuntimeError("没有输入两步验证密码，已取消。")
            await client.sign_in(password=pw)
        except errors.PhoneCodeInvalidError:
            raise RuntimeError("验证码不对，再点一次导出重新发送。")
        except errors.PhoneCodeExpiredError:
            raise RuntimeError("验证码过期了，再点一次导出重新发送。")

    me = await client.get_me()
    log(f"已登录：{me.first_name or ''} {('@' + me.username) if me.username else ''}")

    log("正在读取你的会话列表 ...")
    dialogs = []
    async for d in client.iter_dialogs():
        dialogs.append(d)
    total = len(dialogs)
    log(f"共 {total} 个会话，开始整理 ...")
    progress(0, total)

    groups, bots = [], []
    for i, dialog in enumerate(dialogs, 1):
        entity = dialog.entity
        name = dialog.name or getattr(entity, "title", "") or ""
        last = fmt_dt(dialog.date)

        if isinstance(entity, User):
            if getattr(entity, "bot", False):
                bots.append({
                    "username": getattr(entity, "username", None) or "",
                    "name": name, "last": last, "id": entity.id,
                })
            progress(i, total)
            continue

        if not isinstance(entity, (Channel, Chat)):
            progress(i, total)
            continue

        username = getattr(entity, "username", None) or ""
        ctype = chat_type(entity)
        if username:
            link, state = f"https://t.me/{username}", "公开链接"
        else:
            link, state = await make_invite(client, entity)
            await asyncio.sleep(0.35)

        groups.append({
            "type": ctype, "name": name, "username": username,
            "link": link, "link_state": state,
            "members": getattr(entity, "participants_count", "") or "",
            "last": last, "id": entity.id,
        })
        progress(i, total)

    log("正在读取联系人 ...")
    contacts = []
    try:
        for user in await client.get_contacts():
            if not isinstance(user, User):
                continue
            full = " ".join(x for x in [user.first_name, user.last_name] if x)
            contacts.append({
                "name": full,
                "phone": ("+" + user.phone) if getattr(user, "phone", None) else "",
                "username": getattr(user, "username", None) or "",
                "id": user.id,
            })
    except Exception as exc:
        log(f"联系人读取失败（不影响其它表格）：{exc}")

    await client.disconnect()

    # ---------------- 写表格
    g_head = ["类型", "名称", "@用户名", "链接", "链接状态", "成员数", "最后消息", "ID"]
    b_head = ["@用户名", "昵称", "最后消息", "ID"]
    c_head = ["姓名", "手机号", "@用户名", "ID"]

    g_rows = [[g["type"], g["name"], ("@" + g["username"]) if g["username"] else "",
               g["link"], g["link_state"], g["members"], g["last"], g["id"]] for g in groups]
    b_rows = [[("@" + b["username"]) if b["username"] else "", b["name"],
               b["last"], b["id"]] for b in bots]
    c_rows = [[c["name"], c["phone"], ("@" + c["username"]) if c["username"] else "",
               c["id"]] for c in contacts]

    wb = Workbook()
    ws1 = wb.active
    ws1.title = "群组和频道"
    style_sheet(ws1, g_head, g_rows)
    style_sheet(wb.create_sheet("机器人"), b_head, b_rows)
    style_sheet(wb.create_sheet("联系人"), c_head, c_rows)

    ws4 = wb.create_sheet("统计")
    public_cnt = sum(1 for g in groups if g["username"])
    link_cnt = sum(1 for g in groups if g["link"])
    meta = [
        ("导出时间", datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
        ("账号", f"{me.first_name or ''} {('@' + me.username) if me.username else ''}"),
        ("账号ID", me.id),
        ("会话总数", total),
        ("群组和频道", len(groups)),
        ("其中公开(有用户名)", public_cnt),
        ("能拿到链接的", link_cnt),
        ("机器人", len(bots)),
        ("联系人", len(contacts)),
        ("工具版本", VERSION),
    ]
    ws4["A1"], ws4["B1"] = "项目", "数值"
    for cell in ("A1", "B1"):
        ws4[cell].fill = HEAD_FILL
        ws4[cell].font = HEAD_FONT
    for i, (k, v) in enumerate(meta, start=2):
        ws4[f"A{i}"] = k
        ws4[f"B{i}"] = v
    ws4.column_dimensions["A"].width = 24
    ws4.column_dimensions["B"].width = 44

    stamp = datetime.now().strftime("%Y%m%d_%H%M")
    out = desktop_dir() / f"电报备份_{stamp}.xlsx"
    log("正在保存表格 ...")
    wb.save(out)

    log("")
    log("=" * 34)
    log(f"群组/频道：{len(groups)} 个（公开 {public_cnt}，拿到链接 {link_cnt}）")
    log(f"机器人：{len(bots)} 个")
    log(f"联系人：{len(contacts)} 个")
    log(f"文件已保存到桌面：{out.name}")
    log("=" * 34)
    return out, len(groups), len(bots), len(contacts)


# ---------------------------------------------------------------- 界面

class App:
    def __init__(self, root):
        self.root = root
        self.q = queue.Queue()
        self.busy = False
        self.auth = {"event": None, "value": ""}
        self.advanced_on = False
        self.result = None

        cfg = load_cfg()

        root.title(f"{APP_NAME} v{VERSION}")
        root.geometry("680x540")
        root.minsize(620, 500)
        root.protocol("WM_DELETE_WINDOW", self.on_close)
        try:
            ttk.Style().theme_use("vista")
        except Exception:
            pass

        # 标题
        head = ttk.Frame(root)
        head.pack(fill="x", padx=18, pady=(16, 2))
        ttk.Label(head, text="把电报里关注的群、频道、机器人、联系人",
                  font=("Microsoft YaHei", 14, "bold")).pack(anchor="w")
        ttk.Label(head, text="一次导出成一张 Excel 表，自动存到桌面。不需要懂技术。",
                  foreground="#777").pack(anchor="w", pady=(2, 0))

        # 手机号
        box = ttk.LabelFrame(root, text=" 第一步：填你的手机号 ")
        box.pack(fill="x", padx=16, pady=(14, 6))

        row = ttk.Frame(box)
        row.pack(fill="x", padx=12, pady=10)
        self.v_region = tk.StringVar(value=cfg.get("region") or "中国 +86")
        saved = cfg.get("phone", "") or ""
        saved_code = REGION_MAP.get(self.v_region.get(), "+86")
        self.v_phone = tk.StringVar(
            value=saved.replace(saved_code, "", 1).lstrip("+") if saved else "")

        self.cb = ttk.Combobox(row, textvariable=self.v_region, width=14, state="readonly",
                               values=[r[0] for r in REGIONS])
        self.cb.pack(side="left")
        self.cb.bind("<<ComboboxSelected>>", lambda e: self.refresh_preview())

        self.en_phone = ttk.Entry(row, textvariable=self.v_phone, width=20,
                                  font=("Consolas", 12))
        self.en_phone.pack(side="left", padx=(10, 0))
        ttk.Label(row, text="只填号码就行", foreground="#999").pack(side="left", padx=8)

        self.lb_pre = ttk.Label(box, text="", foreground="#2a7", font=("Consolas", 10))
        self.lb_pre.pack(anchor="w", padx=14, pady=(0, 10))
        self.v_phone.trace_add("write", lambda *a: self.refresh_preview())
        self.refresh_preview()

        # 按钮
        mid = ttk.Frame(root)
        mid.pack(fill="x", padx=16, pady=(4, 4))
        self.btn = ttk.Button(mid, text="一键导出", command=self.on_export)
        self.btn.pack(side="left", ipadx=26, ipady=6)

        self.pb = ttk.Progressbar(mid, mode="determinate", length=230)
        self.pb.pack(side="left", padx=14)
        self.lb_status = ttk.Label(mid, text="准备好了", foreground="#555")
        self.lb_status.pack(side="left")

        # 验证码区
        self.frame_auth = ttk.LabelFrame(root, text=" 需要你输入 ")
        self.lb_auth = ttk.Label(self.frame_auth, text="", wraplength=600, justify="left")
        self.lb_auth.pack(anchor="w", padx=12, pady=(8, 4))
        arow = ttk.Frame(self.frame_auth)
        arow.pack(fill="x", padx=12, pady=(0, 10))
        self.v_code = tk.StringVar()
        self.en_code = ttk.Entry(arow, textvariable=self.v_code, width=22,
                                 font=("Consolas", 13))
        self.en_code.pack(side="left")
        self.en_code.bind("<Return>", lambda e: self.submit_code())
        ttk.Button(arow, text="确定", command=self.submit_code).pack(side="left", padx=8)
        ttk.Label(arow, text="填完按回车", foreground="#999").pack(side="left")

        # 日志
        self.log = ScrolledText(root, height=7, font=("Consolas", 9), state="disabled",
                                bg="#1e1e1e", fg="#d4d4d4")
        self.log.pack(fill="both", expand=True, padx=16, pady=(6, 4))
        self.write_log("点上面的「一键导出」开始。")
        self.write_log("过程会显示在这里，完成以后表格会自动出现在桌面。")

        # 高级
        self.chk_adv = ttk.Checkbutton(root, text="高级设置（不用懂，出问题才看）",
                                       command=self.toggle_advanced)
        self.chk_adv.pack(anchor="w", padx=18)

        self.frame_adv = ttk.LabelFrame(root, text=" 程序身份 ")
        a1 = ttk.Frame(self.frame_adv)
        a1.pack(fill="x", padx=12, pady=6)
        ttk.Label(a1, text="api_id").pack(side="left")
        self.v_id = tk.StringVar(value=str(cfg.get("api_id") or DEFAULT_API_ID))
        ttk.Entry(a1, textvariable=self.v_id, width=12).pack(side="left", padx=(6, 14))
        ttk.Label(a1, text="api_hash").pack(side="left")
        self.v_hash = tk.StringVar(value=cfg.get("api_hash") or DEFAULT_API_HASH)
        ttk.Entry(a1, textvariable=self.v_hash, width=34).pack(side="left", padx=6)
        ttk.Label(self.frame_adv,
                  text="这是程序的身份，不是你的密码，一般不用改。\n"
                       "默认用的是公开值，用的人多时电报会限流；换成自己申请的会更稳。",
                  foreground="#888", justify="left").pack(anchor="w", padx=12, pady=(0, 8))

        self.root.after(120, self.pump)

    # ---------------- 小工具
    def compose_phone(self):
        code = REGION_MAP.get(self.v_region.get(), "+86")
        raw = re.sub(r"[^\d+]", "", self.v_phone.get() or "")
        if raw.startswith("+"):
            return raw
        return code + raw.lstrip("0")

    def refresh_preview(self):
        digits = re.sub(r"\D", "", self.v_phone.get() or "")
        if not digits:
            self.lb_pre.configure(text="填好后这里会显示最终号码", foreground="#999")
            return
        raw = re.sub(r"[^\d+]", "", self.v_phone.get() or "")
        extra = "（开头的 0 已自动去掉）" if (not raw.startswith("+") and raw.startswith("0")) else ""
        self.lb_pre.configure(text=f"✓ 将使用：{self.compose_phone()} {extra}", foreground="#2a7")

    def write_log(self, msg):
        self.log.configure(state="normal")
        self.log.insert("end", str(msg) + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def toggle_advanced(self):
        self.advanced_on = not self.advanced_on
        if self.advanced_on:
            self.frame_adv.pack(fill="x", padx=16, pady=(2, 10), before=self.log)
        else:
            self.frame_adv.pack_forget()

    def open_advanced(self):
        if not self.advanced_on:
            self.chk_adv.state(["selected"])
            self.toggle_advanced()

    # ---------------- 事件循环
    def pump(self):
        try:
            while True:
                item = self.q.get_nowait()
                kind = item[0]
                if kind == "log":
                    self.write_log(item[1])
                elif kind == "prog":
                    cur, total = item[1], item[2]
                    self.pb.configure(maximum=max(total, 1), value=cur)
                    self.lb_status.configure(text=f"正在整理 {cur}/{total}")
                elif kind == "ask":
                    _, prompt, masked, ev = item
                    self.frame_auth.pack(fill="x", padx=16, pady=(0, 6), before=self.log)
                    self.lb_auth.configure(text=prompt)
                    self.v_code.set("")
                    self.en_code.configure(show="*" if masked else "")
                    self.en_code.focus_set()
                    self.auth = {"event": ev, "value": ""}
                elif kind == "done":
                    self.on_done(item[1])
                elif kind == "fail":
                    self.write_log("")
                    self.write_log("出错了：" + str(item[1]))
                    self.open_advanced()
                    self.reset("出错了，看下面的说明")
        except queue.Empty:
            pass
        self.root.after(120, self.pump)

    def submit_code(self):
        ev = self.auth.get("event")
        if not ev:
            return
        self.auth["value"] = self.v_code.get().strip()
        self.auth["event"] = None
        self.frame_auth.pack_forget()
        ev.set()

    def ask(self, prompt, masked=False):
        ev = threading.Event()
        self.q.put(("ask", prompt, masked, ev))
        ev.wait(timeout=1800)
        return self.auth.get("value", "")

    def reset(self, text):
        self.busy = False
        self.btn.configure(state="normal", text="一键导出")
        self.lb_status.configure(text=text)

    def on_done(self, payload):
        path, g, b, c = payload
        self.result = path
        self.reset("完成 ✔")
        self.pb.configure(value=self.pb["maximum"])
        try:
            subprocess.Popen(f'explorer /select,"{path}"')
        except Exception:
            pass
        messagebox.showinfo(
            "导出完成",
            f"已经存到桌面：\n{path.name}\n\n"
            f"群组/频道 {g} 个　机器人 {b} 个　联系人 {c} 个\n\n"
            "表格已经用 Excel 打开了，四张表分别是：\n"
            "群组和频道 / 机器人 / 联系人 / 统计",
        )

    def on_close(self):
        if self.busy and not messagebox.askokcancel("正在导出", "还在跑，确定要关掉吗？"):
            return
        self.root.destroy()

    # ---------------- 开始
    def on_export(self):
        if self.busy:
            return
        phone = self.compose_phone()
        if len(re.sub(r"\D", "", phone)) < 8:
            messagebox.showerror("手机号不对", "号码太短了，检查一下再试。")
            self.en_phone.focus_set()
            return

        proxy = detect_proxy()
        info = {"proxy_type": proxy[0], "addr": proxy[1], "port": proxy[2], "rdns": True} \
            if proxy else None

        cfg = {
            "api_id": int(self.v_id.get().strip() or DEFAULT_API_ID),
            "api_hash": self.v_hash.get().strip() or DEFAULT_API_HASH,
            "phone": phone,
            "region": self.v_region.get(),
            "proxy": info,
        }
        save_cfg(cfg)

        self.busy = True
        self.btn.configure(state="disabled", text="正在导出 ...")
        self.lb_status.configure(text="准备中 ...")
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")
        self.frame_auth.pack_forget()
        self.pb.configure(value=0, maximum=1)

        if info:
            self.write_log(f"网络：已自动使用本机代理 {info['addr']}:{info['port']}")
        else:
            self.write_log("网络：没找到代理软件，尝试直连（国内一般需要先开代理）")

        threading.Thread(target=self.worker, args=(cfg,), daemon=True).start()

    def worker(self, cfg):
        def log(m):
            self.q.put(("log", m))

        def prog(cur, total):
            self.q.put(("prog", cur, total))

        def code_cb():
            return self.ask("电报给你发了一条验证码消息（不是短信），在手机 App 里能看到。\n"
                            "把那几个数字填到下面：")

        def pw_cb():
            return self.ask("这个账号开了两步验证，请输入两步验证密码：", masked=True)

        try:
            payload = asyncio.run(do_export(cfg, code_cb, pw_cb, log, prog))
            self.q.put(("done", payload))
        except Exception as exc:
            msg = str(exc) or traceback.format_exc().strip().splitlines()[-1]
            self.q.put(("fail", msg))


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
