# -*- coding: utf-8 -*-
"""命令行进度条：tqdm 风格的「百分比 + 进度块 + 速度 + 剩余 + 已用」。

显示效果（对齐业务员熟悉的 tqdm 格式）::

    处理文件[德邦]:  43%|█████████▌          | 速度: 3.42it/s | 剩余: 01:12 | 已用: 00:58

设计要点：
- 只占一行，用回车符原地刷新；重定向到文件/被 IDE 捕获时降级成每 10% 打一行。
- **预计时间（剩余）按真实节流算**：每张至少 `config.MIN_SECONDS_PER_FILE` 秒，
  所以 ``剩余 = 剩余张数 × max(实测单张耗时, 节流秒数)``，业务员看到的是真实等待。
- **速度（it/s）只算干活时间**，不含节流 sleep，避免显示成 0.2it/s 让人以为卡死。
- 节流等待期间每 0.2 秒重绘一次，「剩余」一秒一秒往下走，看着像在动。
- 日志/警告要插进来时先 `clear()` 擦掉进度行，打完再 `redraw()`（见 log.py）。
- 中文字符占两列，擦行/补空格按显示宽度算，否则会留残影。
"""
from __future__ import annotations

import shutil
import sys
import time
import unicodedata

BAR_WIDTH = 30          # 终端够宽时用的块数
BAR_MIN = 10
LOG_STEP = 10           # 非终端环境下每 10% 打一行
SLEEP_FRAME = 0.2       # 节流等待时多久重绘一次

_ACTIVE: "Progress | None" = None

# 终端编码能用完整方块就用，否则退回 #
_FULL = "█"
_PARTIAL = ("", "▏", "▎", "▍", "▌", "▋", "▊", "▉")


def disp_width(text: str) -> int:
    """控制台显示宽度：中文/全角算 2 列，擦行时按这个补空格才不残留。"""
    return sum(2 if unicodedata.east_asian_width(ch) in ("W", "F") else 1 for ch in text)


def _fmt_time(seconds: float) -> str:
    """tqdm 风格的时间：mm:ss，超过 1 小时用 h:mm:ss。"""
    total = int(max(0, round(seconds)))
    h, rem = divmod(total, 3600)
    m, s = divmod(rem, 60)
    return f"{h}:{m:02d}:{s:02d}" if h else f"{m:02d}:{s:02d}"


def active_progress() -> "Progress | None":
    """当前正在渲染的进度条（供日志处理器擦行用）。"""
    return _ACTIVE


def set_active(progress: "Progress | None") -> None:
    global _ACTIVE
    _ACTIVE = progress


def _terminal_width(stream) -> int:
    try:
        return shutil.get_terminal_size((100, 20)).columns
    except Exception:  # noqa: BLE001 - 拿不到就当 100
        return 100


class Progress:
    """按「单元」（通常是文件）推进的进度条。

    total 为 0 时视为未知总量：只打印计数，不显示百分比。
    """

    def __init__(self, total: int = 0, stream=None, enabled: bool | None = None,
                 desc: str = "处理文件", unit: str = "it"):
        self.stream = stream if stream is not None else sys.stdout
        if enabled is None:
            enabled = bool(getattr(self.stream, "isatty", lambda: False)())
        self.enabled = enabled
        self.total = max(int(total or 0), 0)
        self.done = 0
        self.desc = desc
        self.unit = unit
        self.label = ""
        self.silent = False       # --no-progress：完全不出声（含收尾那行）
        self._start = time.monotonic()
        self._sleep = 0.0            # 节流等待累计（不算进速度）
        self._deadline: float | None = None
        self._last_logged = -1
        self._last_pct = None
        self._closed = False
        self._drawn_len = 0
        self._bar_w = BAR_WIDTH
        self._uni = self._can_unicode()
        set_active(self)

    # ---------------- 编码能力 ----------------
    def _can_unicode(self) -> bool:
        enc = (getattr(self.stream, "encoding", "") or "").lower().replace("-", "")
        if enc not in ("utf8", "utf8sig", "cp65001"):
            return False
        try:
            _FULL.encode(enc)
        except (UnicodeEncodeError, LookupError):
            return False
        return True

    # ---------------- 状态 ----------------
    @property
    def pct(self) -> float:
        if self.total <= 0:
            return 100.0 if self.done else 0.0
        return min(100.0, 100.0 * self.done / self.total)

    def set_total(self, total: int) -> None:
        self.total = max(int(total or 0), 0)
        self._refresh_deadline()

    def set_label(self, label: str) -> None:
        """把阶段/客户名写进描述后面，例如 ``处理文件[德邦]``。"""
        self.label = label or ""

    def advance(self, n: int = 1, label: str | None = None) -> None:
        """推进 n 个单元（默认 1）；label 不为 None 时顺带改描述。"""
        if label is not None:
            self.label = label
        self.done = min(self.done + max(int(n), 0), self.total) if self.total else self.done + n
        self._refresh_deadline()
        self.redraw()

    def advance_to(self, done: int, label: str | None = None) -> None:
        """补齐到至少 done（插件内部已推进过时不会倒退）。"""
        if label is not None:
            self.label = label
        if done > self.done:
            self.done = min(done, self.total) if self.total else done
        self._refresh_deadline()
        self.redraw()

    def sleep_tick(self, seconds: float) -> None:
        """节流等待期间由 pacing 回调：记下等待时长并重绘（「剩余」会往下走）。"""
        self._sleep += max(0.0, seconds)
        if self.enabled and not self.silent:
            self.redraw()

    def close(self, label: str | None = None) -> None:
        if self._closed:
            return
        self._closed = True
        if label is not None:
            self.label = label
        if self.total:
            self.done = self.total
        self._deadline = None
        self.redraw(force=True)
        if self.enabled:
            self._write("\n")
            self._drawn_len = 0
        set_active(None)

    # ---------------- 时间估算 ----------------
    def _refresh_deadline(self) -> None:
        """剩余时间 = 剩余单元数 × max(实测单张耗时, 节流秒数)。"""
        if not self.total or self.done >= self.total:
            self._deadline = None
            return
        from ..config import min_seconds_per_file  # 延迟导入，避免循环

        elapsed = time.monotonic() - self._start
        floor = min_seconds_per_file() if _pacing_on() else 0.0
        per = max(elapsed / self.done, floor) if self.done else floor
        if per <= 0:
            self._deadline = None
            return
        self._deadline = time.monotonic() + per * (self.total - self.done)

    def _eta(self) -> float:
        if self.total and self.done >= self.total:
            return 0.0
        if self._deadline is None:
            return 0.0
        return max(0.0, self._deadline - time.monotonic())

    def _rate(self) -> float:
        busy = max(0.001, (time.monotonic() - self._start) - self._sleep)
        return self.done / busy if self.done else 0.0

    # ---------------- 渲染 ----------------
    def _desc_text(self) -> str:
        text = self.desc
        if self.label:
            text = f"{text}[{self.label}]"
        return f"{text}: "

    @staticmethod
    def _bar(pct: float, width: int, uni: bool) -> str:
        if width <= 0:
            return ""
        pos = pct / 100.0 * width
        full = int(pos)
        frac = pos - full
        if uni:
            part = _PARTIAL[min(7, int(frac * 8))] if full < width else ""
            body = _FULL * full + part
            empty = " "
        else:
            body = "#" * full
            empty = "-"
        return body + empty * max(0, width - disp_width(body))

    def _line(self) -> str:
        pct = self.pct if self.total else 0.0

        stats = (f" 速度: {self._rate():.2f}{self.unit}/s"
                 f" | 剩余: {_fmt_time(self._eta())}"
                 f" | 已用: {_fmt_time(time.monotonic() - self._start)}")
        head = f"{self._desc_text()}{pct:3.0f}%|"
        tail = f"|{stats}"

        if not self.total:
            return f"{self._desc_text()}{self.done} {self.unit}{stats}"

        # 终端宽度不够就压缩块宽，至少留 BAR_MIN
        width = _terminal_width(self.stream)
        room = width - disp_width(head) - disp_width(tail) - 1
        bar_w = max(BAR_MIN, min(self._bar_w, room))
        return f"{head}{self._bar(pct, bar_w, self._uni)}{tail}"

    def clear(self) -> None:
        """擦掉当前行（日志插入前调用）。"""
        if not self.enabled or not self._drawn_len:
            return
        self._write("\r" + " " * self._drawn_len + "\r")
        self._drawn_len = 0

    def redraw(self, force: bool = False) -> None:
        if self.silent:
            return
        if not self.enabled:
            self._render_log(force)
            return
        line = self._line()
        try:
            pad = max(0, self._drawn_len - disp_width(line))  # 盖住上一行更长的残影
            self._write("\r" + line + " " * pad)
            self._drawn_len = disp_width(line)
        except OSError:
            self.enabled = False

    def _render_log(self, force: bool) -> None:
        """非终端：每 LOG_STEP 个百分点打一行，收尾必打。"""
        pct = int(self.pct)
        if pct == self._last_pct:
            return
        bucket = pct // LOG_STEP
        if not force and bucket <= self._last_logged:
            return
        self._last_logged = bucket
        self._last_pct = pct
        self._write(self._line() + "\n")

    def _write(self, text: str) -> None:
        try:
            self.stream.write(text)
            self.stream.flush()
        except (UnicodeEncodeError, ValueError):  # 奇怪的终端编码/已关闭
            pass


def _pacing_on() -> bool:
    from . import pacing

    return pacing.is_enabled()
