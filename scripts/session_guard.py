#!/usr/bin/env python
"""检查当前 Codex 会话的 rollout 体积，超阈值时提示另开会话。

拼版图被看图工具读进会话后，会以 base64 留在会话历史里并随每次请求重发。
历史过大后上游会拒绝请求或返回误导性报错（历史事故：19 张拼版共 28.8MB，
触发 “the web_search tool is not supported”）。做视觉核对前先跑本脚本；
超过阈值就先另开会话再继续，不要在同一个会话里累积拼版图片。
"""
import argparse
import os
import sys
from pathlib import Path

DEFAULT_THRESHOLD_MB = 20.0


def sessions_root():
    home = Path(os.environ.get("CODEX_HOME") or Path.home() / ".codex")
    return home / "sessions"


def find_rollout(root, session_id):
    """优先按会话 id 找；找不到（或没给）时退回最近写入的 rollout。"""
    if session_id:
        hits = [p for p in root.rglob(f"*{session_id}*.jsonl") if p.is_file()]
        if hits:
            return max(hits, key=lambda p: p.stat().st_mtime)
    files = [p for p in root.rglob("*.jsonl") if p.is_file()]
    return max(files, key=lambda p: p.stat().st_mtime) if files else None


def main():
    ap = argparse.ArgumentParser(description="检查当前 Codex 会话 rollout 是否过大")
    ap.add_argument("--session-id", default=os.environ.get("CODEX_SESSION_ID", ""),
                    help="Codex 线程 id；默认取环境变量 CODEX_SESSION_ID")
    ap.add_argument("--sessions-root", default=None, help="默认 $CODEX_HOME/sessions")
    ap.add_argument("--threshold-mb", type=float, default=DEFAULT_THRESHOLD_MB)
    a = ap.parse_args()

    root = Path(a.sessions_root).expanduser() if a.sessions_root else sessions_root()
    if not root.is_dir():
        print(f"未找到会话目录：{root}（跳过检查）")
        return 0
    rollout = find_rollout(root, a.session_id)
    if rollout is None:
        print(f"未找到 rollout 文件：{root}（跳过检查）")
        return 0

    mb = rollout.stat().st_size / 1024 / 1024
    if mb > a.threshold_mb:
        print(f"WARN 会话 {mb:.1f} MB 已超过 {a.threshold_mb:g} MB："
              "不要再在本会话看图，先另开会话再继续视觉核对")
        print(f"     {rollout}")
        return 1
    print(f"OK 会话 {mb:.1f} MB ≤ {a.threshold_mb:g} MB：{rollout.name}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
