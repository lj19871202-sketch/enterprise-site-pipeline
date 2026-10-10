#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""企业官网资料包 · 域名规则（local_pipeline.py 与 gates.py 共用，避免黑名单漂移）。

聚合目录站（黄页/工商信息/名录/B2B）会把目标企业名写进自己的列表页，
所以“公司名命中正文”无法把它们与真实官网区分；必须在候选与门禁两层
统一识别并排除，否则整站公共素材会污染交付。

判定采用“精确域名后缀 + 域名标签边界”，避免把 a58.com、company-china.cn、
x1688.com 这类合法企业域名误判为目录站。
"""
import re
from urllib.parse import urlparse

MULTI_TLD = ("com.cn", "net.cn", "org.cn", "gov.cn", "edu.cn")

# 已知目录站主域；命中主域本身或任意子域。
DIRECTORY_SUFFIXES = (
    "huangye88.com", "huangye88.cn", "huangye88.com.cn",
    "shuidi.cn", "shuididp.cn", "shuididp.com", "shuidichou.com",
    "qixin.com", "qcc.com", "qichacha.com", "aiqicha.baidu.com",
    "tianyancha.com", "mingluji.com", "qianlima.com", "11467.com",
    "51sole.com", "made-in-china.com", "1688.com", "alibaba.com",
    "hc360.com", "chemnet.com", "china.cn", "58.com", "ganji.com",
    "zhaopin.com", "jobui.com", "bmlink.com", "net114.com", "ebdoor.com",
    "jdzj.com", "chinabidding.com", "bidcenter.com.cn", "zgong.com",
    "goepe.com", "b2b168.com", "sunnet.cn", "chinaso.com",
    "baike.baidu.com", "baike.so.com", "wikipedia.org", "wikiwand.com",
    "doc88.com", "book118.com", "renrendoc.com", "cnki.net",
    "wanfangdata.com.cn",
)

# 只在完整域名标签上命中的品牌词（避免 a58.com / x1688.com 等误伤）。
DIRECTORY_LABEL_RE = re.compile(
    r"(?:^|[.-])(?:huangye|shuidi|shuididp|shuidichou|qixin|qcc|qichacha|"
    r"aiqicha|tianyancha|mingluji|qianlima|11467|51sole|made-in-china|"
    r"1688|alibaba|hc360|chemnet|58|ganji|zhaopin|jobui|bmlink|net114|"
    r"ebdoor|jdzj|chinabidding|bidcenter|zgong|goepe|b2b168|sunnet|"
    r"chinaso|baike|wikipedia|wikiwand|doc88|book118|renrendoc|cnki|"
    r"wanfangdata)(?=[.-]|$)",
    re.I,
)

# 兼容旧调用点；新代码应直接用 is_directory_host()。
DIRECTORY_DOMAIN = DIRECTORY_LABEL_RE


def host_of(url):
    """取 URL 的主机名（去掉 userinfo 与端口）。"""
    try:
        return urlparse(str(url or "")).netloc.lower().split("@")[-1].split(":")[0]
    except Exception:
        return ""


def registrable(host):
    """取可注册域名，用于去重（www.injet.cn 与 injet.cn 同一个）。"""
    host = (host or "").lower().strip(".")
    if not host:
        return ""
    parts = host.split(".")
    if len(parts) >= 3 and ".".join(parts[-2:]) in MULTI_TLD:
        return ".".join(parts[-3:])
    if len(parts) >= 2:
        return ".".join(parts[-2:])
    return host


def same_host(a, b):
    """同一主域（含子域）判定。"""
    return bool(a and b and (a == b or a.endswith("." + b) or b.endswith("." + a)))


def is_directory_host(url_or_host):
    """判断域名是否属于 B2B/黄页/工商/名录等聚合目录站。"""
    host = (url_or_host or "").strip().lower().strip(".")
    if "://" in host:
        host = host_of(host)
    if not host:
        return False
    for suffix in DIRECTORY_SUFFIXES:
        if host == suffix or host.endswith("." + suffix):
            return True
    return bool(DIRECTORY_LABEL_RE.search(host))
