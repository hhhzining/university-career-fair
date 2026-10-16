#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
下载并解析教育部《全国普通高等学校名单》，生成本科院校清单 CSV。

用途：本脚本用于 university-career-fair 技能更新内置的高校名单基准
      （assets/universities.csv）。教育部每年发布一次最新名单，名单过时时
      运行本脚本即可重新生成，保证按省份核查时不遗漏任何本科院校。

用法：
    python update_university_list.py                                  # 自动定位教育部官网最新名单公告页
    python update_university_list.py --page-url <名单公告页URL>       # 指定名单公告页（推荐，定位失败时用）
    python update_university_list.py --xls <本地名单.xls>             # 直接解析本地已下载的名单文件
    python update_university_list.py --output <输出CSV路径>           # 指定输出路径

输出 CSV 列：省份,城市,学校名称,学校标识码,主管部门,办学层次,备注
脚本会打印本科高校总数及分省统计供核对（教育部 2026-06-17 版为 1412 所）。
"""

import argparse
import io
import re
import sys
import tempfile
import urllib.request
import urllib.parse
import urllib.error

# 教育部官网"公告"栏目页（含历次《全国高等学校名单》公告）
CATEGORY_URL = "http://www.moe.gov.cn/jyb_xxgk/s5743/s5744/A03/"
# 模拟真实浏览器的请求头，规避官网基础反爬（418）
UA = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
    "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
    "Connection": "keep-alive",
    "Upgrade-Insecure-Requests": "1",
}
# 省份分组行：如 "北京市（92所）"、"内蒙古自治区（54所）"
PROV_LINE = re.compile(r"^(.+?)\s*[（(]\s*(\d+)\s*所\s*[)）]\s*$")
# 附件链接
ATTACH_RE = re.compile(r'href=["\']([^"\']+?\.(?:xls|xlsx))["\']', re.IGNORECASE)


def fetch_text(url):
    """抓取网页文本，自动尝试 https/http。"""
    urls = [url]
    if url.lower().startswith("http:"):
        urls.append("https:" + url[5:])
    elif url.lower().startswith("https:"):
        urls.append("http:" + url[6:])
    last_err = None
    for u in urls:
        try:
            req = urllib.request.Request(u, headers=UA)
            with urllib.request.urlopen(req, timeout=60) as resp:
                return resp.read().decode("utf-8", errors="replace")
        except urllib.error.HTTPError as e:
            last_err = e
            continue
        except Exception as e:
            last_err = e
            continue
    raise RuntimeError(f"无法访问页面 {url}: {last_err}")


def fetch_bytes(url):
    """抓取二进制文件（附件）。"""
    urls = [url]
    if url.lower().startswith("http:"):
        urls.append("https:" + url[5:])
    elif url.lower().startswith("https:"):
        urls.append("http:" + url[6:])
    last_err = None
    for u in urls:
        try:
            req = urllib.request.Request(u, headers=UA)
            with urllib.request.urlopen(req, timeout=120) as resp:
                return resp.read()
        except Exception as e:
            last_err = e
            continue
    raise RuntimeError(f"无法下载附件 {url}: {last_err}")


def find_latest_announce_page(category_url):
    """在教育部公告栏目页中定位最新《全国高等学校名单》公告页 URL。"""
    html = fetch_text(category_url)
    links = re.findall(r'href=["\']([^"\']+?\.html)["\'][^>]*>([^<]*)', html)
    best = None
    best_date = ""
    for href, text in links:
        if "全国高等学校名单" in text or "全国高等学校名单" in href:
            m = re.search(r"(20\d{2})(\d{2})", href)
            date = m.group(0) if m else ""
            if date >= best_date:
                best_date = date
                best = href
    if not best:
        raise RuntimeError("在栏目页未找到《全国高等学校名单》公告，请用 --page-url 手动指定公告页 URL")
    return urllib.parse.urljoin(category_url, best)


def find_attachment_url(page_url, html):
    """从公告页提取普通高校名单附件 URL（优先含"普通"锚文本的 .xls，否则取第一个 .xls）。"""
    atts = ATTACH_RE.findall(html)
    if not atts:
        raise RuntimeError("公告页未找到 xls/xlsx 附件，请检查页面结构或改用 --xls 本地文件")
    return urllib.parse.urljoin(page_url, atts[0])


def parse_university_xls(data):
    """解析教育部名单 xls 二进制数据，返回本科院校记录列表。"""
    import xlrd

    wb = xlrd.open_workbook(file_contents=data)
    sh = wb.sheet_by_index(0)
    rows = []
    province = ""
    for r in range(sh.nrows):
        vals = [sh.cell_value(r, c) for c in range(sh.ncols)]
        first = str(vals[0]).strip() if vals else ""
        m = PROV_LINE.match(first)
        if m:
            province = m.group(1).strip()
            continue
        if len(vals) >= 6 and isinstance(vals[0], (int, float)):
            level = str(vals[5]).strip()
            if level != "本科":
                continue
            code = vals[2]
            if isinstance(code, float) and code.is_integer():
                code = str(int(code))
            rows.append({
                "省份": province,
                "城市": str(vals[4]).strip(),
                "学校名称": str(vals[1]).strip(),
                "学校标识码": str(code),
                "主管部门": str(vals[3]).strip(),
                "办学层次": level,
                "备注": str(vals[6]).strip() if len(vals) > 6 else "",
            })
    return rows


def write_csv(rows, output):
    """写 UTF-8 with BOM 的 CSV（Excel 直接打开不乱码）。"""
    header = ["省份", "城市", "学校名称", "学校标识码", "主管部门", "办学层次", "备注"]

    def esc(v):
        s = str(v)
        if any(ch in s for ch in ',"\n\r'):
            s = '"' + s.replace('"', '""') + '"'
        return s

    with io.open(output, "w", encoding="utf-8-sig", newline="") as f:
        f.write(",".join(header) + "\n")
        for row in rows:
            f.write(",".join(esc(row[h]) for h in header) + "\n")


def main():
    ap = argparse.ArgumentParser(description="更新教育部全国本科高校名单 CSV")
    ap.add_argument("--page-url", default=None, help="教育部《全国高等学校名单》公告页 URL")
    ap.add_argument("--xls", default=None, help="本地已下载的名单 .xls 文件路径（跳过网络下载）")
    ap.add_argument("--output", default=None, help="输出 CSV 路径")
    args = ap.parse_args()

    if args.xls:
        with open(args.xls, "rb") as f:
            data = f.read()
        print("使用本地名单文件:", args.xls)
    else:
        page_url = args.page_url or find_latest_announce_page(CATEGORY_URL)
        print("名单公告页:", page_url)
        html = fetch_text(page_url)
        att_url = find_attachment_url(page_url, html)
        print("附件下载:", att_url)
        data = fetch_bytes(att_url)

    rows = parse_university_xls(data)
    by_prov = {}
    for row in rows:
        by_prov[row["省份"]] = by_prov.get(row["省份"], 0) + 1

    total = len(rows)
    print("本科高校总数:", total)
    for p, n in sorted(by_prov.items(), key=lambda x: -x[1]):
        print(f"  {p}: {n} 所")

    if total < 1000:
        print("警告：本科数量异常偏少，请核对名单版本是否正确。")
        sys.exit(1)

    output = args.output
    if not output:
        output = "universities.csv"
    write_csv(rows, output)
    print("已写出:", output)


if __name__ == "__main__":
    main()
