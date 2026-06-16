#!/usr/bin/env python3
"""
Turbolink 活动数据拉取脚本
单表纵向堆叠：每个活动的数据依次排列，一行=一个活动+任务
数据统计使用活动本身的起止时间

用法：
  python3 fetch.py --token "Bearer xxx" --project-id "abc123"
  python3 fetch.py --token "Bearer xxx" --project-id "abc123" --coin-cutoff 2026-06-11 --output report.xlsx
"""

import argparse
import sys
import requests
import time
from datetime import datetime, timedelta
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

BASE_URL = "https://api.branchcn.com"
SLASH = "/"

# ============ 带重试的 Session ============

def make_session():
    s = requests.Session()
    retries = Retry(total=5, backoff_factor=1, status_forcelist=[500, 502, 503, 504],
                    allowed_methods=["GET", "POST"])
    adapter = HTTPAdapter(max_retries=retries)
    s.mount("https://", adapter)
    s.mount("http://", adapter)
    return s

SESSION = make_session()


def safe_request(method, url, **kwargs):
    for attempt in range(5):
        try:
            if method == "get":
                resp = SESSION.get(url, **kwargs)
            else:
                resp = SESSION.post(url, **kwargs)
            resp.raise_for_status()
            return resp
        except Exception as e:
            if attempt < 4:
                wait = 2 ** attempt
                print(f"    请求失败 ({e})，{wait}s 后重试...")
                time.sleep(wait)
            else:
                raise


def build_headers(token, project_id):
    return {
        "accept": "application/json, text/plain, */*",
        "accept-language": "en,zh-CN;q=0.9,zh;q=0.8",
        "authorization": token,
        "origin": "https://dashboard.turbolink.cc",
        "referer": "https://dashboard.turbolink.cc/",
        "sec-ch-ua": '"Google Chrome";v="149", "Chromium";v="149", "Not)A;Brand";v="24"',
        "sec-ch-ua-mobile": "?0",
        "sec-ch-ua-platform": '"macOS"',
        "sec-fetch-dest": "empty",
        "sec-fetch-mode": "cors",
        "sec-fetch-site": "cross-site",
        "user-agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/149.0.0.0 Safari/537.36",
        "web-set": f"lang=zh-cn;pjid={project_id}",
    }


def parse_token_expiry(token):
    """解析 JWT token 的过期时间，返回 datetime 或 None"""
    try:
        import base64, json
        parts = token.replace("Bearer ", "").split(".")
        if len(parts) != 3:
            return None
        payload = parts[1]
        payload += "=" * (4 - len(payload) % 4)
        data = json.loads(base64.urlsafe_b64decode(payload))
        exp = data.get("exp")
        if exp:
            return datetime.fromtimestamp(exp)
    except Exception:
        pass
    return None


# ============ 数据拉取 ============

def fetch_fission_type_map(headers):
    url = f"{BASE_URL}/admin/fission/list"
    resp = safe_request("get", url, headers=headers)
    mapping = {}
    for item in resp.json().get("data", {}).get("list", []):
        mapping[item["mark"]] = item["title"]
    return mapping


def fetch_campaign_list(headers, project_id, start, end):
    url = f"{BASE_URL}/admin/fission-campaign/list"
    all_campaigns = []
    page = 1
    while True:
        params = [
            ("page", page), ("per_page", 50),
            ("mate_id", project_id), ("project_id", project_id),
            ("start", start), ("end", end),
            ("targets[]", 1), ("targets[]", 2), ("targets[]", 3),
        ]
        resp = safe_request("get", url, headers=headers, params=params)
        data = resp.json().get("data", {})
        items = data.get("list", [])
        if not items:
            break
        all_campaigns.extend(items)
        total = data.get("total", 0)
        print(f"  活动列表第 {page} 页: 获取 {len(items)} 条 (累计 {len(all_campaigns)}/{total})")
        if len(all_campaigns) >= total:
            break
        page += 1
    return all_campaigns


def filter_campaigns(campaigns, type_map, uv_threshold):
    filtered = []
    print(f"\n过滤条件: UV > {uv_threshold}")
    print(f"{'活动类型':<16} {'活动标题':<24} {'UV':>8} {'状态':>6}")
    print("-" * 60)
    for c in campaigns:
        name = c.get("title", "未知活动")
        fmark = c.get("fission_mark", "")
        act_type = type_map.get(fmark, fmark)
        rid = c.get("id")
        start = c.get("start_date", "")[:10]
        end = c.get("end_date", "")[:10]
        uv = c.get("visitor_num", 0) or 0
        if not rid:
            continue
        status = "保留" if uv > uv_threshold else "跳过"
        print(f"  {act_type:<14} {name:<24} {uv:>8} {status:>6}")
        if uv > uv_threshold:
            filtered.append({
                "activity": act_type,
                "fission_mark": fmark,
                "start": start.replace("/", "-"),
                "end": end.replace("/", "-"),
                "report_id": str(rid),
                "uv": uv,
            })
    return filtered


def fetch_uv(headers, report_id, start, end):
    url = f"{BASE_URL}/fbi/report/overview"
    params = {"id": report_id, "start": start, "end": end, "t_offset": 8}
    resp = safe_request("get", url, headers=headers, params=params)
    for item in resp.json().get("data", {}).get("list", []):
        if item.get("mark") == "uv":
            return item.get("total", 0)
    return 0


def fetch_metric(headers, report_id, metric_type, dems, start, end):
    url = f"{BASE_URL}/fbi/report/custom"
    payload = {"id": report_id, "type": metric_type, "start": start, "end": end, "dems": dems, "t_offset": 8}
    resp = safe_request("post", url, headers=headers, json=payload)
    for row in resp.json().get("data", {}).get("list", []):
        if all(t.get("name") == "all" for t in row.get("title", [])):
            return row.get("user_num", 0)
    return 0


def fetch_task_data(headers, report_id, start, end):
    url = f"{BASE_URL}/fbi/report/custom"
    payload = {"id": report_id, "type": 11, "start": start, "end": end, "dems": ["cond_id"], "t_offset": 8}
    resp = safe_request("post", url, headers=headers, json=payload)
    tasks = []
    for row in resp.json().get("data", {}).get("list", []):
        if all(t.get("name") == "all" for t in row.get("title", [])):
            continue
        tasks.append({
            "task_name": row["title"][0]["name"] if row.get("title") else "未知任务",
            "user_num": row.get("user_num", 0),
            "num": row.get("num", 0),
        })
    return tasks


def fetch_sign_in_tasks(headers, report_id, start, end):
    url = f"{BASE_URL}/fbi/report/custom"
    payload = {"id": report_id, "type": 14, "start": start, "end": end, "dems": ["level_mark"], "t_offset": 8}
    resp = safe_request("post", url, headers=headers, json=payload)
    for row in resp.json().get("data", {}).get("list", []):
        if all(t.get("name") == "all" for t in row.get("title", [])):
            return [{"task_name": "签到参与", "user_num": row.get("user_num", 0), "num": row.get("num", 0)}]
    return [{"task_name": "签到参与", "user_num": 0, "num": 0}]


def fetch_grow_data(headers, report_id, start, end):
    url = f"{BASE_URL}/fbi/report/custom"
    result = {"click": 0, "adopt": 0}
    for mark_type, key in [(16, "click"), (30, "adopt")]:
        payload = {"id": report_id, "type": mark_type, "start": start, "end": end, "dems": ["level_mark"], "t_offset": 8}
        resp = safe_request("post", url, headers=headers, json=payload)
        for row in resp.json().get("data", {}).get("list", []):
            titles = row.get("title", [])
            if all(t.get("name") == "all" for t in titles):
                result[key] = row.get("user_num", 0)
                break
    return result


def fetch_stay_data(headers, report_id, start, end):
    url = f"{BASE_URL}/fbi/report/stay"
    payload = {"id": report_id, "start": start, "end": end}
    resp = safe_request("post", url, headers=headers, json=payload)
    target = (datetime.strptime(start, "%Y-%m-%d") + timedelta(days=1)).strftime("%Y-%m-%d")
    for row in resp.json().get("data", {}).get("list", []):
        if row.get("date") == target:
            def parse_pct(val):
                val = str(val).replace("%", "").strip()
                return float(val) / 100 if val and val != "-" else 0
            return {
                "day1": parse_pct(row.get("day_1", "0%")),
                "day7": parse_pct(row.get("day_7", "0%")),
            }
    return {"day1": 0, "day7": 0}


def fetch_coin_recovery(headers, report_id, cutoff_date):
    """金币大派送：回收金币 = 用户投放金币总数 - 用户领取金币总数"""
    url = f"{BASE_URL}/admin/fission-stat/reward-detail"
    params = {"id": report_id}
    resp = safe_request("get", url, headers=headers, params=params)
    data = resp.json().get("data", {})
    pools = data.get("pools", [])
    cutoff = datetime.strptime(cutoff_date, "%Y-%m-%d")
    total_user = 0
    total_get = 0
    for p in pools:
        date_str = p["title"].split("(")[0]
        pool_date = datetime.strptime(date_str, "%Y-%m-%d")
        if pool_date > cutoff:
            continue
        total_user += p.get("user_coin", 0)
        total_get += p.get("get_coin", 0)
    return total_user - total_get, total_user


def fetch_activity(headers, act, coin_cutoff):
    rid = act["report_id"]
    s, e = act["start"], act["end"]
    fmark = act.get("fission_mark", "")
    is_sign_in = (fmark == "sign_in")
    is_grow = (fmark == "keep")
    is_coin = (fmark == "coin")
    stay = fetch_stay_data(headers, rid, s, e)

    if is_sign_in:
        tasks = fetch_sign_in_tasks(headers, rid, s, e)
        click = 0
    else:
        tasks = fetch_task_data(headers, rid, s, e)
        click = fetch_metric(headers, rid, 5, ["channel_id"], s, e)

    grow_click = 0
    grow_adopt = 0
    if is_grow:
        grow = fetch_grow_data(headers, rid, s, e)
        grow_click = grow["click"]
        grow_adopt = grow["adopt"]

    coin_recovery = 0
    coin_user_total = 0
    if is_coin and coin_cutoff:
        coin_recovery, coin_user_total = fetch_coin_recovery(headers, rid, coin_cutoff)

    return {
        "activity": act["activity"],
        "period": f"{s} ~ {e}",
        "is_grow": is_grow,
        "is_coin": is_coin,
        "uv": act.get("uv", 0),
        "click": click,
        "grow_click": grow_click,
        "grow_adopt": grow_adopt,
        "tasks": tasks,
        "day1": stay["day1"], "day7": stay["day7"],
        "coin_recovery": coin_recovery,
        "coin_user_total": coin_user_total,
    }


# ============ Excel ============

def set_cell(ws, row, col, value=None, fmt=None):
    cell = ws.cell(row=row, column=col)
    if value is not None:
        cell.value = value
    cell.border = Border(left=Side(style="thin"), right=Side(style="thin"),
                         top=Side(style="thin"), bottom=Side(style="thin"))
    if fmt:
        cell.number_format = fmt
    return cell


def build_excel(all_data, output_file, has_coin):
    wb = Workbook()
    ws = wb.active
    ws.title = "活动数据"

    pct = "0.00%"
    hfill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    hfont = Font(bold=True, size=11, color="FFFFFF")

    headers = [
        "活动名称", "活动时间", "UV",                              # A-C
        "点击活动主按钮人数", "活动参与率",                          # D-E
        "任务名称", "完成任务人数", "任务完成率", "完成任务次数",      # F-I
        "次留", "7日留",                                            # J-K
    ]
    if has_coin:
        headers.extend(["回收金币", "金币回收比例"])

    for i, h in enumerate(headers, 1):
        c = ws.cell(row=1, column=i, value=h)
        c.font = hfont
        c.fill = hfill
        c.alignment = Alignment(horizontal="center")
        c.border = Border(left=Side(style="thin"), right=Side(style="thin"),
                          top=Side(style="thin"), bottom=Side(style="thin"))

    coin_col_start = 12  # L column if has_coin

    row = 2
    for d in all_data:
        first_row = row
        tasks = d["tasks"]
        is_grow = d.get("is_grow", False)
        has_grow_adopt = is_grow and d.get("grow_adopt", 0) > 0
        num_rows = max(1, len(tasks) + (1 if has_grow_adopt else 0)) if is_grow else max(1, len(tasks))

        for i in range(num_rows):
            is_first = (i == 0)
            is_grow_adopt_row = is_grow and has_grow_adopt and i == 1
            task = tasks[i] if 0 <= i < len(tasks) else None

            # A-C: 基础信息
            if is_first:
                set_cell(ws, row, 1, d["activity"])
                set_cell(ws, row, 2, d["period"])
                set_cell(ws, row, 3, d["uv"])
            else:
                for c in range(1, 4):
                    set_cell(ws, row, c)

            # D-E: 点击主按钮 + 参与率
            if is_first:
                click_val = d["grow_click"] if is_grow else d["click"]
                set_cell(ws, row, 4, click_val if click_val > 0 else SLASH)
                set_cell(ws, row, 5, f'=IF(OR(C{first_row}=0,D{first_row}="/"),"/",D{first_row}/C{first_row})', pct)
            elif is_grow_adopt_row:
                set_cell(ws, row, 4, d["grow_adopt"])
                set_cell(ws, row, 5, f'=IF(OR(C{first_row}=0,D{row}=0),"/",D{row}/C{first_row})', pct)
            else:
                set_cell(ws, row, 4)
                set_cell(ws, row, 5)

            # F-I: 任务
            if task:
                set_cell(ws, row, 6, task["task_name"])
                set_cell(ws, row, 7, task["user_num"] if task["user_num"] > 0 else SLASH)
                set_cell(ws, row, 8, f'=IF(OR($C${first_row}=0,G{row}="/"),"/",G{row}/$C${first_row})', pct)
                set_cell(ws, row, 9, task["num"] if task["num"] > 0 else SLASH)
            else:
                set_cell(ws, row, 6, SLASH)
                set_cell(ws, row, 7, SLASH)
                set_cell(ws, row, 8)
                set_cell(ws, row, 9, SLASH)

            # J-K: 留存
            if is_first:
                set_cell(ws, row, 10, d["day1"] if d["day1"] > 0 else SLASH, pct)
                set_cell(ws, row, 11, d["day7"] if d["day7"] > 0 else SLASH, pct)
            else:
                set_cell(ws, row, 10)
                set_cell(ws, row, 11)

            # L-M: 金币回收（可选）
            if has_coin:
                if is_first and d.get("is_coin"):
                    set_cell(ws, row, coin_col_start, d["coin_recovery"] if d["coin_recovery"] > 0 else SLASH)
                    ratio = d["coin_recovery"] / d["coin_user_total"] if d["coin_user_total"] > 0 else 0
                    set_cell(ws, row, coin_col_start + 1, ratio if ratio > 0 else SLASH, pct)
                else:
                    set_cell(ws, row, coin_col_start)
                    set_cell(ws, row, coin_col_start + 1)

            row += 1

    # 列宽
    widths = [14, 24, 8, 20, 10, 40, 12, 10, 12, 10, 10]
    if has_coin:
        widths.extend([14, 12])
    for i, w in enumerate(widths, 1):
        col_letter = chr(64 + i) if i <= 26 else "A" + chr(64 + i - 26)
        ws.column_dimensions[col_letter].width = w

    wb.save(output_file)


# ============ 主流程 ============

def main():
    parser = argparse.ArgumentParser(description="Turbolink 活动数据拉取")
    parser.add_argument("--token", required=True, help="Bearer token (从 Turbolink 后台获取)")
    parser.add_argument("--project-id", required=True, help="项目 ID (Turbolink 后台 URL 中的 pjid)")
    parser.add_argument("--uv-threshold", type=int, default=15, help="UV 过滤阈值 (默认 15)")
    parser.add_argument("--coin-cutoff", default=None, help="金币回收统计截止日期，如 2026-06-11")
    parser.add_argument("--output", default="活动数据分析.xlsx", help="输出文件名")
    parser.add_argument("--search-start", default="2025/01/01 00:00", help="活动搜索起始时间")
    parser.add_argument("--search-end", default="2099/12/31 23:59", help="活动搜索结束时间")
    args = parser.parse_args()

    # 检查 token 是否过期
    exp = parse_token_expiry(args.token)
    if exp and exp < datetime.now():
        print(f"错误: Token 已过期 ({exp.strftime('%Y-%m-%d %H:%M')})，请从 Turbolink 后台获取新 Token")
        sys.exit(1)
    elif exp:
        print(f"Token 有效至: {exp.strftime('%Y-%m-%d %H:%M')}")

    headers = build_headers(args.token, args.project_id)

    print(f"活动搜索范围: {args.search_start} ~ {args.search_end}")

    print("正在获取活动类型映射...")
    type_map = fetch_fission_type_map(headers)
    print(f"  共 {len(type_map)} 种活动类型")

    print("正在拉取活动列表...")
    campaigns = fetch_campaign_list(headers, args.project_id, args.search_start, args.search_end)
    print(f"共获取 {len(campaigns)} 个活动")

    activities = filter_campaigns(campaigns, type_map, args.uv_threshold)
    print(f"\n符合条件的活动: {len(activities)} 个")

    if not activities:
        print("没有符合条件的活动，退出")
        return

    has_coin = any(a["fission_mark"] == "coin" for a in activities)

    print("\n开始拉取详细数据...")
    all_data = []
    for act in activities:
        print(f"\n拉取: {act['activity']} ({act['start']} ~ {act['end']})")
        data = fetch_activity(headers, act, args.coin_cutoff)
        all_data.append(data)
        print(f"  UV: {data['uv']} | 点击: {data['click']} | 任务: {len(data['tasks'])} 个")

    build_excel(all_data, args.output, has_coin)
    print(f"\nExcel 已保存: {args.output} ({len(all_data)} 个活动)")


if __name__ == "__main__":
    main()
