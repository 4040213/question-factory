"""抓取莫顿一道正常题目的原始 JSON，用于和我们生成的做结构对比。

用法（不写死凭证，避免泄露）：
  set MORTON_COOKIE=staffId=...tokenCrm=...
  set MORTON_TOKEN=Bearer ...
  python fetch_morton_question.py --id 25867

可选参数：
  --id        题目ID（默认 25867，即已知的正常题）
  --host      莫顿域名（默认 https://tms-mx.xueqiulearning.com）
  --out       输出文件路径（默认 morton_question_<id>.json）

脚本会依次试探几组常见接口，第一个返回 200 的就被 dump 出来。
"""
import os
import sys
import json
import argparse

try:
    import httpx
except ImportError:
    sys.exit("缺少 httpx，请先装：pip install httpx")

API_HOST = "https://tms-mx.xueqiulearning.com"


def _headers(cookie: str, token: str) -> dict:
    return {
        "Cookie": cookie,
        "Authorization": token,
        "Content-Type": "application/json",
        "Referer": f"{API_HOST}/fe_morton/morton/",
    }


def _candidates(qid: str):
    """返回待试探的 (method, path, json_body, params) 列表。"""
    return [
        ("POST", "/merton-backend/question/detail", {"id": qid}, None),
        ("POST", "/merton-backend/question/detail", {"questionId": qid}, None),
        ("GET",  f"/merton-backend/question/detail?id={qid}", None, None),
        ("POST", "/merton-backend/question/getById", {"id": qid}, None),
        ("POST", "/merton-backend/question/page", {"id": qid, "pageNum": 1, "pageSize": 10}, None),
        ("POST", "/merton-backend/question/info", {"id": qid}, None),
    ]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--id", default="25867")
    ap.add_argument("--host", default=API_HOST)
    ap.add_argument("--out", default=None)
    args = ap.parse_args()

    cookie = os.environ.get("MORTON_COOKIE", "")
    token = os.environ.get("MORTON_TOKEN", "")
    if not cookie or not token:
        sys.exit("请先设置环境变量 MORTON_COOKIE 和 MORTON_TOKEN")

    headers = _headers(cookie, token)
    out_path = args.out or f"morton_question_{args.id}.json"

    with httpx.Client(base_url=args.host, headers=headers, timeout=30) as cli:
        for method, path, body, params in _candidates(args.id):
            try:
                if method == "POST":
                    r = cli.post(path, json=body)
                else:
                    r = cli.get(path)
                print(f"[{method}] {path} -> HTTP {r.status_code}")
                if r.status_code == 200:
                    try:
                        data = r.json()
                    except Exception:
                        print("  响应不是 JSON，前 200 字符：")
                        print(r.text[:200])
                        continue
                    # 尝试定位 questionData
                    qd = None
                    if isinstance(data, dict):
                        qd = data.get("questionData") or data.get("data") or data.get("result")
                    with open(out_path, "w", encoding="utf-8") as f:
                        json.dump(data, f, ensure_ascii=False, indent=2)
                    print(f"  ✓ 已保存原始响应到: {out_path}")
                    if qd is not None:
                        print("  ✓ 找到 questionData 字段")
                    else:
                        print("  ⚠ 未直接找到 questionData，已保存整包响应，人工查看")
                    return
                else:
                    print(f"  响应片段: {r.text[:160]}")
            except Exception as e:
                print(f"  ✗ 异常: {e}")

    print("\n所有候选接口都失败了。请用方法一(F12)确认莫顿前端实际调用的接口路径和参数，发我后我立刻修正脚本。")


if __name__ == "__main__":
    main()
