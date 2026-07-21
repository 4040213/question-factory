"""莫顿题目 JSON 逐字段对比工具。

用法:
    python diff_question.py <working.json> <ours.json>

两个文件各自包含一道题目的完整 questionData（JSON 对象，或 {"questionData": {...}} 包裹）。
脚本会递归对比，输出:
    - 只在 A(正常题) 有的字段
    - 只在 B(我们生成的) 有的字段
    - 两边都有但值/类型不同的字段
    - 顶层缺失的关键字段告警
"""
import json
import sys
from typing import Any


def _unwrap(obj: Any) -> Any:
    """兼容 {"questionData": {...}} 或直接是 {...} 两种情况。"""
    if isinstance(obj, dict) and "questionData" in obj and isinstance(obj["questionData"], dict):
        return obj["questionData"]
    return obj


def _diff(path: str, a: Any, b: Any, only_a: list, only_b: list, diffs: list):
    # 类型不同
    if type(a) is not type(b):
        diffs.append(f"[类型不同] {path}: A={type(a).__name__}({_short(a)})  B={type(b).__name__}({_short(b)})")
        return
    if isinstance(a, dict):
        keys_a = set(a.keys())
        keys_b = set(b.keys())
        for k in sorted(keys_a - keys_b):
            only_a.append(f"{path}.{k}")
        for k in sorted(keys_b - keys_a):
            only_b.append(f"{path}.{k}")
        for k in sorted(keys_a & keys_b):
            _diff(f"{path}.{k}", a[k], b[k], only_a, only_b, diffs)
    elif isinstance(a, list):
        if len(a) != len(b):
            diffs.append(f"[数组长度不同] {path}: A={len(a)}  B={len(b)}")
        for i in range(min(len(a), len(b))):
            _diff(f"{path}[{i}]", a[i], b[i], only_a, only_b, diffs)
    else:
        if a != b:
            diffs.append(f"[值不同] {path}: A={_short(a)}  B={_short(b)}")


def _short(v: Any, n: int = 80) -> str:
    s = json.dumps(v, ensure_ascii=False) if not isinstance(v, str) else v
    return s if len(s) <= n else s[:n] + "..."


def main():
    if len(sys.argv) != 3:
        print("用法: python diff_question.py <working.json> <ours.json>")
        sys.exit(1)
    with open(sys.argv[1], "r", encoding="utf-8") as f:
        A = _unwrap(json.load(f))
    with open(sys.argv[2], "r", encoding="utf-8") as f:
        B = _unwrap(json.load(f))

    only_a, only_b, diffs = [], [], []
    _diff("questionData", A, B, only_a, only_b, diffs)

    print("=" * 60)
    print(" 只在【正常题】里有、我们缺失的字段:")
    print("=" * 60)
    if only_a:
        for x in only_a:
            print("  -", x)
    else:
        print("  (无)")

    print("=" * 60)
    print(" 只在我们生成的题里有、正常题没有的字段:")
    print("=" * 60)
    if only_b:
        for x in only_b:
            print("  +", x)
    else:
        print("  (无)")

    print("=" * 60)
    print(" 两边都有但值/类型不同的字段:")
    print("=" * 60)
    if diffs:
        for d in diffs:
            print("  ", d)
    else:
        print("  (无)")

    # 顶层关键字段告警
    print("=" * 60)
    print(" 顶层关键字段速览:")
    print("=" * 60)
    for key in ["questionType", "subjectId", "resolution", "skinVersion",
                "displayBox", "newDisplayBox", "webBgmUrl", "name"]:
        av = A.get(key, "<缺失>")
        bv = B.get(key, "<缺失>")
        flag = "" if av == bv else "  ⚠️不同"
        print(f"  {key}: 正常={_short(av,30)} | 我们={_short(bv,30)}{flag}")
    qi_a = A.get("questionItem", {})
    qi_b = B.get("questionItem", {})
    for key in ["displayBox", "newDisplayBox", "displayNumber"]:
        av = qi_a.get(key, "<缺失>")
        bv = qi_b.get(key, "<缺失>")
        flag = "" if av == bv else "  ⚠️不同"
        print(f"  questionItem.{key}: 正常={_short(av,30)} | 我们={_short(bv,30)}{flag}")


if __name__ == "__main__":
    main()
