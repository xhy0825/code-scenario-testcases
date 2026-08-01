#!/usr/bin/env python3
"""确定性判定点枚举脚本（多语言，正则启发式）。

用途：作为覆盖率的分母基准，把源码中的判定点（if / else if / switch / case /
default / 三目 / catch）按语言语法枚举出来，输出 JSON。与 SKILL.md 第 4 步配套：

    python scripts/enumerate_decision_points.py <目录或文件> [--output dp.json]

输出结构：
    {
      "tool": "enumerate_decision_points",
      "scanned_files": 2,
      "total_decision_points": 5,
      "files": [
        {
          "path": "src/.../OrderService.java",
          "language": "java",
          "decision_points": [
            {"line": 3, "type": "if", "text": "if (req.getUserId() == null)",
             "compound": false, "boundary": true}
          ]
        }
      ]
    }

说明：
- 正则启发式，追求"确定性 + 显式列出"，而非完美解析；注释行已剔除。
- `compound` = 条件行含 `&&` / `||`（提示需按 SKILL.md A1 拆子条件）。
- `boundary` = 条件行含比较运算符（提示需按 SKILL.md A2 补边界用例）。
- 不把 `for` / `while` 循环本身计为判定点（SKILL 的判定点模型把"循环内 if"计为判定）。
"""

import argparse
import json
import os
import re
import sys
from datetime import date

# 扩展名 → 语言
LANG_BY_EXT = {
    ".java": "java", ".kt": "kotlin", ".kts": "kotlin",
    ".go": "go", ".py": "python",
    ".js": "js", ".mjs": "js", ".cjs": "js", ".ts": "ts", ".tsx": "tsx",
    ".cs": "cs", ".c": "c", ".h": "c", ".cpp": "cpp", ".hpp": "cpp", ".cc": "cpp",
    ".rs": "rust", ".php": "php", ".rb": "ruby", ".swift": "swift",
}

# 每种判定点类型的识别正则（按语言归类）
# type, regex, languages
PATTERNS = [
    ("if", r"\bif\s*\(", ["java", "kotlin", "js", "ts", "tsx", "cs", "c", "cpp", "php", "swift", "ruby"]),
    ("if", r"\bif\s+", ["go", "rust", "python"]),
    ("elif", r"\belse\s+if\s*\(", ["java", "kotlin", "js", "ts", "tsx", "cs", "c", "cpp", "php", "swift", "ruby"]),
    ("elif", r"\belse\s+if\s+", ["go", "rust"]),
    ("elif", r"^\s*elif\s+", ["python"]),
    ("switch", r"\bswitch\s*\(", ["java", "kotlin", "js", "ts", "tsx", "cs", "c", "cpp", "php", "swift", "ruby"]),
    ("switch", r"\bswitch\s", ["go"]),
    ("match", r"\bmatch\s", ["python", "rust"]),
    ("case", r"\bcase\s+[^:{}]*:", ["java", "kotlin", "js", "ts", "tsx", "cs", "c", "cpp", "php", "swift", "ruby"]),
    ("case", r"\bcase\s+", ["go"]),
    ("default", r"\bdefault\s*:", ["java", "kotlin", "js", "ts", "tsx", "cs", "c", "cpp", "php", "swift", "go"]),
    ("catch", r"\bcatch\s*\(", ["java", "kotlin", "js", "ts", "tsx", "cs", "c", "cpp", "php", "swift"]),
    ("except", r"^\s*except\s+", ["python"]),
    # 三目运算符（保守匹配，可能误报类型参数/对象字面量，供 LLM 复核）
    ("ternary", r"\?\s*[^?=:]+:", ["java", "kotlin", "cs", "c", "cpp", "js", "ts", "tsx", "php", "ruby"]),
]

COMPOUND_RE = re.compile(r"&&|\|\|")
BOUNDARY_RE = re.compile(r">=|<=|===|!==|==|!=|<>|>|<")
# 参数约束注解（Java Bean Validation / Jakarta 为主，其它语言映射见 references/input-constraints.md）
CONSTRAINT_RE = re.compile(r"@(NotNull|NotBlank|NotEmpty|Min|Max|Size|Pattern|Email|Positive|Negative|DecimalMin|DecimalMax|Range|Length|MinLength|MaxLength)\b")
# 剔除箭头函数 -> 对 boundary 的干扰
ARROW_RE = re.compile(r"->|\w+\s*:\s*\w+\s*=>")


def strip_line_comment(line: str, lang: str) -> str:
    """剔除行注释，返回参与匹配的代码片段。"""
    if lang in ("python", "ruby"):
        idx = line.find("#")
        return line[:idx] if idx >= 0 else line
    # C 系 / 类 C 系：//
    idx = line.find("//")
    return line[:idx] if idx >= 0 else line


def classify(lang: str, cleaned: str, line_no: int):
    """在 cleaned 行上匹配判定点，返回 (type, text, compound, boundary) 或 None。"""
    for dp_type, regex, langs in PATTERNS:
        if lang not in langs:
            continue
        m = re.search(regex, cleaned)
        if m:
            # 文本：取该行自匹配起点到行尾前 60 字符
            text = cleaned[m.start():m.start() + 60].strip()
            # compound 标记：Python/Ruby 用 and/or，其余语言用 &&/||
            if lang in ("python", "ruby"):
                compound = bool(re.search(r"\band\b|\bor\b", cleaned))
            else:
                compound = bool(COMPOUND_RE.search(cleaned))
            boundary = False
            if dp_type in ("if", "elif", "switch", "match", "ternary"):
                stripped = ARROW_RE.sub("", cleaned)
                boundary = bool(BOUNDARY_RE.search(stripped))
            return dp_type, text, compound, boundary
    return None


def scan_file(path: str):
    lang = LANG_BY_EXT.get(os.path.splitext(path)[1].lower())
    if not lang:
        return None
    points = []
    constraints = []
    try:
        with open(path, encoding="utf-8", errors="replace") as f:
            for line_no, raw in enumerate(f, start=1):
                cleaned = strip_line_comment(raw, lang)
                if not cleaned.strip():
                    continue
                hit = classify(lang, cleaned, line_no)
                if hit:
                    dp_type, text, compound, boundary = hit
                    points.append({
                        "line": line_no,
                        "type": dp_type,
                        "text": text,
                        "compound": compound,
                        "boundary": boundary,
                    })
                for cm in CONSTRAINT_RE.finditer(cleaned):
                    constraints.append({
                        "line": line_no,
                        "type": "constraint",
                        "annotation": cm.group(1),
                        "text": cleaned.strip()[:60],
                    })
    except OSError as e:
        print(f"warning: 无法读取 {path}: {e}", file=sys.stderr)
        return None
    return {"path": path.replace("\\", "/"), "language": lang,
            "decision_points": points, "constraints": constraints}


def collect_files(path):
    if os.path.isfile(path):
        yield path
        return
    for root, dirs, files in os.walk(path):
        # 跳过构建产物/依赖目录
        dirs[:] = [d for d in dirs if d not in
                   ("target", "build", "dist", "bin", "out", "node_modules",
                    "vendor", ".venv", "site-packages", ".git", ".svn", "__pycache__")]
        for f in sorted(files):
            if f.endswith((".pyc",)):
                continue
            yield os.path.join(root, f)


def main():
    parser = argparse.ArgumentParser(description="确定性判定点枚举（覆盖率分母基准）")
    parser.add_argument("path", help="源码目录或单个文件")
    parser.add_argument("--output", "-o", help="输出 JSON 路径（默认 stdout）")
    args = parser.parse_args()

    files = []
    total = 0
    for fp in collect_files(args.path):
        if not os.path.splitext(fp)[1].lower() in LANG_BY_EXT:
            continue
        info = scan_file(fp)
        if info:
            files.append(info)
            total += len(info["decision_points"])

    const_total = sum(len(f.get("constraints") or []) for f in files)
    result = {
        "tool": "enumerate_decision_points",
        "generated": date.today().isoformat(),
        "scanned_files": len(files),
        "total_decision_points": total,
        "total_input_constraints": const_total,
        "files": files,
    }
    out = json.dumps(result, ensure_ascii=False, indent=2)
    if args.output:
        with open(args.output, "w", encoding="utf-8") as f:
            f.write(out)
        print(f"已枚举 {total} 个判定点 -> {args.output}")
    else:
        print(out)


if __name__ == "__main__":
    main()
