#!/usr/bin/env python3
"""统计每个文件中由指定作者最后修改的存活行数（基于 git blame）。

用法:
    python scripts/blame_stats.py                      # 默认作者列表
    python scripts/blame_stats.py --author Muu --author Muuyo
    python scripts/blame_stats.py --author "Alice" --author "Bob"
    python scripts/blame_stats.py -A Alice -A Bob
    python scripts/blame_stats.py --repo /path/to/repo -A Muu
    python scripts/blame_stats.py --top 20
    python scripts/blame_stats.py --output report.txt
"""

import concurrent.futures
import io
import os
import subprocess
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

_REPO_ROOT = Path(__file__).resolve().parent.parent.parent
if str(_REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(_REPO_ROOT))

from utils.argutils import parse_args as argutils_parse_args  # noqa: E402


@dataclass
class BlameStatsConfig:
    authors: Optional[list[str]] = field(
        default=None,
        metadata={
            "meta": {
                "arg_names": ["-A", "--author"],
                "action": "append",
                "dest": "authors",
                "type": str,
                "nargs": None,
                "default": None,
            }
        },
    )
    """要统计的作者名，可重复指定多个（默认: Muu, Muuyo）。"""

    repo: str = "."
    """git 仓库路径（默认: 当前目录）。"""

    top: int = 0
    """仅输出匹配行数最多的前 N 个文件（0 表示全部）。"""

    min_lines: int = 1
    """过滤掉匹配行数小于该值的文件（默认: 1）。"""

    output: Optional[Path] = None
    """将结果写入指定文件（同时在终端打印）。"""

    include_blank: bool = False
    """统计时包含空行（默认排除空行）。"""

    workers: int = field(
        default_factory=lambda: max(1, (os.cpu_count() or 4)),
    )
    """并行 blame 的进程数（默认: CPU 核数）。"""


@dataclass
class FileStats:
    path: str
    matched: int
    total: int

    @property
    def ratio(self) -> float:
        return self.matched / self.total if self.total else 0.0


def run_git(repo: Path, *args: str) -> str:
    result = subprocess.run(
        ["git", "-C", str(repo)] + list(args),
        capture_output=True,
        check=False,
    )
    return result.stdout.decode("utf-8", errors="replace")


def list_tracked_files(repo: Path) -> list[str]:
    out = run_git(repo, "ls-files", "-z")
    return [f for f in out.split("\x00") if f]


def blame_file(repo: Path, file: str, authors: set[str], include_blank: bool = False) -> FileStats:
    proc = subprocess.run(
        ["git", "-C", str(repo), "blame", "--line-porcelain", "--", file],
        capture_output=True,
        check=False,
    )
    if proc.returncode != 0 or not proc.stdout:
        return FileStats(path=file, matched=0, total=0)

    text = proc.stdout.decode("utf-8", errors="replace")
    matched = 0
    total = 0
    current_author: str | None = None
    for line in text.splitlines():
        if line.startswith("author "):
            current_author = line[7:]
        elif line.startswith("\t") and current_author is not None:
            content = line[1:]
            if include_blank or content.strip():
                total += 1
                if current_author in authors:
                    matched += 1
            current_author = None
    return FileStats(path=file, matched=matched, total=total)


def collect_stats(
    repo: Path, files: list[str], authors: set[str], workers: int, include_blank: bool = False
) -> list[FileStats]:
    stats: list[FileStats] = []
    with concurrent.futures.ThreadPoolExecutor(max_workers=workers) as pool:
        futures = {pool.submit(blame_file, repo, f, authors, include_blank): f for f in files}
        for fut in concurrent.futures.as_completed(futures):
            try:
                stats.append(fut.result())
            except Exception as exc:  # noqa: BLE001
                print(f"[warn] blame {futures[fut]} failed: {exc}", file=sys.stderr)
    return stats


def render(
    stats: list[FileStats],
    authors: list[str],
    top: int,
    min_lines: int,
) -> str:
    filtered = [s for s in stats if s.matched >= min_lines]
    filtered.sort(key=lambda s: (-s.matched, s.path))

    if top > 0:
        filtered = filtered[:top]

    matched_total = sum(s.matched for s in stats)
    line_total = sum(s.total for s in stats)
    ratio = (matched_total / line_total * 100) if line_total else 0.0

    max_path_len = max((len(s.path) for s in filtered), default=10)
    path_width = ((max_path_len + 4) // 5) * 5 + 5

    lines: list[str] = [
        f"作者: {', '.join(authors)}",
        f"匹配文件数    : {len(filtered)}（共 {len(stats)} 个跟踪文件）",
        f"存活行数      : {matched_total}",
        f"项目总行数    : {line_total}",
        f"占比          : {ratio:.2f}%",
        "",
    ]
    header = f"{'File':<{path_width}} {'Match':>8} {'Total':>8} {'Ratio':>8}"
    lines.append(header)
    lines.append("-" * len(header))
    for s in filtered:
        pct = f"{s.ratio * 100:.2f}%" if s.total else "-"
        lines.append(f"{s.path:<{path_width}} {s.matched:>8} {s.total:>8} {pct:>8}")
    return "\n".join(lines)


def main() -> int:
    if sys.platform == "win32":
        for stream in (sys.stdout, sys.stderr):
            try:
                stream.reconfigure(encoding="utf-8")  # type: ignore[attr-defined]
            except (AttributeError, io.UnsupportedOperation):
                pass

    args = argutils_parse_args(BlameStatsConfig, "Git Blame Stats")
    repo = Path(args.repo).resolve()
    if not (repo / ".git").exists():
        print(f"error: {repo} 不是 git 仓库", file=sys.stderr)
        return 1

    authors = args.authors or ["Muu", "Muuyo"]
    authors_set = set(authors)

    files = list_tracked_files(repo)
    if not files:
        print("error: 仓库中没有跟踪文件", file=sys.stderr)
        return 1

    stats = collect_stats(repo, files, authors_set, args.workers, args.include_blank)
    output = render(stats, authors, args.top, args.min_lines)

    print(output)
    if args.output is not None:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(output + "\n", encoding="utf-8")
    return 0


if __name__ == "__main__":
    sys.exit(main())
