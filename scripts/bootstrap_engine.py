#!/usr/bin/env python3
"""定位 PPTX 引擎；必要时从本 skill 内置的归档自举解压。

用法::

    python scripts/bootstrap_engine.py                 # 自举（如需）并打印引擎目录
    python scripts/bootstrap_engine.py --check         # 只报告状态，不写任何文件
    python scripts/bootstrap_engine.py --json          # 机器可读输出
    python scripts/bootstrap_engine.py --cache-dir DIR # 指定解压缓存的基础目录
    python scripts/bootstrap_engine.py --force         # 忽略已有缓存，重新解压
    python scripts/bootstrap_engine.py --verify        # 解压后额外运行引擎版权守卫

定位顺序：

1. 环境变量 ``PPT_MASTER_DIR`` 指向有效引擎目录时直接使用，不做任何解压；
2. 否则解析缓存目录，优先级为 ``--cache-dir`` > ``PPT_MASTER_CACHE_DIR`` >
   skill 同级 ``_engine/`` > 当前目录 ``_engine/``；
3. 缓存目录已含**完整**引擎时复用；不完整时重新解压（解压可覆盖写入，因此可直接
   修复被中断的缓存）；否则从 ``vendor/`` 归档解压。解压前校验归档 SHA-256，
   解压后校验文件数与 LICENSE 摘要是否与清单一致。

完好性判定同时看四件事：``SKILL.md`` 与 ``LICENSE`` 存在、LICENSE 摘要与清单一致、
``workflows/`` ``templates/`` ``scripts/`` ``references/`` 四项结构存在且
``scripts/attribution_guard.py`` 就位、文件总数等于清单的 ``file_count``。只按
前两项判定会把被中断的残缺解压当成有效引擎，因此必须带上文件数校验。

安全边界：只读取本 skill 的 ``vendor/`` 目录；只写入缓存目录；不删除任何文件；
候选缓存目录全部落在系统盘时拒绝执行。逐条解压并在瞬时占用时重试，避免单条
写入失败丢掉整次解压。引擎目录内的文件不得修改。
"""
from __future__ import annotations

import argparse
import hashlib
import importlib.util
import json
import os
import sys
import time
import zipfile
from pathlib import Path

SKILL_DIR = Path(__file__).resolve().parent.parent
VENDOR_DIR = SKILL_DIR / "vendor"
ARCHIVE = VENDOR_DIR / "ppt-master-6.4.0.zip"
MANIFEST = VENDOR_DIR / "ppt-master-6.4.0.manifest.json"
ENGINE_DIR_NAME = "ppt-master-6.4.0"


def load_manifest() -> dict:
    if not MANIFEST.is_file():
        return {}
    return json.loads(MANIFEST.read_text(encoding="utf-8"))


def normalized_digest(path: Path) -> str:
    """LICENSE 摘要按行尾归一化后计算，与引擎守卫口径一致。"""
    data = path.read_bytes()
    if data.startswith(b"\xef\xbb\xbf"):
        return ""
    text = data.decode("utf-8")
    return hashlib.sha256(text.replace("\r\n", "\n").replace("\r", "\n").encode("utf-8")).hexdigest()


REQUIRED_STRUCTURE = (
    "workflows",
    "templates",
    "references",
    "scripts",
)
REQUIRED_FILES = (
    "SKILL.md",
    "LICENSE",
    "scripts/attribution_guard.py",
    "workflows/generate-pptx.md",
    "scripts/project_manager.py",
)


def count_files(path: Path) -> int:
    return sum(1 for item in path.rglob("*") if item.is_file())


def engine_is_valid(path: Path, expected_licence: str, expected_count: int = 0) -> tuple[bool, str]:
    """返回 (是否可用, 不可用原因)。

    只检查 ``SKILL.md`` 与 LICENSE 会把被中断的残缺解压误判为有效引擎，
    因此同时校验关键结构、关键文件与文件总数。
    """
    if not path.is_dir():
        return False, "目录不存在"
    for rel in REQUIRED_STRUCTURE:
        if not (path / rel).is_dir():
            return False, f"缺少目录 {rel}/"
    for rel in REQUIRED_FILES:
        if not (path / rel).is_file():
            return False, f"缺少文件 {rel}"
    licence = path / "LICENSE"
    if expected_licence and normalized_digest(licence) != expected_licence:
        return False, "LICENSE 摘要与清单不一致"
    if expected_count:
        actual = count_files(path)
        if actual != expected_count:
            return False, f"文件数不符（实际 {actual}，清单 {expected_count}）"
    return True, ""


def on_system_drive(path: Path) -> bool:
    return os.path.splitdrive(str(path.resolve()))[0].upper() == "C:"


def resolve_cache_dir(args: argparse.Namespace) -> Path:
    bases = []
    if args.cache_dir:
        bases.append(Path(args.cache_dir))
    env_dir = os.environ.get("PPT_MASTER_CACHE_DIR", "").strip()
    if env_dir:
        bases.append(Path(env_dir))
    bases.append(SKILL_DIR.parent / "_engine")
    bases.append(Path.cwd() / "_engine")

    rejected = []
    for base in bases:
        target = base / ENGINE_DIR_NAME
        if on_system_drive(target):
            rejected.append(str(target))
            continue
        return target
    raise SystemExit(
        "缓存目录候选位置全部落在系统盘，请用 --cache-dir 或 PPT_MASTER_CACHE_DIR 指定可写目录。"
        f"已排除：{'; '.join(rejected)}"
    )


def extract_archive(archive: Path, target: Path) -> int:
    """逐条校验条目不越出目标目录，再逐条写入；单条瞬时占用时重试。

    Windows 上杀毒软件或索引服务会短暂锁住刚写入的文件，``extractall`` 遇到
    这种瞬时 ``PermissionError`` 会整体中断并留下残缺缓存。这里改为逐条写入
    并在失败时重试若干次，仍失败才中止，避免因一次占用丢掉整次解压。
    """
    target.mkdir(parents=True, exist_ok=True)
    root = target.resolve()
    count = 0
    with zipfile.ZipFile(archive) as handle:
        for member in handle.infolist():
            if member.is_dir():
                continue
            write_member(handle, member, root)
            count += 1
    return count


def write_member(handle: zipfile.ZipFile, member: zipfile.ZipInfo, root: Path) -> None:
    destination = (root / member.filename).resolve()
    try:
        destination.relative_to(root)
    except ValueError:
        raise SystemExit(f"归档条目越出目标目录，已中止：{member.filename}")
    destination.parent.mkdir(parents=True, exist_ok=True)
    last = None
    for attempt in range(3):
        try:
            with handle.open(member) as source, open(destination, "wb") as sink:
                while True:
                    chunk = source.read(1 << 20)
                    if not chunk:
                        break
                    sink.write(chunk)
            return
        except (PermissionError, OSError) as exc:  # 杀软或索引服务瞬时占用
            last = exc
            time.sleep(0.2 * (attempt + 1))
    raise SystemExit(
        f"写入失败（{last}）：{member.filename}。"
        "缓存目录可能被占用或不允许写入，请重试，或换用 --cache-dir 指定其他位置。"
    )


def run_guard(engine: Path) -> None:
    """在进程内加载引擎守卫并执行完整性检查。

    直接调用守卫的 ``require_skill_integrity()``，不启动子进程；守卫按
    ``Path(__file__).parent.parent`` 自定位引擎根，因此加载即指向目标引擎。
    """
    guard = engine / "scripts" / "attribution_guard.py"
    if not guard.is_file():
        raise SystemExit("引擎缺少 scripts/attribution_guard.py，无法校验完整性")
    spec = importlib.util.spec_from_file_location("_ppt_master_attribution_guard", guard)
    module = importlib.util.module_from_spec(spec)
    try:
        spec.loader.exec_module(module)
        module.require_skill_integrity()
    except SystemExit as exc:
        raise SystemExit(f"引擎完整性守卫未通过：{exc.code}")
    except Exception as exc:
        raise SystemExit(f"引擎完整性守卫未通过：{exc}")
    print("[OK] 引擎完整性守卫通过", file=sys.stderr)


def emit(result: dict, args: argparse.Namespace) -> int:
    if args.json:
        print(json.dumps(result, ensure_ascii=False, indent=2))
    else:
        print(f"引擎目录: {result['engine']}")
        print(f"来源: {result['source']}｜动作: {result['action']}")
        if result.get("reason"):
            print(f"说明: {result['reason']}")
    return 0


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="定位或自举 ppt-master 引擎")
    parser.add_argument("--check", action="store_true", help="只报告状态，不写任何文件")
    parser.add_argument("--json", action="store_true", help="输出 JSON")
    parser.add_argument("--cache-dir", help="解压缓存的基础目录")
    parser.add_argument("--force", action="store_true", help="忽略已有缓存并重新解压")
    parser.add_argument("--verify", action="store_true", help="解压后运行引擎版权守卫")
    args = parser.parse_args(argv)

    manifest = load_manifest()
    expected_licence = manifest.get("licence_sha256", "")
    expected_count = int(manifest.get("file_count") or 0)
    result = {"engine": None, "source": None, "action": None, "file_count": None, "reason": ""}

    env_dir = os.environ.get("PPT_MASTER_DIR", "").strip()
    if env_dir:
        candidate = Path(env_dir)
        ok, reason = engine_is_valid(candidate, expected_licence)
        if ok:
            result.update(engine=str(candidate.resolve()), source="env:PPT_MASTER_DIR", action="none")
            return emit(result, args)
        if not args.json:
            print(f"[提示] PPT_MASTER_DIR={env_dir} 不是有效引擎目录（{reason}），继续尝试内置归档。", file=sys.stderr)

    if not ARCHIVE.is_file():
        raise SystemExit(f"内置归档缺失：{ARCHIVE.relative_to(SKILL_DIR).as_posix()}")

    cache = resolve_cache_dir(args)
    cache_exists = cache.exists()
    ok, reason = engine_is_valid(cache, expected_licence, expected_count)
    if ok and not args.force:
        result.update(engine=str(cache), source="cache", action="reuse")
        if args.verify:
            run_guard(cache)
        return emit(result, args)

    if args.check:
        result.update(
            engine=str(cache),
            source="none",
            action="needs_repair" if cache_exists else "needs_extract",
            reason=reason,
        )
        return emit(result, args)

    archive_digest = hashlib.sha256(ARCHIVE.read_bytes()).hexdigest()
    expected_archive = manifest.get("archive_sha256", "")
    if expected_archive and archive_digest != expected_archive:
        raise SystemExit("归档 SHA-256 与清单不符，已中止解压")

    count = extract_archive(ARCHIVE, cache)
    ok, reason = engine_is_valid(cache, expected_licence, expected_count)
    if not ok:
        raise SystemExit(f"解压后校验未通过（{reason}），已中止；请检查 vendor/ 下的归档与清单")

    result.update(
        engine=str(cache),
        source="archive",
        action="repair" if cache_exists else "extract",
        file_count=count,
    )
    if args.verify:
        run_guard(cache)
    return emit(result, args)


if __name__ == "__main__":
    raise SystemExit(main())
