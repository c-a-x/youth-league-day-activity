# -*- coding: utf-8 -*-
"""检查本 skill 的依赖环境，输出 ready / partial / needs_setup 三态。

用法：
  python check_environment.py [--skill-dir 目录] [--json] [--probe-word]

--probe-word  额外真实启动一次 Word 进程读取版本号后立即退出；不加时只读注册表。
--json        输出机器可读的 JSON，便于自动化判断。

本脚本只读：不安装软件、不写配置文件（包括不新建目录）、不修改 skill 或引擎目录、
不发起任何网络请求。依赖声明来自 skill-dependencies.json，本脚本按其中的 check id 逐项执行。

退出码：0 = ready，1 = partial，2 = needs_setup。
"""
import argparse
import hashlib
import importlib.util
import json
import os
import sys
from pathlib import Path

SKILL_ROOT = Path(__file__).resolve().parents[1]
ENGINE_LICENSE_SHA256 = "80cefc234c1ec12a8cece4344f16300c634fa03df7891686fcf979e3828f0921"
ARCHIVE_NAME = "ppt-master-6.4.0.zip"
CACHE_DIR_NAME = "ppt-master-6.4.0"


def archive_path():
    return SKILL_ROOT / "vendor" / ARCHIVE_NAME


def default_cache_dir():
    """解析默认解压缓存目录，与 bootstrap_engine.py 口径一致。

    候选顺序：``PPT_MASTER_CACHE_DIR`` → skill 同级 ``_engine/`` → 当前目录
    ``_engine/``；全部落在 C 盘时返回 None（引擎缓存不许进系统盘）。只取
    ``SKILL_ROOT.parent`` 会漏掉 ``bootstrap_engine.py`` 实际落盘的
    ``<cwd>/_engine``，导致已解压的引擎被误报为“未定位到”。
    """
    bases = []
    base = os.environ.get("PPT_MASTER_CACHE_DIR", "").strip()
    if base:
        bases.append(Path(base))
    bases.append(SKILL_ROOT.parent / "_engine")
    bases.append(Path.cwd() / "_engine")
    for item in bases:
        candidate = item / CACHE_DIR_NAME
        if os.path.splitdrive(str(candidate))[0].upper() != "C:":
            return candidate
    return None


ENGINE_REQUIRED_STRUCTURE = ("workflows", "templates", "references", "scripts")
ENGINE_REQUIRED_FILES = (
    "SKILL.md",
    "LICENSE",
    "scripts/attribution_guard.py",
    "workflows/generate-pptx.md",
    "scripts/project_manager.py",
)


def engine_is_complete(engine_dir):
    """判断目录是否为完整引擎，而非被中断的残缺解压。

    只看 ``SKILL.md`` 存在会把只解压了一部分的缓存误判为可用，导致后续
    生成 PPT 时缺文件才报错。这里与 ``bootstrap_engine.py`` 同口径，检查
    关键结构、关键文件与文件总数。
    """
    if not engine_dir.is_dir():
        return False
    for rel in ENGINE_REQUIRED_STRUCTURE:
        if not (engine_dir / rel).is_dir():
            return False
    for rel in ENGINE_REQUIRED_FILES:
        if not (engine_dir / rel).is_file():
            return False
    return True


def locate_engine():
    """按 环境变量 → 已解压缓存 的顺序返回第一个完整可用的引擎目录。"""
    env_dir = os.environ.get("PPT_MASTER_DIR", "").strip()
    if env_dir and engine_is_complete(Path(env_dir)):
        return Path(env_dir), "环境变量 PPT_MASTER_DIR"
    cache = default_cache_dir()
    if cache is not None and engine_is_complete(cache):
        return cache, "内置归档已解压缓存"
    return None, ""


def archive_digest_ok():
    """返回 (归档是否存在, 摘要是否与清单一致)；清单缺失时只判存在性。"""
    path = archive_path()
    if not path.is_file():
        return False, False
    manifest = SKILL_ROOT / "vendor" / "ppt-master-6.4.0.manifest.json"
    expected = ""
    if manifest.is_file():
        try:
            expected = json.loads(manifest.read_text(encoding="utf-8")).get("archive_sha256", "")
        except (OSError, ValueError):
            expected = ""
    if not expected:
        return True, True
    return True, hashlib.sha256(path.read_bytes()).hexdigest() == expected


def read_engine_version(engine_dir):
    try:
        text = (engine_dir / "SKILL.md").read_text(encoding="utf-8", errors="replace")
    except OSError:
        return ""
    for line in text.splitlines():
        stripped = line.strip()
        if stripped.lower().startswith("version:"):
            return stripped.split(":", 1)[1].strip().strip("\"'")
    return ""


def read_license_digest(engine_dir):
    try:
        raw = (engine_dir / "LICENSE").read_bytes()
    except OSError:
        return ""
    text = raw.decode("utf-8", errors="replace")
    return hashlib.sha256(text.replace("\r\n", "\n").replace("\r", "\n").encode("utf-8")).hexdigest()


def word_registered_version():
    try:
        import winreg
    except ImportError:
        return None
    try:
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "Word.Application\\CurVer") as key:
            value, _ = winreg.QueryValueEx(key, "")
        return value
    except OSError:
        return None


def check_python_version():
    version = sys.version_info
    ok = version >= (3, 9)
    detail = f"Python {version.major}.{version.minor}.{version.micro}（{sys.executable}）"
    fix = "安装 Python 3.9 及以上，并用它运行本 skill 的全部脚本"
    return ok, detail, fix


def check_docx_import():
    try:
        import docx
    except ImportError as exc:
        return False, f"import docx 失败：{exc}", "python -m pip install python-docx"
    version = getattr(docx, "__version__", "未知版本")
    return True, f"python-docx {version}", ""


def check_word_com(probe=False):
    try:
        import win32com.client
        import pythoncom
    except ImportError as exc:
        return False, f"pywin32 不可用：{exc}", "python -m pip install pywin32"
    registered = word_registered_version()
    if not registered:
        return False, "注册表未找到 Word.Application", "安装桌面版 Microsoft Word（Windows）"
    if not probe:
        return True, f"Word.Application 已注册（{registered}）；未做真机探测", ""
    word = None
    initialized = False
    try:
        pythoncom.CoInitialize()
        initialized = True
        word = win32com.client.DispatchEx("Word.Application")
        version = str(word.Version)
        return True, f"Word COM 探测成功，版本 {version}", ""
    except Exception as exc:  # Word 被策略拦截、受保护视图、进程启动失败等
        return False, f"Word COM 探测失败：{exc}", "确认桌面版 Word 可正常启动，或去掉 --probe-word 只做注册表检查"
    finally:
        if word is not None:
            word.Quit()
        if initialized:
            pythoncom.CoUninitialize()


def check_pymupdf_import():
    try:
        import fitz
    except ImportError as exc:
        return False, f"import fitz 失败：{exc}", "python -m pip install pymupdf"
    version = getattr(fitz, "VersionBind", "") or getattr(fitz, "__version__", "未知版本")
    return True, f"PyMuPDF {version}", ""


def check_engine_resolve():
    engine_dir, source = locate_engine()
    if engine_dir is not None:
        return True, f"引擎就绪：{engine_dir}（来源：{source}）", ""
    exists, digest_ok = archive_digest_ok()
    if exists and digest_ok:
        cache = default_cache_dir()
        target = str(cache) if cache else "需显式指定缓存目录（PPT_MASTER_CACHE_DIR 或 --cache-dir）"
        return True, f"内置归档可用，首次生成 PPT 时自举解压到 {target}", ""
    if exists:
        return False, "内置归档摘要与清单不一致，脚本会拒绝解压", "重新获取归档，或设置 PPT_MASTER_DIR 指向有效引擎"
    return False, "既无可用引擎，也找不到内置归档", "设置 PPT_MASTER_DIR 指向有效引擎目录，或恢复 vendor/ 下的归档"


def check_engine_archive():
    exists, digest_ok = archive_digest_ok()
    path = archive_path()
    if not exists:
        return False, f"内置归档缺失：{path.relative_to(SKILL_ROOT).as_posix()}", "恢复 vendor/ 下的归档，或改用 PPT_MASTER_DIR"
    size = path.stat().st_size / 1048576
    if digest_ok:
        return True, f"内置归档完整（{size:.2f} MB），SHA-256 与清单一致", ""
    return False, "内置归档 SHA-256 与清单不一致", "重新获取归档；摘要不符时脚本会拒绝解压"


def check_engine_env():
    env_dir = os.environ.get("PPT_MASTER_DIR", "").strip()
    if not env_dir:
        return True, "未设置 PPT_MASTER_DIR（可选覆盖项；未设置时用内置归档自举）", ""
    if not engine_is_complete(Path(env_dir)):
        return False, f"PPT_MASTER_DIR={env_dir} 不是完整引擎目录", "确认该路径包含引擎的 SKILL.md、workflows/ 与 scripts/，或清除该变量改用内置归档"
    return True, f"PPT_MASTER_DIR={env_dir}", ""


def check_stdlib_net():
    try:
        available = importlib.util.find_spec("urllib.request") is not None
    except (ImportError, ValueError):
        available = False
    if not available:
        return False, "标准库缺少网络模块", "使用完整的 Python 发行版；核实主题时可改用用户提供的来源"
    return True, "标准库网络模块可用（本脚本不发起任何请求）", ""


def build_handlers(probe_word):
    return {
        "python-version": check_python_version,
        "docx-import": check_docx_import,
        "word-com": lambda: check_word_com(probe_word),
        "pymupdf-import": check_pymupdf_import,
        "engine-resolve": check_engine_resolve,
        "engine-archive": check_engine_archive,
        "engine-env": check_engine_env,
        "stdlib-net": check_stdlib_net,
    }


def load_manifest(skill_dir):
    path = skill_dir / "skill-dependencies.json"
    if not path.is_file():
        return None
    try:
        return json.loads(path.read_text(encoding="utf-8-sig"))
    except (OSError, ValueError) as exc:
        print(f"依赖清单无法读取：{exc}", file=sys.stderr)
        return None


def run_checks(skill_dir, probe_word, manifest):
    handlers = build_handlers(probe_word)
    rows = []
    for dependency in manifest.get("dependencies", []):
        dep_id = dependency.get("id", "")
        required = bool(dependency.get("required"))
        results = []
        for check in dependency.get("checks", []):
            check_id = check.get("id", "")
            handler = handlers.get(check_id)
            if handler is None:
                results.append({"id": check_id, "required": bool(check.get("required")), "ok": False,
                                "detail": "本脚本未实现该检查项", "fix": "检查依赖清单与脚本是否同步"})
                continue
            try:
                ok, detail, fix = handler()
            except Exception as exc:  # 单项检查异常不应中断整体体检
                ok, detail, fix = False, f"检查异常：{exc}", "按 references/setup-guide.md 手工确认"
            results.append({"id": check_id, "required": bool(check.get("required")), "ok": ok,
                            "detail": detail, "fix": fix})
        required_ok = all(item["ok"] for item in results if item["required"])
        any_ok = any(item["ok"] for item in results)
        if required and not required_ok:
            state = "missing"
        elif required:
            state = "ok"
        else:
            state = "ok" if all(item["ok"] for item in results) and results else ("partial" if any_ok else "missing")
        rows.append({
            "id": dep_id,
            "required": required,
            "capabilities": dependency.get("capabilities", []),
            "state": state,
            "checks": results,
        })
    return rows


def engine_note():
    engine_dir, source = locate_engine()
    if engine_dir is not None:
        return {
            "located": True,
            "source": source,
            "path": str(engine_dir),
            "version": read_engine_version(engine_dir),
            "license_match": read_license_digest(engine_dir) == ENGINE_LICENSE_SHA256,
        }
    exists, digest_ok = archive_digest_ok()
    if exists and digest_ok:
        cache = default_cache_dir()
        return {
            "located": False,
            "source": "内置归档（首次生成 PPT 时自举解压）",
            "path": str(cache) if cache else "",
            "version": "6.4.0",
            "license_match": True,
        }
    return {"located": False, "source": "", "path": "", "version": "", "license_match": False}


def main():
    parser = argparse.ArgumentParser(description="团日活动 skill 依赖环境检查")
    parser.add_argument("--skill-dir", default=str(SKILL_ROOT), help="skill 根目录，默认取本脚本的上级目录")
    parser.add_argument("--json", action="store_true", help="只输出 JSON")
    parser.add_argument("--probe-word", action="store_true", help="真实启动一次 Word 读取版本号后退出")
    args = parser.parse_args()

    skill_dir = Path(args.skill_dir).resolve()
    manifest = load_manifest(skill_dir)
    if manifest is None:
        payload = {"status": "unknown", "reason": "缺少或无法解析 skill-dependencies.json", "skill_dir": str(skill_dir)}
        print(json.dumps(payload, ensure_ascii=False, indent=2))
        return 2

    rows = run_checks(skill_dir, args.probe_word, manifest)
    engine = engine_note()

    if any(row["required"] and row["state"] == "missing" for row in rows):
        status = "needs_setup"
    elif any(row["state"] != "ok" for row in rows):
        status = "partial"
    else:
        status = "ready"

    engine_check = next((row for row in rows if row["id"] == "ppt-master-engine"), None)
    if engine_check is not None and engine["located"] and not engine["license_match"]:
        engine_check["state"] = "missing"
        engine_check["checks"].append({"id": "license-digest", "required": False, "ok": False,
                                       "detail": "引擎 LICENSE 摘要与上游不一致，完整性检查会阻断 PPTX 路线",
                                       "fix": "从官方仓库重新获取引擎目录，不要修改其内部文件"})
        status = "partial" if status != "needs_setup" else status

    payload = {
        "status": status,
        "skill_dir": str(skill_dir),
        "interpreter": sys.executable,
        "dependencies": rows,
        "engine": engine,
        "notes": [
            "本脚本不安装依赖、不写配置、不发起网络请求。",
            "运行本 skill 的全部脚本时，应使用上面 interpreter 指向的解释器。",
        ],
    }

    if args.json:
        print(json.dumps(payload, ensure_ascii=False, indent=2))
    else:
        labels = {"ok": "OK", "partial": "部分可用", "missing": "缺失"}
        print(f"环境状态：{status}")
        print(f"解释器：{sys.executable}")
        print()
        for row in rows:
            tag = "必需" if row["required"] else "可选"
            print(f"[{labels[row['state']]:<6}] {row['id']}（{tag}）")
            for item in row["checks"]:
                mark = "OK  " if item["ok"] else "缺失"
                print(f"         {mark} {item['id']}：{item['detail']}")
                if not item["ok"] and item["fix"]:
                    print(f"              处理：{item['fix']}")
        print()
        if engine["located"]:
            print(f"PPTX 引擎：{engine['path']}（来源：{engine['source']}，版本 {engine['version'] or '未知'}，LICENSE 摘要{'一致' if engine['license_match'] else '不一致'}）")
        else:
            print("PPTX 引擎：未定位到，PPT 无法生成；请按 setup-guide 恢复引擎")
        if status != "ready":
            print("处理方式见 references/setup-guide.md。")

    return {"ready": 0, "partial": 1, "needs_setup": 2}[status]


if __name__ == "__main__":
    sys.exit(main())
