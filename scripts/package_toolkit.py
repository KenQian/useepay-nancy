from datetime import datetime
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED


REPO_ROOT = Path(__file__).resolve().parents[1]
DIST_DIR = REPO_ROOT / "dist"
PACKAGE_NAME = "useepay_toolkit"

INCLUDE_PATHS = [
    REPO_ROOT / "src" / "各通道需换汇情况汇总.cmd",
    REPO_ROOT / "src" / "商户交易异常检测.cmd",
    REPO_ROOT / "src" / "fx_summary_workflow",
    REPO_ROOT / "src" / "merchant_analyzer",
    REPO_ROOT / "src" / "tools" / "compare_csv_files.py",
]

EXCLUDE_NAMES = {
    "__pycache__",
    ".DS_Store",
}


def iter_files(path: Path):
    if path.is_file():
        yield path
        return

    for child in path.rglob("*"):
        if child.is_dir():
            continue
        if any(part in EXCLUDE_NAMES for part in child.parts):
            continue
        yield child


def main():
    DIST_DIR.mkdir(exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d%H%M")
    zip_path = DIST_DIR / f"{PACKAGE_NAME}_{timestamp}.zip"
    if zip_path.exists():
        zip_path.unlink()

    with ZipFile(zip_path, "w", compression=ZIP_DEFLATED) as zip_file:
        for include_path in INCLUDE_PATHS:
            for file_path in iter_files(include_path):
                if file_path.is_relative_to(REPO_ROOT / "src"):
                    arcname = file_path.relative_to(REPO_ROOT / "src")
                else:
                    arcname = file_path.relative_to(REPO_ROOT)
                zip_file.write(file_path, arcname)

    print(f"Created {zip_path}")


if __name__ == "__main__":
    main()
