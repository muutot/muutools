import configparser
import getpass
import os
import shutil
from pathlib import Path, PurePath
from typing import Optional

import win32com.client

# 当前的用户
USER = getpass.getuser()

# 忽略输出文件, 文件名+后缀
IGNORE_OUTPUT = {"desktop.ini", "此电脑.lnk"}

# 从起始Folder开始的路径
# 如果为文件夹, 则文件夹下的文件全部忽略
PROTECTED_FILES = {"System"}

SRC_FOLDER = r"B:\Start Menu\Programs"

TARGET_FOLDER = (
    # rf"C:\Users\{user}\AppData\Roaming\Microsoft\Windows\Start Menu\Program\sync"
    # r"C:\ProgramData\Microsoft\Windows\Start Menu\Programs"
    r"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\sync"
)


def check_overwrite_ignore(dest_item: PurePath):
    return dest_item.name in IGNORE_OUTPUT


def check_is_protected(item: PurePath):
    relative_path = str(item)
    return any(relative_path.startswith(ignore) for ignore in PROTECTED_FILES)


def check_is_same(src_item, dest_item, relative_path):
    if not dest_item.exists():
        return False
    # 对于快捷方式文件，比较目标路径是否相同
    if (
        src_item.is_file()
        and dest_item.is_file()
        and src_item.suffix == dest_item.suffix
    ):
        return get_checker_by_suffix(src_item.suffix).check(
            src_item, dest_item, relative_path
        )
    elif src_item.is_dir() and dest_item.is_dir():
        # 如果都是目录，则认为相同
        show_output(f"[已存在目录]  {relative_path}")
        return True
    return False


def sync_folders(src_folder, dest_folder):
    """
    同步两个文件夹的内容

    Args:
        src_folder (str): 源文件夹路径
        dest_folder (str): 目标文件夹路径
    """
    src_path = Path(src_folder)
    dest_path = Path(dest_folder)

    # 确保源文件夹存在
    if not src_path.exists():
        raise FileNotFoundError(f"源文件夹不存在: {src_folder}")

    # 创建目标文件夹（如果不存在）
    dest_path.mkdir(parents=True, exist_ok=True)

    src_items = set()
    # 复制或更新源文件夹中的内容到目标文件夹
    for item in src_path.rglob("*"):
        relative_path = item.relative_to(src_path)
        if check_is_protected(relative_path):
            continue
        # 获取源文件夹中的所有文件和文件夹
        src_items.add(relative_path)

        dest_item = dest_path / relative_path
        # 判断文件是否相同
        if check_is_same(item, dest_item, relative_path):
            continue
        if item.is_dir():
            # 如果是目录，创建目录
            dest_item.mkdir(parents=True, exist_ok=True)
            print(f"创建目录: {relative_path}")
        elif item.is_file():
            # 如果是文件，复制文件（覆盖现有文件）
            dest_item.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(item, dest_item)
            if check_overwrite_ignore(dest_item):
                continue
            print(f"复制文件: {relative_path}")

    delete_redundant_files(dest_path, src_items)


def delete_redundant_files(dest_path, src_items):
    # 删除目标文件夹中源文件夹不存在的内容
    # 重新遍历并删除多余内容（从最深层开始，避免目录非空问题）
    items_to_delete: list[tuple[bool, Path]] = []
    for item in dest_path.rglob("*"):
        relative_path = item.relative_to(dest_path)
        if check_is_protected(relative_path):
            continue
        if relative_path not in src_items:
            items_to_delete.append((item.is_file(), item))

    # 按路径深度降序排列，确保文件在目录之前删除
    items_to_delete.sort(key=lambda x: len(str(x[1]).split(os.sep)), reverse=True)

    for is_file, item in items_to_delete:
        try:
            if is_file:
                item.unlink()
                print(f"删除文件: {item.relative_to(dest_path)}")
            else:
                item.rmdir()
                print(f"删除目录: {item.relative_to(dest_path)}")
        except OSError as e:
            print(f"删除失败 {item}: {e}")


def get_checker_by_suffix(suffix):
    if suffix == ".lnk":
        return LnkChecker()
    elif suffix == ".url":
        return URLChecker()
    return BaskChecker()


class BaskChecker:
    suffix = ""

    @classmethod
    def check(cls, src_item, dest_item, relative_path):
        return False


class LnkChecker(BaskChecker):
    def check(self, src_item, dest_item, relative_path):
        try:
            src_target = self.get_lnk_target(str(src_item))
            dest_target = self.get_lnk_target(str(dest_item))
            if not src_target or not dest_target:
                return False
            if src_target == dest_target:
                show_output(f"快捷方式已存在: {relative_path}")
                return True
        except (OSError, NotImplementedError) as e:
            print(e, src_item, dest_item)
            # 如果无法读取链接或不是链接文件，则按普通文件处理
            return False

    @staticmethod
    def get_lnk_target(lnk_path):
        """获取快捷方式的目标路径"""
        try:
            # 创建Windows脚本外壳对象
            shell = win32com.client.Dispatch("WScript.Shell")
            shortcut = shell.CreateShortCut(lnk_path)
            return shortcut.Targetpath
        except OSError as e:
            print(f"无法读取快捷方式 {lnk_path}: {e}")
            return None


class URLChecker(BaskChecker):
    def check(self, src_item, dest_item, relative_path):
        try:
            src_target = self.read_url_file(str(src_item))
            dest_target = self.read_url_file(str(dest_item))
            if not src_target or not dest_target:
                return False
            if src_target == dest_target:
                show_output(f"快捷方式已存在: {relative_path}")
                return True
        except (OSError, NotImplementedError) as e:
            # 如果无法读取链接或不是链接文件，则按普通文件处理
            return False

    @staticmethod
    def read_url_file(file_path: str) -> Optional[str]:
        """
        读取Windows .url文件并返回网址信息
        """
        # 检查文件是否存在
        if not os.path.exists(file_path):
            return None

        # 创建配置解析器
        config = configparser.ConfigParser()

        # 读取文件（指定编码为UTF-8）
        config.read(file_path, encoding="utf-8")

        # 检查是否存在[InternetShortcut]节
        if "InternetShortcut" not in config:
            return None

        # 提取网址信息
        shortcut_section = config["InternetShortcut"]

        # 提取常用字段
        if "URL" in shortcut_section:
            return shortcut_section["URL"].strip()
        return None


def main():
    try:
        sync_folders(SRC_FOLDER, TARGET_FOLDER)
        print("同步完成!")
    except Exception as e:
        print(f"同步失败: {e}")


exist_output = False


def show_output(info):
    if not exist_output:
        return
    print(info)


if __name__ == "__main__":
    main()
