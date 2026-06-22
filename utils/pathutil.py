import configparser
import os
import shutil
from concurrent.futures import ProcessPoolExecutor, as_completed
from pathlib import Path
from typing import Optional

import psutil
import win32com.client

from log import log_error, log_info, log_warning


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


def paths_actual_equal(path1: str, path2: str) -> bool:
    """
    对比两个路径的最终物理路径是否相同，包括消解符号链接。

    Args:
        path1: 第一个路径字符串.
        path2: 第二个路径字符串.

    Returns:
        如果路径存在且实际物理路径相同，返回True.

    Raises:
        FileNotFoundError: 当路径不存在时抛出.
    """
    path1_resolved = Path(path1).resolve()
    path2_resolved = Path(path2).resolve()
    return path1_resolved == path2_resolved


def is_symlink_pointing_to(link_path: str, target_dir: str) -> bool:
    """
    判断符号链接是否指向指定的目录。

    Args:
        link_path (str): 要检查的符号链接路径。
        target_dir (str): 需要判断的目标目录的绝对路径。

    Returns:
        bool: 符号链接是否指向该目录。
    """
    link = Path(link_path)
    target = Path(target_dir)

    if not target.exists() or not link.exists():
        return False
    abs_symlink_target = None
    try:
        # 验证路径是符号链接
        if not link.is_symlink():
            return False

        # 获取符号链接的目标路径并转换为绝对路径
        symlink_target = link.readlink()

        abs_symlink_target = link.parent.joinpath(symlink_target).absolute()

        # 比较绝对路径
        return os.path.samefile(abs_symlink_target, target.absolute())
    except OSError as e:
        log_error(
            f"is_symlink_pointing_to link:{link} link_targe:{abs_symlink_target} target:{target} e:{e}"
        )
    return False


def copy_directory_contents(
    source_dir: str,
    target_dir: str,
    overwrite: bool = False,
    ignore_errors: bool = False,
):
    """
    将源目录内的所有内容（文件和子目录）复制到目标目录。

    Args:
        source_dir (str): 源目录路径（需包含内容）
        target_dir (str): 目标目录路径（直接接收内容，不新增子目录）
        overwrite (bool): 是否覆盖目标已存在的同名文件/目录 [default: False]
        ignore_errors (bool): 是否静默处理复制失败的条目 [default: False]
    """
    try:
        # 确保目标目录存在
        os.makedirs(target_dir, exist_ok=True)

        # 遍历源目录下的所有条目
        for entry in os.listdir(source_dir):
            source_path = os.path.join(source_dir, entry)
            target_path = os.path.join(target_dir, entry)

            try:
                # 处理子目录
                if os.path.isdir(source_path):
                    if os.path.exists(target_path) and overwrite:
                        log_info(f"删除已存在的子目录：{target_path}")
                        shutil.rmtree(target_path)
                    # 递归复制子目录（允许覆盖）
                    shutil.copytree(
                        source_path,
                        target_path,
                        dirs_exist_ok=overwrite,  # 根据是否覆盖设置参数
                    )

                # 处理文件
                else:
                    if os.path.exists(target_path):
                        if overwrite:
                            log_info(f"删除已存在的文件：{target_path}")
                            os.remove(target_path)
                        else:
                            continue  # 不覆盖则跳过

                    # 复制文件（保留元数据）
                    shutil.copy2(source_path, target_path)
                    log_info(f"成功复制文件：{source_path} → {target_path}")

            except OSError as e:
                if ignore_errors:
                    log_warning(f"跳过错误：{str(e)}")
                else:
                    raise  # 若不忽略，则抛出异常

        log_info("目录内容复制完成")

    except FileNotFoundError as e:
        log_error(f"源目录不存在: {source_dir}")
    except Exception as e:
        log_error(f"发生未知错误：{str(e)}")
        raise


class OpenFileMgr:
    open_files = set()

    @classmethod
    def refresh_open_files(cls):
        def check_process(proc):
            try:
                cls.open_files.update(
                    os.path.normcase(opened_file.path)
                    for opened_file in proc.open_files()
                )
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                pass

        processes = psutil.process_iter()
        with ProcessPoolExecutor(max_workers=os.cpu_count() * 2) as executor:
            futures = {executor.submit(check_process, p): p for p in processes}
            for _ in as_completed(futures):
                continue

    @classmethod
    def is_path_occupied_windows(cls, path):
        if not cls.open_files:
            cls.refresh_open_files()
        if path in cls.open_files:
            return True
        for opened_file in cls.open_files:
            if opened_file.startswith(path + os.sep):
                return True
        return False
