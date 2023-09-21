from __future__ import annotations
import os
import shutil
from pathlib import Path
from typing import Any, Iterable, List
import os
from shutil import which
from contextlib import contextmanager
import uno


@contextmanager
def change_dir(directory):
    """
    A context manager that changes the current working directory to the specified directory
    temporarily and then changes it back when the context is exited.
    """
    current_dir = os.getcwd()
    os.chdir(directory)
    try:
        yield
    finally:
        os.chdir(current_dir)


def find_file_in_parent_dirs(filename: str, path: str = "") -> str:
    """
    Recursively searches parent directories for a file with the given filename.
    Returns the absolute path to the file if found, or None if not found.
    """
    if path is None:
        path = os.getcwd()

    file_path = os.path.join(path, filename)
    if os.path.exists(file_path):
        return os.path.abspath(file_path)

    parent_dir = os.path.abspath(os.path.join(path, os.pardir))
    if parent_dir == path:
        # We've reached the root directory and haven't found the file
        return ""

    return find_file_in_parent_dirs(filename, parent_dir)


def mkdirp(self, dest_dir):
    # Python ≥ 3.5
    if isinstance(dest_dir, Path):
        dest_dir.mkdir(parents=True, exist_ok=True)
    else:
        Path(dest_dir).mkdir(parents=True, exist_ok=True)


def find_files_matching_patterns(root_dir: str | Path, ext: Iterable[str]) -> List[str]:
    """
    Finds all files in the given directory and its subdirectories that match the patterns *.txt and *.xml.
    Returns a list of absolute file paths.

    Args:
        root_dir (str | Path): The root directory to search.
        ext (List[str]): The file extensions to search for.
    Returns:
        List[str]: A list of absolute file paths.
    """
    if isinstance(root_dir, str):
        root_path = Path(root_dir)
    else:
        root_path = root_dir
    if not root_path.exists():
        raise FileNotFoundError(f"Directory '{root_dir}' not found")
    if not ext:
        return []

    file_paths = []
    extensions = set([f".{e.lower()}" for e in ext])
    for file_path in root_path.glob("**/*"):
        if file_path.is_file() and (file_path.suffix in extensions):
            file_paths.append(str(file_path.absolute()))

    return file_paths


def read_file(file_path: str, encoding="UTF-8") -> str:
    """Read the contents of the given file and return it as a string."""
    with open(file_path, "r", encoding=encoding) as f:
        return f.read()


def write_string_to_file(file_path: str, content: str, encoding="UTF-8") -> None:
    """Write the given string to the specified file."""
    with open(file_path, "w", encoding=encoding) as f:
        f.write(content)


def zip_folder(folder: str | Path, base_name: str = "", dest_dir: str | Path = "") -> None:
    """
    Zips all files in the given folder to the specified zip file.

    Args:
        folder (str | Path): is a directory that will be the root directory of the archive;
        base_name (str): is the name of the file to create, minus any format-specific
            extension; 'format' is the archive format: one of "zip", "tar", "gztar",
            "bztar", or "xztar".  Or any other registered format.

    Returns:
        None
    """
    if isinstance(folder, str):
        folder_path = Path(folder)
    else:
        folder_path = folder
    if not folder_path.is_dir():
        raise ValueError(f"Expected folder, got '{folder_path}'")
    if not folder_path.is_absolute():
        folder_path = folder_path.absolute()
    if not folder_path.exists():
        raise FileNotFoundError(f"Folder '{folder_path}' not found")
    if not base_name:
        base_name = folder_path.name

    if isinstance(dest_dir, str):
        if dest_dir == "":
            dest_dir = folder_path.parent
        else:
            dest_dir = Path(dest_dir)
    else:
        dest_dir = dest_dir.absolute()

    if not dest_dir.is_dir():
        raise ValueError(f"Expected folder, got '{dest_dir}'")
    if not dest_dir.exists():
        raise FileNotFoundError(f"Folder '{dest_dir}' not found")

    with change_dir(dest_dir):
        shutil.make_archive(base_name, "zip", folder_path)


def get_which(name: str | Path) -> str:
    """
    Returns the path to the executable which would be executed in the current
    environment, or empty string if no executable is found.

    Such as 'python', 'pip', 'git', etc.
    """
    result = which(name)
    if result is None:
        return ""
    return str(result)


def is_on_path(name: str | Path) -> bool:
    """
    Returns True if the given name is on the PATH, False otherwise.

    Such as 'python', 'pip', 'git', etc.
    """
    return which(name) is not None


def get_user_profile_path(as_sys_path: bool = True, ctx: Any = None) -> str:
    """
    Returns the path to the user profile directory.

    Args:
        as_sys_path (bool, optional): If True, returns the path as a system path entry otherwise ``file:///`` format.
            Defaults to True.
        ctx (Any, optional): The context to use. Defaults to None.
    """
    if ctx is None:
        ctx = uno.getComponentContext()
    result = ctx.ServiceManager.createInstance(
        "com.sun.star.util.PathSubstitution"
    ).substituteVariables(  # type: ignore
        "$(user)", True
    )
    return uno.fileUrlToSystemPath(result) if as_sys_path else result


def get_package_location(pkg_id: str, as_sys_path: bool = True, ctx: Any = None) -> str:
    """
    Gets the package location.

    Something link this: 'file:///home/user/.config/libreoffice/4/user/uno_packages/cache/uno_packages/lu323960zf5pw.tmp_/OooPip.oxt'

    Args:
        pkg_id (str): _description_
        as_sys_path (bool, optional): If True, returns the path as a system path entry otherwise ``file:///`` format.
            Defaults to True.
        ctx (Any, optional): The context to use. Defaults to None.

    Returns:
        str: _description_
    """
    # sourcery skip: reintroduce-else, swap-if-else-branches, use-named-expression
    if ctx is None:
        ctx = uno.getComponentContext()
    pip = ctx.getValueByName("/singletons/com.sun.star.deployment.PackageInformationProvider")
    # pip.getPackageLocation("org.openoffice.extensions.ooopip")
    result = pip.getPackageLocation(pkg_id)
    if not result:
        return ""
    return uno.fileUrlToSystemPath(result) if as_sys_path else result
