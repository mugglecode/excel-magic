#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""Chromium download module."""

from io import BytesIO
import logging
import os
from pathlib import Path
import stat
import sys
from zipfile import ZipFile
import requests
import urllib3

logger = logging.getLogger(__name__)

DOWNLOADS_FOLDER = Path('') / 'local-chromium'
DEFAULT_DOWNLOAD_HOST = 'https://storage.googleapis.com'
DOWNLOAD_HOST = os.environ.get(
    'PYPPETEER_DOWNLOAD_HOST', DEFAULT_DOWNLOAD_HOST)
BASE_URL = f'{DOWNLOAD_HOST}/chromium-browser-snapshots'

REVISION = '737027'

NO_PROGRESS_BAR = os.environ.get('PYPPETEER_NO_PROGRESS_BAR', '')
if NO_PROGRESS_BAR.lower() in ('1', 'true'):
    NO_PROGRESS_BAR = True  # type: ignore

downloadURLs = {
    'linux': f'{BASE_URL}/Linux_x64/{REVISION}/chrome-linux.zip',
    'mac': f'{BASE_URL}/Mac/{REVISION}/chrome-mac.zip',
    'win32': f'{BASE_URL}/Win/{REVISION}/chrome-win32.zip',
    'win64': f'{BASE_URL}/Win_x64/{REVISION}/chrome-win32.zip',
}

chromiumExecutable = {
    'linux': 'https://npm.taobao.org/mirrors/chromium-browser-snapshots/Linux_x64/737027/chrome-linux.zip',
    'mac': 'https://npm.taobao.org/mirrors/chromium-browser-snapshots/Mac/737027/chrome-mac.zip',
    'win32': 'https://npm.taobao.org/mirrors/chromium-browser-snapshots/Win/737027/chrome-win.zip',
    'win64': 'https://npm.taobao.org/mirrors/chromium-browser-snapshots/Win_x64/737027/chrome-win.zip',
}


def current_platform() -> str:
    """Get current platform name by short string."""
    if sys.platform.startswith('linux'):
        return 'linux'
    elif sys.platform.startswith('darwin'):
        return 'mac'
    elif (sys.platform.startswith('win') or
          sys.platform.startswith('msys') or
          sys.platform.startswith('cyg')):
        if sys.maxsize > 2 ** 31 - 1:
            return 'win64'
        return 'win32'
    raise OSError('Unsupported platform: ' + sys.platform)


def get_url() -> str:
    """Get chromium download url."""
    return downloadURLs[current_platform()]


def extract_zip(data: BytesIO, path: Path) -> None:
    """Extract zipped data to path."""
    # On mac zipfile module cannot extract correctly, so use unzip instead.
    if current_platform() == 'mac':
        import subprocess
        import shutil
        zip_path = path / 'chrome.zip'
        if not path.exists():
            path.mkdir(parents=True)
        with zip_path.open('wb') as f:
            f.write(data.getvalue())
        if not shutil.which('unzip'):
            raise OSError('Failed to automatically extract chromium.'
                          f'Please unzip {zip_path} manually.')
        proc = subprocess.run(
            ['unzip', str(zip_path)],
            cwd=str(path),
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
        )
        if proc.returncode != 0:
            logger.error(proc.stdout.decode())
            raise OSError(f'Failed to unzip {zip_path}.')
    else:
        with ZipFile(data) as zf:
            zf.extractall(str(path))
    exec_path = chromium_executable()
    exec_path.chmod(exec_path.stat().st_mode | stat.S_IXOTH | stat.S_IXGRP |
                    stat.S_IXUSR)
    logger.warning(f'chromium extracted to: {path}')


def download_chromium() -> None:
    """Download and extract chromium."""
    # extract_zip(download_zip(get_url()), DOWNLOADS_FOLDER / REVISION)
    file = requests.get(chromiumExecutable['win32'])
    file = BytesIO(file.content)
    with open('chrome.zip', 'wb') as f:
        f.write(file.getvalue())
    extract_zip(file, Path(''))

def chromium_excutable() -> Path:
    """[Deprecated] miss-spelled function.

    Use `chromium_executable` instead.
    """
    logger.warning(
        '`chromium_excutable` function is deprecated. '
        'Use `chromium_executable instead.'
    )
    return chromium_executable()


def chromium_executable() -> Path:
    """Get path of the chromium executable."""
    return Path(chromiumExecutable[current_platform()])


def check_chromium() -> bool:
    """Check if chromium is placed at correct path."""
    return chromium_executable().exists()
