import sys
import os
from pathlib import Path
import logging

logger = logging.getLogger(__name__)

_app_path = None
_user_data_path = None


def get_app_path() -> Path:
    global _app_path

    if _app_path is not None:
        return _app_path

    if getattr(sys, 'frozen', False):
        _app_path = Path(os.path.dirname(sys.executable))
        logger.info(f"Running from exe, app path: {_app_path}")
    else:
        _app_path = Path(__file__).parent.parent
        logger.info(f"Running from script, app path: {_app_path}")

    return _app_path


def get_user_data_path() -> Path:
    global _user_data_path

    if _user_data_path is not None:
        return _user_data_path

    if getattr(sys, 'frozen', False):
        appdata = os.environ.get('APPDATA', os.path.expanduser('~'))
        _user_data_path = Path(appdata) / 'ExcelRealtimeController'
        logger.info(f"User data path (exe): {_user_data_path}")
    else:
        _user_data_path = Path(__file__).parent.parent
        logger.info(f"User data path (script): {_user_data_path}")

    _user_data_path.mkdir(parents=True, exist_ok=True)
    return _user_data_path


def get_config_path(relative_path: str) -> Path:
    user_data = get_user_data_path()
    config_path = user_data / relative_path

    config_path.parent.mkdir(parents=True, exist_ok=True)

    logger.debug(f"Config path for '{relative_path}': {config_path}")
    return config_path

