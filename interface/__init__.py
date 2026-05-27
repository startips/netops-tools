from .connection import *
from .public_env import *
from .bitFunctions import *

__all__ = ['logg', 'autoThreadingPool',
           'deviceControl_auto', 'excel', 'init', 'set_value', 'get_value', 'readTxt', 'ping_check',
           'CHECK_ITEM_NAMES', 'DEVICE_TYPE_CONFIGS', 'DEVICE_TYPE_PATTERNS', 'CABLE_CHECK_CONFIG']


def __getattr__(name):
    # 懒加载：check_items 的导出名在首次访问时才读 YAML
    if name in ('CHECK_ITEM_NAMES', 'DEVICE_TYPE_CONFIGS', 'DEVICE_TYPE_PATTERNS', 'CABLE_CHECK_CONFIG'):
        from . import check_items
        return getattr(check_items, name)
    raise AttributeError(f"module {__name__!r} has no attribute {name!r}")
