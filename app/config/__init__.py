# VoterMail Engine — High-Volume Word & Excel DPT Mail-Merge Automation
# Copyright (c) 2026 Smurffahh (https://github.com/Smurffahh/votermail-engine)

__author__ = "Smurffahh"

from .settings import (
    BASE_DIR,
    DATA_DIR,
    TEMPLATES_DIR,
    OUTPUT_DIR,
    TEMP_DIR,
    LOGS_DIR,
    CONFIG_FILE_PATH,
    AppConfig,
    load_app_config,
    save_app_config,
    clean_temp_dir,
    get_default_word_path,
    get_default_excel_path,
)
from .theme import ThemeTokens, get_theme
from .version import (
    __app_name__,
    __author__,
    __repository__,
    __license__,
    __version__,
    FEATURE_REGISTRY,
    get_current_version,
)

__all__ = [
    "BASE_DIR",
    "DATA_DIR",
    "TEMPLATES_DIR",
    "OUTPUT_DIR",
    "TEMP_DIR",
    "LOGS_DIR",
    "CONFIG_FILE_PATH",
    "AppConfig",
    "load_app_config",
    "save_app_config",
    "clean_temp_dir",
    "get_default_word_path",
    "get_default_excel_path",
    "ThemeTokens",
    "get_theme",
    "__app_name__",
    "__author__",
    "__repository__",
    "__license__",
    "__version__",
    "FEATURE_REGISTRY",
    "get_current_version",
]
