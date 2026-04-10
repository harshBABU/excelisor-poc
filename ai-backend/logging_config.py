import logging
import os

DEFAULT_LOG_FORMAT = "%(asctime)s | %(levelname)-8s | %(name)s | %(message)s"
DEFAULT_DATE_FMT = "%Y-%m-%d %H:%M:%S"


def configure_logging() -> None:
    root_logger = logging.getLogger()
    if not root_logger.handlers:
        level_name = os.getenv("AI_BACKEND_LOG_LEVEL", "DEBUG").upper()
        level = getattr(logging, level_name, logging.DEBUG)
        handler = logging.StreamHandler()
        formatter = logging.Formatter(fmt=DEFAULT_LOG_FORMAT, datefmt=DEFAULT_DATE_FMT)
        handler.setFormatter(formatter)
        root_logger.addHandler(handler)
    root_logger.setLevel(getattr(logging, os.getenv("AI_BACKEND_LOG_LEVEL", "DEBUG").upper(), logging.DEBUG))
