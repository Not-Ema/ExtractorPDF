import logging
import json
import os

def setup_logger(config_path='config/config.json'):
    """Setup centralized logging based on config."""
    with open(config_path, 'r', encoding='utf-8') as f:
        config = json.load(f)

    settings = config.get('settings', {})
    log_level = getattr(logging, settings.get('log_level', 'INFO').upper(), logging.INFO)
    log_file = settings.get('log_file', 'extractor.log')

    logger = logging.getLogger('extractor')
    logger.setLevel(log_level)

    # Remove existing handlers
    for handler in logger.handlers[:]:
        logger.removeHandler(handler)

    # File handler
    file_handler = logging.FileHandler(log_file, encoding='utf-8')
    file_handler.setLevel(log_level)

    # Console handler
    console_handler = logging.StreamHandler()
    console_handler.setLevel(log_level)

    # Formatter
    formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
    file_handler.setFormatter(formatter)
    console_handler.setFormatter(formatter)

    logger.addHandler(file_handler)
    logger.addHandler(console_handler)

    return logger

# Global logger instance
logger = setup_logger()