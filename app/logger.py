import logging

def setup_logger(name: str="logs", level=logging.DEBUG) -> logging.Logger:
    """Set up a logger with the specified name and log file.

    Args:
        name (str): The logger's title
        level (int): The logging level (default is logging.INFO).
    Returns:
        logging.Logger
    """
    # Create a logger
    log_file = name + ".log"
    logger = logging.getLogger(name)
    logger.setLevel(level)

    # Create a file handler
    file_handler = logging.FileHandler(log_file, mode="w+")
    file_handler.setLevel(level)

    # Create a console handler
    console_handler = logging.StreamHandler()
    console_handler.setLevel(level)

    # Create a formatter
    formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')

    # Add the formatter to the handlers
    file_handler.setFormatter(formatter)
    console_handler.setFormatter(formatter)

    # Add the handlers to the logger
    if not logger.handlers:
        logger.addHandler(file_handler)
        logger.addHandler(console_handler)

    return logger

# Example usage
if __name__ == "__main__":
    logger = setup_logger('my_logger')
    logger.info('This is an info message.')
    logger.error('This is an error message.')
