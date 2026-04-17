import logging


def setup_logging(level: str = "INFO") -> None:
    """Configure root logger with a clean, single-line format."""
    logging.basicConfig(
        level=level,
        format="%(asctime)s  %(levelname)-8s  %(name)s  %(message)s",
        datefmt="%H:%M:%S",
    )
