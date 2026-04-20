import logging


def setup_logging(level: str = "INFO"):
    """
    Configura el nivel de logging de la librería.

    Args:
        level: nivel de logging. Opciones: "DEBUG", "INFO", "WARNING", "ERROR"

    Ejemplo:
        from msgraph_client.utils import setup_logging
        setup_logging("DEBUG")  # ver todo
        setup_logging("ERROR")  # ver solo errores
    """
    levels = {
        "DEBUG": logging.DEBUG,
        "INFO": logging.INFO,
        "WARNING": logging.WARNING,
        "ERROR": logging.ERROR,
    }

    log_level = levels.get(level.upper(), logging.INFO)

    logger = logging.getLogger("msgraph_client")
    logger.setLevel(log_level)

    # Si no tiene handlers, agregar uno a consola
    if not logger.handlers:
        handler = logging.StreamHandler()
        handler.setLevel(log_level)
        formatter = logging.Formatter(
            "[%(asctime)s] %(levelname)s - %(name)s - %(message)s",
            datefmt="%Y-%m-%d %H:%M:%S"
        )
        handler.setFormatter(formatter)
        logger.addHandler(handler)
    else:
        # Actualizar nivel en handlers existentes
        for handler in logger.handlers:
            handler.setLevel(log_level)