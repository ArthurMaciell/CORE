import logging
from pathlib import Path

# Cria a pasta de logs se não existir
Path("logs").mkdir(exist_ok=True)

# Configuração do logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler("logs/core.log"),
        logging.StreamHandler()  # opcional: mostra no terminal também
    ]
)

logger = logging.getLogger("CORE")