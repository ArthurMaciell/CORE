from datetime import datetime,timedelta
from app.utils.logging_config import logger



HORA_INICIO_UTEIS = 8
HORA_FIM_UTEIS = 18


def calcular_horas_uteis(data_solicitacao: str, hora_solicitacao: str, data_conclusao: str, hora_conclusao: str) -> float:
    """
    Calcula horas úteis entre data/hora de início e conclusão considerando dias úteis e expediente.
    """
    data_inicio = datetime.strptime(data_solicitacao, "%d/%m/%Y").date()
    data_fim = datetime.strptime(data_conclusao, "%d/%m/%Y").date()
    hora_inicio = datetime.strptime(hora_solicitacao, "%H:%M").time()
    hora_fim = datetime.strptime(hora_conclusao, "%H:%M").time()

    inicio = datetime.combine(data_inicio, hora_inicio)
    fim = datetime.combine(data_fim, hora_fim)

    if inicio >= fim:
        return 0

    total_horas = 0
    while inicio < fim:
        if inicio.weekday() < 5 and HORA_INICIO_UTEIS <= inicio.hour < HORA_FIM_UTEIS:
            proximo_ponto = min(fim, inicio.replace(hour=HORA_FIM_UTEIS, minute=0, second=0))
            total_horas += (proximo_ponto - inicio).total_seconds() / 3600
            inicio = proximo_ponto

        inicio += timedelta(days=1)
        inicio = inicio.replace(hour=HORA_INICIO_UTEIS, minute=0, second=0)
    
    logger.info(f"Calculando horas úteis entre {data_solicitacao} {hora_solicitacao} e {data_conclusao} {hora_conclusao}.")

    return round(total_horas, 2)

def calcular_horas_uteis_simples(inicio: datetime, fim: datetime) -> float:
    """
    Versão alternativa que recebe dois objetos datetime diretamente.
    """
    total_horas = 0
    atual = inicio

    while atual.date() <= fim.date():
        inicio_dia = max(atual.replace(hour=HORA_INICIO_UTEIS, minute=0), inicio)
        fim_dia = min(atual.replace(hour=HORA_FIM_UTEIS, minute=0), fim)

        if inicio_dia < fim_dia:
            total_horas += (fim_dia - inicio_dia).total_seconds() / 3600

        atual += timedelta(days=1)
        atual = atual.replace(hour=HORA_INICIO_UTEIS, minute=0)

    return round(total_horas, 2)