# ==============================================================================
# connectors/__init__.py
# Pacote de conectores do DataMiner Planeja+.
#
# Cada conector define os padrões de nomes de colunas esperados
# para um tipo específico de documento orçamentário municipal.
# A função detectar_colunas() de cada conector retorna um dicionário
# com o mapeamento das colunas encontradas no arquivo real.
# ==============================================================================

from connectors.siconfi import SICONFI
from connectors.loa import LOA
from connectors.ppa import PPA
from connectors.generico import Generico

# Registro central de conectores disponíveis
# Chave: nome exibido na interface
# Valor: classe do conector
CONECTORES = {
    "SICONFI": SICONFI,
    "LOA": LOA,
    "PPA": PPA,
    "LDO": Generico,
    "Balanço Orçamentário": Generico,
    "Outro": Generico
}


def obter_conector(tipo_documento: str):
    """
    Retorna a classe do conector correspondente ao tipo de documento.
    Usa o conector Genérico como fallback para tipos não mapeados.
    """
    return CONECTORES.get(tipo_documento, Generico)