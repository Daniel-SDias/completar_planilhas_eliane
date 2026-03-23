from pathlib import Path
from .utils import *


def main():

    caminho_arquivo = Path(
        r"C:\Users\danie\OneDrive\Desktop\Velho burguer1- Planilha para importar-2026.xlsm")

    if identificar_extensao_arquivo(caminho_arquivo) == ".xlsm":
        wb = carregar_planilha(caminho_arquivo)
    elif identificar_extensao_arquivo(caminho_arquivo) == ".xlsx":
        wb = carregar_planilha(caminho_arquivo)
    else:
        raise Exception

    abas_planilhas = wb.sheetnames

    aba_exemplo = obter_aba_exemplo(abas_planilhas)
    if not aba_exemplo:
        raise Exception

    aba_referencia = wb[aba_exemplo]
    dados_aba_referencia = obter_dados_aba(aba_referencia)

    dict_referencia = criar_dict_referencia(dados_aba_referencia, aba_referencia)

    aba_atual = obter_aba_atual(abas_planilhas)
    if not aba_atual:
        raise Exception

    aba_atual = wb[aba_atual]
    dados_aba_atual = obter_dados_aba(aba_atual)

    preencher_dados(dados_aba_atual, dict_referencia, aba_atual)
                
    wb.save(caminho_arquivo)
