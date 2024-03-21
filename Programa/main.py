import pandas as pd
import sys
import re
import keyboard
import time


def separa_valores(linha):
    lista_macs = linha.split(", ")
    return lista_macs


def compara_macs(mac1, mac2):
    if mac1.upper() == mac2.upper():
        return True
    return False


def copia_dados(planilha1, planilha2, linha1, linha2, valores1, valores2):
    planilha1.at[linha1, valores1[0]] = "Sim"
    valores1 = valores1[1:]

    for i in range(len(valores1)):
        planilha1.at[linha1, valores1[i]] = planilha2.at[linha2, valores2[i]]

    return planilha1


def mac_eh_valido(mac):
    return bool(re.match(r"^([0-9A-Fa-f]{2}[ :\-]?){5}([0-9A-Fa-f]{2})$", mac))


def formata_mac(mac):
    mac_hex = re.sub(r"[^0-9A-Fa-f]", "", mac)
    mac_parts = [mac_hex[i:i+2] for i in range(0, len(mac_hex), 2)]
    formatted_mac = ":".join(mac_parts)
    return formatted_mac


def verifica_mac(planilha):
    for coluna in planilha.columns:
        contem_macs = planilha[coluna].apply(lambda x: isinstance(x, str) and mac_eh_valido(x)).any()
        if contem_macs:
            return coluna
    return "Nenhuma coluna encontrada com valores de MAC"
    

def armazena_planilha(pergunta="Arraste aqui o arquivo desejado: "):
    caminho_principal = sys.argv[1] if len(sys.argv) > 1 else input(pergunta)
    caminho_principal = caminho_principal.strip('"')
    xls = pd.ExcelFile(caminho_principal)
    nome_sheets = xls.sheet_names
    sheet_name = mostra_abas(nome_sheets)
    return pd.read_excel(caminho_principal, sheet_name=sheet_name, dtype=str)


def separa_colunas(planilha, mensagem="Digite o numero das colunas que serão armazenadas os valores: "):
    colunas = planilha.columns
    return mostra_abas(colunas, mensagem)


def mostra_abas(planilha, mensagem="Digite o número da aba que vai ser adicionado os valores: "):
    abas = [aba for aba in planilha]
    if len(abas) == 1:
        print("Planilha adicionada com sucesso!")
        time.sleep(3)
        return abas[0]
    for index, value in enumerate(abas, start=1): 
        print(f"{index} - {value}")
    index_value = str(input(mensagem))
    if index_value.isdigit():
        print("Planilha adicionada com sucesso!")
        time.sleep(3)
        return str(abas[int(index_value) - 1])
    else:
        numeros = re.findall(r'\d+', index_value)
        print("Colunas adicionadas com sucesso!")
        time.sleep(3)
        return [abas[int(indice) - 1] for indice in numeros]


def esvazia_coluna(planilha, colunas):
    for valor in colunas:
        planilha[valor] = ""
    return planilha


def percorre_valores(planilha_resultado, planilha_comparacao,
                    coluna_mac_principal, coluna_mac_info,
                    colunas_principal, colunas_info):
    for linha_principal, valor_principal in enumerate(planilha_resultado[coluna_mac_principal]):
        for linha_info, valor_info in enumerate(planilha_comparacao[coluna_mac_info]):
            if valor_principal == "nan":
                continue
            elif (",") in valor_info:
                lista_macs = separa_valores(str(valor_info))
                for  valor_lista in lista_macs:
                    mac_principal_formatado = formata_mac(valor_principal)
                    mac_lista_formatado = formata_mac(valor_lista)
                    if mac_principal_formatado == mac_lista_formatado:
                        planilha_resultado = copia_dados(planilha_resultado, planilha_comparacao,
                                                        linha_principal, linha_info,
                                                        colunas_principal, colunas_info)
            else:
                mac_principal_formatado = formata_mac(valor_principal)
                mac_info_formatado = formata_mac(valor_info)
                if mac_principal_formatado == mac_info_formatado:
                    planilha_resultado = copia_dados(planilha_resultado, planilha_comparacao,
                                                    linha_principal, linha_info,
                                                    colunas_principal, colunas_info)
    return planilha_resultado


def percorre_macs():
    principal = armazena_planilha()
    informacoes = armazena_planilha()

    coluna_mac_principal = verifica_mac(principal)
    coluna_mac_info = verifica_mac(informacoes)

    colunas_principal = separa_colunas(planilha=principal)
    colunas_info = separa_colunas(planilha=informacoes)

    principal = esvazia_coluna(principal, colunas_principal)

    principal[coluna_mac_principal] = principal[coluna_mac_principal].astype(str)
    informacoes[coluna_mac_info] = informacoes[coluna_mac_info].astype(str)

    return percorre_valores(principal, informacoes,
                            coluna_mac_principal, coluna_mac_info,
                            colunas_principal, colunas_info)


def salva_planilha(planilha):
    try:
        planilha.to_excel("Resultado.xlsx", index=False)
        print("Planilha salva com sucesso! Verifique o resultado na pasta do programa.")
        print("Pressione Enter para fechar o programa...")
        keyboard.wait('enter')
    except Exception as e:
        print(e)


def main():
    planilha_final = percorre_macs()
    salva_planilha(planilha_final)


if __name__ == "__main__":
    main()
