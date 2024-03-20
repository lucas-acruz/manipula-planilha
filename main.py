import pandas as pd
import re


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
    return bool(re.match(r'^([0-9A-Fa-f]{2}[ :\-]?){5}([0-9A-Fa-f]{2})$', mac))


def verifica_mac(planilha):
    for coluna in planilha.columns:
        contem_macs = planilha[coluna].apply(lambda x: isinstance(x, str) and mac_eh_valido(x)).any()
        if contem_macs:
            return coluna
    else:
        return "Nenhuma coluna encontrada com valores de MAC"
    

def armazena_planilha(pergunta="Arraste aqui o arquivo desejado: "):
    caminho_principal = str(input(pergunta))
    caminho_principal_formatado = r"{}".format(re.sub(r"^\s*&\s*'|'\s*$", "", caminho_principal))
    xls = pd.ExcelFile(caminho_principal_formatado)
    nome_sheets = xls.sheet_names
    sheet_name = mostra_abas(nome_sheets)
    return pd.read_excel(caminho_principal_formatado, sheet_name=sheet_name, dtype=str)


def separa_colunas(planilha, mensagem="Digite o numero das colunas que serão armazenadas os valores: "):
    colunas = planilha.columns
    return mostra_abas(colunas, mensagem)


def mostra_abas(planilha, mensagem="Digite o número da aba que vai ser adicionado os valores: "):
    abas = []
    colunas = []
    for contador, aba in enumerate(planilha, start=1):
        abas.append(aba)
        print(f"{contador} - {aba}")
    index_value = str(input(mensagem))
    if index_value.isdigit():
        return str(abas[int(index_value) - 1])
    else:
        numeros = re.findall(r'\d+', index_value)
        for indice in numeros:
            colunas.append(abas[int(indice) - 1])
        return colunas


def esvazia_coluna(planilha, colunas):
    for valor in colunas:
        planilha[valor] = ""
    return planilha


def percorre_valores(planilha_resultado, planilha_comparacao, coluna_mac_principal, coluna_mac_info, colunas_principal, colunas_info):
    for linha_principal, valor_principal in enumerate(planilha_resultado[coluna_mac_principal]):
        for linha_info, valor_info in enumerate(planilha_comparacao[coluna_mac_info]):
            if valor_principal == "nan":
                continue
            elif (",") in valor_info:
                lista_macs = separa_valores(str(valor_info))
                for  valor_lista in lista_macs:
                    mac = compara_macs(valor_principal, valor_lista)
                    if mac:
                        planilha_resultado = copia_dados(planilha_resultado, planilha_comparacao, linha_principal, linha_info, colunas_principal, colunas_info)
            else:
                if compara_macs(valor_principal, valor_info):
                    planilha_resultado = copia_dados(planilha_resultado, planilha_comparacao, linha_principal, linha_info, colunas_principal, colunas_info)
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

    return percorre_valores(principal, informacoes, coluna_mac_principal, coluna_mac_info, colunas_principal, colunas_info)


def salva_planilha(planilha):
    try:
        planilha.to_excel("Resultado.xlsx", index=False)
        print("Planilha salva com sucesso!")
    except Exception as e:
        print(e)


def main():
    planilha_final = percorre_macs()
    salva_planilha(planilha_final)


if __name__ == "__main__":
    main()
