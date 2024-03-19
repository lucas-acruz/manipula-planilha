import pandas as pd
import re


def separa_valores(linha):
    lista_macs = linha.split(", ")
    return lista_macs


def compara_macs(mac1, mac2):
    if mac1.upper() == mac2.upper():
        return True
    return False


def copia_dados(planilha1, planilha2, linha1, linha2):
    planilha1.at[linha1, "Patrimônio no STAR"] = planilha2.at[linha2, "endpoint_name"]
    planilha1.at[linha1, "SO do STAR"] = planilha2.at[linha2, "os_version"]
    planilha1.at[linha1, "STAR"] = "Sim"
    return planilha1


def mac_eh_valido(mac):
    return bool(re.match(r'^([0-9A-Fa-f]{2}[ :\-]?){5}([0-9A-Fa-f]{2})$', mac))


def verifica_mac(planilha):
    for coluna in planilha.columns:
        contem_macs = planilha[coluna].apply(lambda x: isinstance(x, str) and mac_eh_valido(x)).any()
        if contem_macs:
            print(coluna)
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
    mostra_abas(colunas)
    colunas = input(mensagem)


def mostra_abas(planilha):
    abas = []
    for contador, aba in enumerate(planilha, start=1):
        abas.append(aba)
        print(f"{contador} - {aba}")
    index_value = str(input("Digite o número da aba que vai ser adicionado os valores: "))
    if index_value.isdigit():
        return str(abas[int(index_value) - 1])
    else:
        numeros = re.findall(r'\d+', index_value)
        


def percorre_macs():
    principal = armazena_planilha()
    informacoes = armazena_planilha()

    coluna_mac_principal = verifica_mac(principal)
    coluna_mac_info = verifica_mac(informacoes)

    principal["SO do STAR"] = ""
    principal["STAR"] = ""
    principal["Patrimônio no STAR"] = ""

    principal[coluna_mac_principal] = principal[coluna_mac_principal].astype(str)
    informacoes[coluna_mac_info] = informacoes[coluna_mac_info].astype(str)

    for linha_principal, valor_principal in enumerate(principal[coluna_mac_principal]):
        for linha_info, valor_info in enumerate(informacoes[coluna_mac_info]):
            if valor_principal == "nan":
                continue
            elif (",") in valor_info:
                lista_macs = separa_valores(str(valor_info))
                for  valor_lista in lista_macs:
                    mac = compara_macs(valor_principal, valor_lista)
                    if mac:
                        principal = copia_dados(principal, informacoes, linha_principal, linha_info)
            else:
                mac = compara_macs(valor_principal, valor_info)
                if mac:
                    principal = copia_dados(principal, informacoes, linha_principal, linha_info)
    
    return principal


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
