import pandas as pd


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


def percorre_macs():
    sheet_name = str(input("Digite o nome da aba que vai ser adicionado os valores: "))
    principal = pd.read_excel("C:/Users/LUCASARRUDADACRUZ/Downloads/Trabalho Cristiano/Inventário de Ativos.xlsx", sheet_name=sheet_name, dtype=str)
    informacoes = pd.read_excel("C:/Users/LUCASARRUDADACRUZ/Downloads/Trabalho Cristiano/STAR.xlsx", dtype=str)

    principal["SO do STAR"] = ""
    principal["STAR"] = ""
    principal["Patrimônio no STAR"] = ""

    principal["Levantamento de Ativos - MACs"] = principal["Levantamento de Ativos - MACs"].astype(str)
    informacoes["mac"] = informacoes["mac"].astype(str)

    for linha_principal, valor_principal in enumerate(principal["Levantamento de Ativos - MACs"]):
        for linha_info, valor_info in enumerate(informacoes["mac"]):
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
        planilha.to_excel("C:/Users/LUCASARRUDADACRUZ/Downloads/Trabalho Cristiano/Resultado_Rio.xlsx", index=False)
        print("Planilha salva com sucesso!")
    except Exception as e:
        print(e)


def main():
    planilha_final = percorre_macs()
    salva_planilha(planilha_final)


if __name__ == "__main__":
    main()