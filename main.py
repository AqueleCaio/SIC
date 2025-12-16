import os
from openpyxl import load_workbook
from colorama import Fore, Style, init

init(autoreset=True)

# ================= CONFIGURAÇÕES =================
PASTAS_PLANILHAS = [
    r"\\fileceduc\grupos\ceduc_secretaria\PATRIMÔNIO\CEDUC_LEVANTAMENTO PATRIMÔNIO_2025",
    r"\\fileceduc\grupos\ceduc_secretaria\PATRIMÔNIO\2025_PATRIMÔNIO_NEOA"
]

COLUNAS = {
    "1": ("tombamento", 2),
    "2": ("patrimonio", 3),
    "3": ("inventario", 4),
    "4": ("especificacao", 5)
}
# =================================================


def limpar_terminal():
    os.system("cls" if os.name == "nt" else "clear")


def linha():
    print(Fore.CYAN + "." * 70)


def cabecalho(titulo):
    linha()
    print(Fore.YELLOW + Style.BRIGHT + titulo.center(70))
    linha()


def texto_com_tecla(texto, tecla, cor_tecla=Fore.GREEN):
    return texto.replace(
        tecla,
        cor_tecla + Style.BRIGHT + tecla + Style.RESET_ALL + Fore.WHITE
    )

def procurar(coluna_idx, valor, criterio):
    valor_original = valor.strip()
    valor = valor_original.upper()
    criterio = criterio.lower()
    encontrou = False

    # Mensagem inicial personalizada
    if criterio == "especificacao":
        print(
            Fore.MAGENTA + Style.BRIGHT +
            f"\n🔎 Procurando por: {Fore.YELLOW}{valor_original}\n"
        )
    else:
        print(
            Fore.MAGENTA + Style.BRIGHT +
            f"\n🔎 Procurando pelo item com número de "
            f"{Fore.YELLOW}{criterio.upper()}: {valor_original}\n"
        )

    for pasta in PASTAS_PLANILHAS:
        print(Fore.YELLOW + Style.BRIGHT + f"📂 Vasculhando pasta: {pasta}")

        if not os.path.exists(pasta):
            print(Fore.RED + "  Pasta não encontrada.\n")
            continue

        for arquivo in os.listdir(pasta):
            if not arquivo.endswith(".xlsx"):
                continue

            caminho = os.path.join(pasta, arquivo)

            try:
                wb = load_workbook(caminho, data_only=True)
            except:
                continue

            for aba in wb.sheetnames:
                ws = wb[aba]

                for i, linha_dados in enumerate(
                    ws.iter_rows(min_row=2, values_only=True), start=2
                ):
                    if not linha_dados or len(linha_dados) < 8:
                        continue

                    conteudo = (
                        str(linha_dados[coluna_idx]).strip().upper()
                        if linha_dados[coluna_idx] else ""
                    )

                    if conteudo == valor:
                        encontrou = True

                        # Identificação da origem
                        if "NEOA" in pasta.upper():
                            origem = Fore.GREEN + Style.BRIGHT + "NEOA"
                        elif "CEDUC" in pasta.upper():
                            origem = Fore.BLUE + Style.BRIGHT + "CEDUC"
                        else:
                            origem = Fore.WHITE + "DESCONHECIDA"

                        print("\n")
                        print(Fore.WHITE + "Origem: " + origem)
                        print(Fore.WHITE + f"Sala (arquivo): {arquivo}")
                        print(Fore.WHITE + f"Aba: {aba}")
                        print(Fore.WHITE + f"Linha: {i}")

                        linha()
                        print(Fore.CYAN + f"Item: {linha_dados[1]}")
                        print(Fore.CYAN + f"Tombamento: {linha_dados[2]}")
                        print(Fore.CYAN + f"Patrimônio: {linha_dados[3]}")
                        print(Fore.CYAN + f"Inventário: {linha_dados[4]}")
                        print(Fore.CYAN + f"Especificação: {linha_dados[5]}")
                        print(Fore.CYAN + f"TR: {linha_dados[6]}")
                        print(Fore.CYAN + f"Situação: {linha_dados[7]}")
                        linha()

    print(Fore.MAGENTA + Style.BRIGHT + "\nVarredura finalizada.")

    if not encontrou:
        print(Fore.RED + "✖ Nenhum resultado encontrado.")



def menu():
    while True:
        limpar_terminal()
        cabecalho("CONSULTA DE PATRIMÔNIO - CEDUC")

        print(Fore.WHITE + "Escolha o critério de busca:\n")
        print(Fore.GREEN + "1 - Número de Tombamento")
        print(Fore.GREEN + "2 - Número de Patrimônio")
        print(Fore.GREEN + "3 - Número de Inventário")
        print(Fore.GREEN + "4 - Especificação")
        print(Fore.RED + "0 - Sair")

        linha()
        opcao = input(Fore.YELLOW + "Opção: ").strip()

        if opcao == "0":
            limpar_terminal()
            print(Fore.MAGENTA + Style.BRIGHT + "Programa encerrado. Até mais 👋")
            break

        if opcao not in COLUNAS:
            print(Fore.RED + "\nOpção inválida.")
            input(texto_com_tecla(
                "\nPressione ENTER para continuar...",
                "ENTER",
                Fore.YELLOW
            ))
            continue

        nome, coluna_idx = COLUNAS[opcao]
        valor = input(Fore.YELLOW + f"Digite o valor para {nome.upper()}: ")

        limpar_terminal()
        cabecalho(f"RESULTADO DA BUSCA - {nome.upper()}")
        procurar(coluna_idx, valor, nome)

        input(texto_com_tecla(
            "\nPressione ENTER para voltar ao menu...",
            "ENTER",
            Fore.GREEN
        ))


# ========= EXECUÇÃO =========
menu()
