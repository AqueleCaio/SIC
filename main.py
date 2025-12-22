import os
from openpyxl import load_workbook
from colorama import Fore, Style, init
from itens import extract_items_from_pdf
import fitz  

init(autoreset=True)


REPORTS_FOLDER = os.path.join(os.getcwd(), "relatorios")


SPREADSHEET_FOLDERS = [
    r"\\fileceduc\grupos\ceduc_secretaria\PATRIMÔNIO\CEDUC_LEVANTAMENTO PATRIMÔNIO_2025",
    r"\\fileceduc\grupos\ceduc_secretaria\PATRIMÔNIO\2025_PATRIMÔNIO_NEOA"
]


SEARCH_COLUMNS = {
    "1": ("tombamento", 2),
    "2": ("patrimonio", 3),
    "3": ("inventario", 4),
    "4": ("especificacao", 5)
}


def clear_terminal():
    os.system("cls" if os.name == "nt" else "clear")


def print_line():
    print(Fore.CYAN + "." * 70)


def print_header(title):
    print_line()
    print(Fore.YELLOW + Style.BRIGHT + title.center(70))
    print_line()


def highlight_key(text, key, key_color=Fore.GREEN):
    return text.replace(
        key,
        key_color + Style.BRIGHT + key + Style.RESET_ALL + Fore.WHITE
    )


def extract_room_from_filename(filename):
    """
    Extrai somente código + nome da sala a partir do nome do arquivo.
    """

    # Remove extensão
    name = os.path.splitext(filename)[0]

    # Remove tudo entre parênteses
    name = re.sub(r"\(.*?\)", "", name)

    # Normaliza espaços
    name = name.strip()

    parts = [p.strip() for p in name.split("_") if p.strip()]

    if not parts:
        return name

    room_parts = []

    for part in parts:
        upper = part.upper()

        # Para quando começar dados administrativos ou nomes
        if any(keyword in upper for keyword in [
            "CEDUC",
            "NEOA",
            "SUPORTE",
            "GABINETE",
            "PROF",
            "DOCENTE",
            "AL",
            "ANA",
            "CAROLINA",
            "RODRIGUES",
            "OLIVEIRA"
        ]):
            break

        room_parts.append(part)

    return "_".join(room_parts)



def draw_check(page, x, y, size, color, width=1.5):
    """
    Desenha um ✓ vetorial usando 2 linhas
    """
    page.draw_line(
        fitz.Point(x - size * 0.4, y),
        fitz.Point(x - size * 0.1, y + size * 0.4),
        color=color,
        width=width
    )

    page.draw_line(
        fitz.Point(x - size * 0.1, y + size * 0.4),
        fitz.Point(x + size * 0.5, y - size * 0.5),
        color=color,
        width=width
    )


def draw_x(page, x, y, size, color, width=1.5):
    """
    Desenha um ✗ vetorial usando 2 linhas
    """
    page.draw_line(
        fitz.Point(x - size * 0.5, y - size * 0.5),
        fitz.Point(x + size * 0.5, y + size * 0.5),
        color=color,
        width=width
    )

    page.draw_line(
        fitz.Point(x - size * 0.5, y + size * 0.5),
        fitz.Point(x + size * 0.5, y - size * 0.5),
        color=color,
        width=width
    )


def generate_checked_pdf(original_pdf, output_pdf, tombamento_results):
    """
    Gera uma cópia do PDF com ícones vetoriais:

    ✓ → encontrado (verde)
    ✗ → não encontrado (vermelho)
    """

    doc = fitz.open(original_pdf)

    GREEN = (0.0, 0.6, 0.25)
    RED = (0.75, 0.15, 0.15)

    ICON_X = 16      # distância da borda esquerda
    ICON_SIZE = 8    # tamanho do ícone
    STROKE = 1.5     # espessura do traço

    for page in doc:
        words = page.get_text("words")

        for tombamento, found in tombamento_results.items():
            line_words = [w for w in words if tombamento in w[4]]
            if not line_words:
                continue

            for w in line_words:
                _, y0, _, y1, _, block, line_no, _ = w

                same_line = [
                    word for word in words
                    if word[5] == block and word[6] == line_no
                ]
                if not same_line:
                    continue

                center_y = (y0 + y1) / 2
                color = GREEN if found else RED

                if found:
                    draw_check(page, ICON_X, center_y, ICON_SIZE, color, STROKE)
                else:
                    draw_x(page, ICON_X, center_y, ICON_SIZE, color, STROKE)

    doc.save(output_pdf)
    doc.close()


def search_items_from_pdf(pdf_path):
    report_items = extract_items_from_pdf(pdf_path)

    if not report_items:
        print(Fore.RED + "Nenhum item encontrado no relatório.")
        return

    found_tombamentos = {}

    # Pré-varredura: coleta todos os tombamentos existentes nas planilhas
    for folder in SPREADSHEET_FOLDERS:
        if not os.path.exists(folder):
            continue

        for file in os.listdir(folder):
            if not file.endswith(".xlsx"):
                continue

            file_path = os.path.join(folder, file)

            try:
                workbook = load_workbook(file_path, data_only=True)
            except:
                continue

            for sheet_name in workbook.sheetnames:
                sheet = workbook[sheet_name]

                for row in sheet.iter_rows(min_row=2, values_only=True):
                    if not row or len(row) < 3:
                        continue

                    tombamento_cell = str(row[2]).strip() if row[2] else ""
                    if tombamento_cell and tombamento_cell not in found_tombamentos:
                        # nome do arquivo sem extensão representa a sala
                        sala = extract_room_from_filename(file)
                        found_tombamentos[tombamento_cell] = sala




    print(Fore.MAGENTA + Style.BRIGHT + "\nVerificando itens do relatório:\n")

    tombamento_results = {}

    for item in report_items:
        tombamento = item["tombamento"]
        denominacao = item["denominacao"]

        if tombamento in found_tombamentos:
            sala = found_tombamentos[tombamento]
            tombamento_results[tombamento] = True

            print(
                Fore.GREEN + Style.BRIGHT +
                f"✔ Tombamento: {tombamento} | Item: {denominacao} "
                f"- encontrado na sala ({sala})"
            )
        else:
            tombamento_results[tombamento] = False

            print(
                Fore.RED + Style.BRIGHT +
                f"✖ Tombamento: {tombamento} | Item: {denominacao}"
            )

    print(Fore.MAGENTA + Style.BRIGHT + "\nVerificação finalizada.")

    # Gera PDF marcado
    output_pdf = os.path.join(
        REPORTS_FOLDER,
        "RELATORIO_VERIFICADO.pdf"
    )

    generate_checked_pdf(
        original_pdf=pdf_path,
        output_pdf=output_pdf,
        tombamento_results=tombamento_results
    )

    print(
        Fore.CYAN + Style.BRIGHT +
        f"\n📄 PDF gerado com marcações: {output_pdf}"
    )



def search_items(column_index, value, criterion):
    original_value = value.strip()
    search_value = original_value.upper()
    criterion = criterion.lower()
    found = False

    # Mensagem inicial personalizada
    if criterion == "especificacao":
        print(
            Fore.MAGENTA + Style.BRIGHT +
            f"\n🔎 Procurando por: {Fore.YELLOW}{original_value}\n"
        )
    else:
        print(
            Fore.MAGENTA + Style.BRIGHT +
            f"\n🔎 Procurando pelo item com número de "
            f"{Fore.YELLOW}{criterion.upper()}: {original_value}\n"
        )

    for folder in SPREADSHEET_FOLDERS:
        print(Fore.YELLOW + Style.BRIGHT + f"📂 Vasculhando pasta: {folder}")

        if not os.path.exists(folder):
            print(Fore.RED + "  Pasta não encontrada.\n")
            continue

        for file in os.listdir(folder):
            if not file.endswith(".xlsx"):
                continue

            file_path = os.path.join(folder, file)

            try:
                workbook = load_workbook(file_path, data_only=True)
            except:
                continue

            for sheet_name in workbook.sheetnames:
                sheet = workbook[sheet_name]

                for row_index, row in enumerate(
                    sheet.iter_rows(min_row=2, values_only=True), start=2
                ):
                    if not row or len(row) < 8:
                        continue

                    cell_value = (
                        str(row[column_index]).strip().upper()
                        if row[column_index] else ""
                    )

                    if cell_value == search_value:
                        found = True

                        # Identificação da origem
                        if "NEOA" in folder.upper():
                            origin = Fore.GREEN + Style.BRIGHT + "NEOA"
                        elif "CEDUC" in folder.upper():
                            origin = Fore.BLUE + Style.BRIGHT + "CEDUC"
                        else:
                            origin = Fore.WHITE + "DESCONHECIDA"

                        print("\n")
                        print(Fore.WHITE + "Origem: " + origin)
                        print(Fore.WHITE + f"Sala (arquivo): {file}")
                        print(Fore.WHITE + f"Aba: {sheet_name}")
                        print(Fore.WHITE + f"Linha: {row_index}")

                        print_line()
                        print(Fore.CYAN + f"Item: {row[1]}")
                        print(Fore.CYAN + f"Tombamento: {row[2]}")
                        print(Fore.CYAN + f"Patrimônio: {row[3]}")
                        print(Fore.CYAN + f"Inventário: {row[4]}")
                        print(Fore.CYAN + f"Especificação: {row[5]}")
                        print(Fore.CYAN + f"TR: {row[6]}")
                        print(Fore.CYAN + f"Situação: {row[7]}")
                        print_line()

    print(Fore.MAGENTA + Style.BRIGHT + "\nVarredura finalizada.")

    if not found:
        print(Fore.RED + "✖ Nenhum resultado encontrado.")


def list_pdf_reports():
    if not os.path.exists(REPORTS_FOLDER):
        print(Fore.RED + "Pasta 'relatorios' não encontrada.")
        return []

    pdf_files = [
        f for f in os.listdir(REPORTS_FOLDER)
        if f.lower().endswith(".pdf")
    ]

    if not pdf_files:
        print(Fore.RED + "Nenhum relatório PDF encontrado na pasta 'relatorios'.")
        return []

    print(Fore.WHITE + "Relatórios disponíveis:\n")

    for index, pdf in enumerate(pdf_files, start=1):
        print(Fore.GREEN + f"{index} - {pdf}")

    return pdf_files


def run_menu():
    while True:
        clear_terminal()
        print_header("CONSULTA DE PATRIMÔNIO - CEDUC")

        print(Fore.WHITE + "Escolha o critério de busca:\n")
        print(Fore.GREEN + "1 - Número de Tombamento")
        print(Fore.GREEN + "2 - Número de Patrimônio")
        print(Fore.GREEN + "3 - Número de Inventário")
        print(Fore.GREEN + "4 - Especificação")
        print(Fore.GREEN + "5 - Procurar itens do relatório PDF nas planilhas")
        print(Fore.RED + "0 - Sair")

        print_line()
        option = input(Fore.YELLOW + "Opção: ").strip()

        if option == "0":
            clear_terminal()
            print(Fore.MAGENTA + Style.BRIGHT + "Programa encerrado. Até mais 👋")
            break


        if option == "5":
            clear_terminal()
            print_header("VERIFICAÇÃO DE ITENS DO RELATÓRIO PDF")

            pdf_files = list_pdf_reports()

            if not pdf_files:
                input(highlight_key(
                    "\nPressione ENTER para voltar ao menu...",
                    "ENTER",
                    Fore.GREEN
                ))
                continue

            print_line()
            choice = input(Fore.YELLOW + "Escolha o relatório pelo número: ").strip()

            if not choice.isdigit() or not (1 <= int(choice) <= len(pdf_files)):
                print(Fore.RED + "\nOpção inválida.")
                input(highlight_key(
                    "\nPressione ENTER para voltar ao menu...",
                    "ENTER",
                    Fore.GREEN
                ))
                continue

            selected_pdf = pdf_files[int(choice) - 1]
            pdf_path = os.path.join(REPORTS_FOLDER, selected_pdf)


            clear_terminal()
            print_header("RESULTADO DA VERIFICAÇÃO DO RELATÓRIO")
            search_items_from_pdf(pdf_path)

            input(highlight_key(
                "\nPressione ENTER para voltar ao menu...",
                "ENTER",
                Fore.GREEN
            ))
            continue


        if option not in SEARCH_COLUMNS:
            print(Fore.RED + "\nOpção inválida.")
            input(highlight_key(
                "\nPressione ENTER para continuar...",
                "ENTER",
                Fore.YELLOW
            ))
            continue

        criterion_name, column_index = SEARCH_COLUMNS[option]
        value = input(Fore.YELLOW + f"Digite o valor para {criterion_name.upper()}: ")

        clear_terminal()
        print_header(f"RESULTADO DA BUSCA - {criterion_name.upper()}")
        search_items(column_index, value, criterion_name)

        input(highlight_key(
            "\nPressione ENTER para voltar ao menu...",
            "ENTER",
            Fore.GREEN
        ))


# ==================================================
# APPLICATION ENTRY POINT
# ==================================================

if __name__ == "__main__":
    run_menu()
