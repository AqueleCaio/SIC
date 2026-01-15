import os
import re
import fitz  
import time
import threading
from InquirerPy import prompt
from openpyxl import load_workbook
from collections import defaultdict
from colorama import Fore, Style, init
from openpyxl.styles import PatternFill
from itens import extract_items_from_pdf

init(autoreset=True)


# ==================================================
# CONFIGURAÇÕES GLOBAIS
# ==================================================

REPORTS_FOLDER = os.path.join(os.getcwd(), "relatorios")
VERIFIED_REPORTS_FOLDER = os.path.join(os.getcwd(), "relatorios_verificados")

# pastas reais do CEDUC E DO NEOA
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

# ==================================================
# CLASSE: VIEW (INTERFACE DO USUÁRIO)
# ==================================================

class View:
    """Responsável por toda a interface com o usuário"""
    
    @staticmethod
    def clear_terminal():
        """Limpa o terminal"""
        os.system("cls" if os.name == "nt" else "clear")
    
    @staticmethod
    def print_line():
        """Imprime linha decorativa"""
        print(Fore.CYAN + "." * 70)
    
    @staticmethod
    def print_header(title):
        """Imprime cabeçalho formatado"""
        View.print_line()
        print(Fore.YELLOW + Style.BRIGHT + title.center(70))
        View.print_line()
    
    @staticmethod
    def highlight_key(text, key, key_color=Fore.GREEN):
        """Destaca uma palavra-chave no texto"""
        return text.replace(
            key,
            key_color + Style.BRIGHT + key + Style.RESET_ALL + Fore.WHITE
        )
    
    @staticmethod
    def loading_animation(stop_event, text="🔎 Procurando itens"):
        """Animação de carregamento"""
        dots = 0
        while not stop_event.is_set():
            print(f"\r{text}{'.' * dots}   ", end="", flush=True)
            dots = (dots + 1) % 4
            time.sleep(0.5)
    
    @staticmethod
    def tui_main_menu():
        """Exibe menu principal"""
        print(Fore.YELLOW + Style.BRIGHT + "OPÇÕES DISPONÍVEIS:")
        print(Fore.CYAN + "Ctrl + C para voltar / sair\n")
        
        pergunta = [{
            "type": "list",
            "name": "opcao",
            "message": "",
            "choices": [
                {"name": " Buscar por Número de Tombamento", "value": "1"},
                {"name": " Buscar por Número de Patrimônio", "value": "2"},
                {"name": " Buscar por Número de Inventário", "value": "3"},
                {"name": " Buscar por Especificação", "value": "4"},
                {"name": " Verificar itens do relatório PDF", "value": "5"},
                {"name": " Sair", "value": "0"},
            ],
        }]
        
        try:
            resposta = prompt(pergunta)
            return resposta["opcao"]
        except KeyboardInterrupt:
            return None
    
    @staticmethod
    def tui_select_pdf(pdf_files):
        """Seleção de PDFs"""
        print(Fore.YELLOW + Style.BRIGHT + "RELATÓRIOS DISPONÍVEIS:")
        print(Fore.CYAN + "Ctrl + C para voltar\n")
        
        pergunta = [{
            "type": "list",
            "name": "pdf",
            "message": "",
            "choices": pdf_files,
        }]
        
        try:
            resposta = prompt(pergunta)
            return resposta["pdf"]
        except KeyboardInterrupt:
            return None
    
    @staticmethod
    def display_search_results(item_data, sala, criterion, original_value):
        """Exibe resultados da busca"""
        if criterion == "especificacao":
            print(Fore.MAGENTA + Style.BRIGHT + 
                  f"\n🔎 Procurando por: {Fore.YELLOW}{original_value}\n")
        else:
            print(Fore.MAGENTA + Style.BRIGHT + 
                  f"\n🔎 Procurando pelo item com número de "
                  f"{Fore.YELLOW}{criterion.upper()}: {original_value}\n")
        
        print(Fore.YELLOW + Style.BRIGHT + f"📂 Vasculhando pasta: {item_data['folder']}")
        
        if not os.path.exists(item_data['folder']):
            print(Fore.RED + "  Pasta não encontrada.\n")
            return
        
        print("\n")
        print(Fore.WHITE + "Origem: " + item_data['origin'])
        print(Fore.WHITE + f"Sala (arquivo): {item_data['file']}")
        print(Fore.WHITE + f"Aba: {item_data['sheet']}")
        print(Fore.WHITE + f"Linha: {item_data['row']}")
        
        View.print_line()
        print(Fore.CYAN + f"Item: {item_data['item']}")
        print(Fore.CYAN + f"Tombamento: {item_data['tombamento']}")
        print(Fore.CYAN + f"Patrimônio: {item_data['patrimonio']}")
        print(Fore.CYAN + f"Inventário: {item_data['inventario']}")
        print(Fore.CYAN + f"Especificação: {item_data['especificacao']}")
        print(Fore.CYAN + f"TR: {item_data['tr']}")
        print(Fore.CYAN + f"Situação: {item_data['situacao']}")
        View.print_line()
    
    @staticmethod
    def display_report_results(itens_por_sala):
        """Exibe resultados da verificação de relatório"""
        for sala, itens in itens_por_sala.items():
            print(Fore.CYAN + Style.BRIGHT + f"\n📍 ITENS DA SALA - {sala}")
            View.print_line()
            
            for item in itens:
                if item["status"]:
                    print(Fore.GREEN + Style.BRIGHT +
                          f"✔ Tombamento: {item['tombamento']} | Item: {item['denominacao']}")
                else:
                    print(Fore.RED + Style.BRIGHT +
                          f"✖ Tombamento: {item['tombamento']} | Item: {item['denominacao']}")

# ==================================================
# CLASSE: MODEL (MANIPULAÇÃO DE DADOS E ARQUIVOS)
# ==================================================

class Model:
    """Responsável por manipulação de dados e arquivos"""
    
    @staticmethod
    def extract_room_from_filename(filename):
        """Extrai código + nome da sala a partir do nome do arquivo"""
        name = os.path.splitext(filename)[0]
        name = re.sub(r"\(.*?\)", "", name)
        name = name.strip()
        parts = [p.strip() for p in name.split("_") if p.strip()]
        
        if not parts:
            return name
        
        room_parts = []
        
        for i, part in enumerate(parts):
            upper = part.upper()
            
            if i > 0 and any(keyword in upper for keyword in [
                "CEDUC", "NEOA", "SUPORTE", "GABINETE", "PROF", "DOCENTE",
                "AL", "ANA", "CAROLINA", "RODRIGUES", "OLIVEIRA"
            ]):
                break
            
            room_parts.append(part)
        
        return " ".join(room_parts)
    
    @staticmethod
    def list_pdf_reports():
        """Lista arquivos PDF na pasta de relatórios"""
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
        
        return pdf_files
    
    @staticmethod
    def load_spreadsheet_data():
        """Carrega todos os tombamentos das planilhas"""
        found_tombamentos = {}
        
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
                            sala = Model.extract_room_from_filename(file)
                            found_tombamentos[tombamento_cell] = sala
        
        return found_tombamentos
    
    @staticmethod
    def draw_check(page, x, y, size, color, width=1.5):
        """Desenha um ✓ vetorial usando 2 linhas"""
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
    
    @staticmethod
    def draw_x(page, x, y, size, color, width=1.5):
        """Desenha um ✗ vetorial usando 2 linhas"""
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
    
    @staticmethod
    def generate_checked_pdf(original_pdf, output_pdf, tombamento_results):
        """Gera uma cópia do PDF com ícones vetoriais"""
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
                        Model.draw_check(page, ICON_X, center_y, ICON_SIZE, color, STROKE)
                    else:
                        Model.draw_x(page, ICON_X, center_y, ICON_SIZE, color, STROKE)
        
        doc.save(output_pdf)
        doc.close()

# ==================================================
# CLASSE: CONTROLLER (LÓGICA DE NEGÓCIO)
# ==================================================

class Controller:
    """Responsável pela lógica de negócio e coordenação"""
    
    def __init__(self):
        self.view = View()
        self.model = Model()
        self.GOOD_FILL = PatternFill(fill_type="solid", fgColor="FFC6EFCE")
        self.BAD_FILL = PatternFill(fill_type="solid", fgColor="FFFFC7CE")
    
    def search_items(self, column_index, value, criterion):
        """Busca itens nas planilhas"""
        original_value = value.strip()
        search_value = original_value.upper()
        criterion = criterion.lower()
        found = False
        
        for folder in SPREADSHEET_FOLDERS:
            item_data = {
                'folder': folder,
                'origin': '',
                'file': '',
                'sheet': '',
                'row': '',
                'item': '',
                'tombamento': '',
                'patrimonio': '',
                'inventario': '',
                'especificacao': '',
                'tr': '',
                'situacao': ''
            }
            
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
                                item_data['origin'] = Fore.GREEN + Style.BRIGHT + "NEOA"
                            elif "CEDUC" in folder.upper():
                                item_data['origin'] = Fore.BLUE + Style.BRIGHT + "CEDUC"
                            else:
                                item_data['origin'] = Fore.WHITE + "DESCONHECIDA"
                            
                            item_data.update({
                                'file': file,
                                'sheet': sheet_name,
                                'row': row_index,
                                'item': row[1],
                                'tombamento': row[2],
                                'patrimonio': row[3],
                                'inventario': row[4],
                                'especificacao': row[5],
                                'tr': row[6],
                                'situacao': row[7]
                            })
                            
                            self.view.display_search_results(item_data, None, criterion, original_value)
        
        print(Fore.MAGENTA + Style.BRIGHT + "\nVarredura finalizada.")
        
        if not found:
            print(Fore.RED + "✖ Nenhum resultado encontrado.")
    
    def search_items_from_pdf(self, pdf_path):
        """Processa itens de um relatório PDF"""
        os.makedirs(VERIFIED_REPORTS_FOLDER, exist_ok=True)
        
        stop_event = threading.Event()
        loader = threading.Thread(
            target=self.view.loading_animation,
            args=(stop_event,)
        )
        loader.start()
        
        report_items = extract_items_from_pdf(pdf_path)
        
        if not report_items:
            print(Fore.RED + "Nenhum item encontrado no relatório.")
            return
        
        found_tombamentos = self.model.load_spreadsheet_data()
        
        print(Fore.MAGENTA + Style.BRIGHT + "\nVerificando itens do relatório:\n")
        
        tombamento_results = {}
        itens_por_sala = defaultdict(list)
        
        for item in report_items:
            tombamento = item["tombamento"]
            denominacao = item["denominacao"]
            
            if tombamento in found_tombamentos:
                sala = found_tombamentos[tombamento]
                tombamento_results[tombamento] = True
                
                itens_por_sala[sala].append({
                    "status": True,
                    "tombamento": tombamento,
                    "denominacao": denominacao
                })
            else:
                tombamento_results[tombamento] = False
                
                itens_por_sala["NÃO ENCONTRADO"].append({
                    "status": False,
                    "tombamento": tombamento,
                    "denominacao": denominacao
                })
        
        stop_event.set()
        loader.join()
        print("\r" + " " * 50 + "\r", end="")
        
        self.view.display_report_results(itens_por_sala)
        
        choice = input(
            Fore.YELLOW +
            "\nDeseja aplicar o resultado na planilha REAL de teste controlado? (s/n): "
        )
        
        if choice.lower() == "s":
            self.apply_results_to_spreadsheets(tombamento_results)
        
        # Gera PDF marcado
        original_name = os.path.splitext(os.path.basename(pdf_path))[0]
        output_filename = f"{original_name} - verificado.pdf"
        output_pdf = os.path.join(VERIFIED_REPORTS_FOLDER, output_filename)
        
        if os.path.exists(output_pdf):
            print(Fore.YELLOW + Style.BRIGHT +
                  f"\n📄 PDF já verificado anteriormente:"
                  f"\n➡ {output_pdf}")
        else:
            self.model.generate_checked_pdf(pdf_path, output_pdf, tombamento_results)
            print(Fore.CYAN + Style.BRIGHT +
                  f"\n📄 PDF gerado com marcações: {output_pdf}")
    
    def apply_results_to_spreadsheets(self, tombamento_results):
        """Aplica resultados às planilhas"""
        print(Fore.RED + Style.BRIGHT + "\n⚠️ MODO PRODUÇÃO ⚠️")
        print("Você está prestes a alterar TODAS as planilhas reais.")
        print("Essa ação NÃO pode ser desfeita automaticamente.\n")
        
        confirm = input("Digite APLICAR para continuar: ")
        
        if confirm.strip().upper() != "APLICAR":
            print(Fore.YELLOW + "Operação cancelada.")
            return
        
        total_files = 0
        total_changes = 0
        
        for folder in SPREADSHEET_FOLDERS:
            if not os.path.exists(folder):
                continue
            
            for file in os.listdir(folder):
                if not file.endswith(".xlsx"):
                    continue
                
                file_path = os.path.join(folder, file)
                
                try:
                    wb = load_workbook(file_path)
                except Exception:
                    continue
                
                file_changes = 0
                
                for sheet in wb.sheetnames:
                    ws = wb[sheet]
                    
                    for row in ws.iter_rows(min_row=2):
                        tombamento_cell = row[2]
                        
                        if tombamento_cell.value is None:
                            continue
                        
                        tombamento = str(tombamento_cell.value).strip()
                        
                        if tombamento in tombamento_results:
                            status = tombamento_results[tombamento]
                            target_cell = row[0]
                            
                            target_cell.fill = self.GOOD_FILL if status else self.BAD_FILL
                            file_changes += 1
                
                if file_changes > 0:
                    wb.save(file_path)
                    total_files += 1
                    total_changes += file_changes
                    
                    print(Fore.GREEN + f"✔ {file} — {file_changes} linhas atualizadas")
        
        print(Fore.CYAN + Style.BRIGHT +
              f"\n✅ PROCESSO FINALIZADO"
              f"\n📂 Planilhas alteradas: {total_files}"
              f"\n🧾 Linhas atualizadas: {total_changes}")
    
    def run_menu(self):
        """Loop principal do menu"""
        while True:
            self.view.clear_terminal()
            self.view.print_header("CONSULTA DE PATRIMÔNIO - CEDUC")
            
            option = self.view.tui_main_menu()
            
            if option is None or option == "0":
                self.view.clear_terminal()
                print(Fore.MAGENTA + Style.BRIGHT + "Programa encerrado. Até mais 👋")
                break
            
            if option == "5":
                self.view.clear_terminal()
                self.view.print_header("VERIFICAÇÃO DE ITENS DO RELATÓRIO PDF")
                
                pdf_files = self.model.list_pdf_reports()
                
                if not pdf_files:
                    input(self.view.highlight_key(
                        "\nPressione ENTER para voltar ao menu...",
                        "ENTER",
                        Fore.GREEN
                    ))
                    continue
                
                selected_pdf = self.view.tui_select_pdf(pdf_files)
                
                if selected_pdf is None:
                    continue
                
                pdf_path = os.path.join(REPORTS_FOLDER, selected_pdf)
                
                self.view.clear_terminal()
                self.view.print_header("RESULTADO DA VERIFICAÇÃO DO RELATÓRIO")
                self.search_items_from_pdf(pdf_path)
                
                input(self.view.highlight_key(
                    "\nPressione ENTER para voltar ao menu...",
                    "ENTER",
                    Fore.GREEN
                ))
                continue
            
            if option not in SEARCH_COLUMNS:
                print(Fore.RED + "\nOpção inválida.")
                input(self.view.highlight_key(
                    "\nPressione ENTER para continuar...",
                    "ENTER",
                    Fore.YELLOW
                ))
                continue
            
            criterion_name, column_index = SEARCH_COLUMNS[option]
            value = input(Fore.YELLOW + f"Digite o valor para {criterion_name.upper()}: ")
            
            self.view.clear_terminal()
            self.view.print_header(f"RESULTADO DA BUSCA - {criterion_name.upper()}")
            self.search_items(column_index, value, criterion_name)
            
            input(self.view.highlight_key(
                "\nPressione ENTER para voltar ao menu...",
                "ENTER",
                Fore.GREEN
            ))

# ==================================================
# PONTO DE ENTRADA DA APLICAÇÃO
# ==================================================

if __name__ == "__main__":
    app = Controller()
    app.run_menu()
