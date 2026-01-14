import pdfplumber
import re


def extract_items_from_pdf(pdf_path):
    items = []

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()

            if not text:
                continue

            lines = text.split("\n")

            for line in lines:
                line = line.strip()

                # Ignora linhas vazias ou de número de série
                if not line or line.lower().startswith("número de série"):
                    continue

                # Ex: 2017004687 TELEFONE IP 23/11/2017 ...
                match = re.match(
                    r"^(\d{6,})\s+(.+?)\s+\d{2}/\d{2}/\d{4}",
                    line
                )

                if match:
                    tombamento = match.group(1)
                    denominacao = match.group(2).strip()

                    items.append({
                        "tombamento": tombamento,
                        "denominacao": denominacao
                    })

    return items
