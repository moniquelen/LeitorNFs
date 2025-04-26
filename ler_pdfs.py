import os
import re
from pathlib import Path
import fitz  # PyMuPDF
import pandas as pd
from concurrent.futures import ThreadPoolExecutor

def converter_pdf_para_xml(pdf_path, output_folder):
    try:
        os.makedirs(output_folder, exist_ok=True)
        xml_filename = os.path.splitext(os.path.basename(pdf_path))[0] + ".xml"
        xml_path = os.path.join(output_folder, xml_filename)

        doc = fitz.open(pdf_path)
        with open(xml_path, "w", encoding="utf-8") as xml_file:
            xml_file.write("<pdf>\n")
            for page_num, page in enumerate(doc, start=1):
                texto = page.get_text("text")
                xml_file.write(f"  <page number='{page_num}'>\n")
                xml_file.write(f"    <content><![CDATA[{texto}]]></content>\n")
                xml_file.write("  </page>\n")
            xml_file.write("</pdf>")

        return xml_path
    except Exception as e:
        print(f"[ERRO] Falha ao converter PDF para XML: {e}")
        return None

def extrair_texto_do_xml(xml_path):
    try:
        with open(xml_path, "r", encoding="utf-8") as xml_file:
            return xml_file.read()
    except Exception as e:
        print(f"[ERRO] Falha ao ler o arquivo XML: {e}")
        return ""

def extrair_valor(texto, padrao):
    match = re.search(padrao, texto, flags=re.IGNORECASE)
    if match:
        valor_str = match.group(1)
        valor_str = valor_str.replace('.', '').replace(',', '.').strip()
        try:
            return float(valor_str)
        except ValueError:
            print(f"[AVISO] Valor inválido para conversão: {valor_str}")
            return 0.0
    return 0.0

def extrair_relevantes_com_regex(texto):
    texto = re.sub(r'\s+', ' ', texto)

    padrao_numero_nf = r'N[°º]\s?[.]?\s?\d{1,10}(?:\.\d{3})*|[0-9]{3}\.[0-9]{3}\.[0-9]{3}'
    padrao_data_emissao = r'(0[1-9]|[12][0-9]|3[01])/(0[1-9]|1[0-2])/\d{4}'
    padrao_cnpj = r'\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}'

    padrao_valor_cte = (
        r'(?:VALOR TOTAL DA PRESTAÇÃO DO SERVIÇO|'
        r'VALOR TOTAL DO SERVIÇO|'
        r'VALOR TOTAL DA NOTA|'
        r'VALOR TOTAL DA NF|'
        r'VALOR LÍQUIDO|'
        r'VALOR LÍQUIDO DA NOTA|'
        r'VALOR A RECEBER)[^\d]*(\d[\d.,]*)'
    )

    cte_number_match = re.search(padrao_numero_nf, texto)
    emission_date_match = re.search(padrao_data_emissao, texto)
    cnpj_match = re.search(padrao_cnpj, texto)

    cte_value = extrair_valor(texto, padrao_valor_cte)

    return {
        "Numero da NF": cte_number_match.group(0).strip() if cte_number_match else "N/A",
        "Valor da NF": f"{cte_value:,.2f}" if cte_value else "N/A",
        "Data de Emissão": emission_date_match.group(0) if emission_date_match else "N/A",
        "CNPJ": cnpj_match.group(0) if cnpj_match else "N/A",
    }

def process_single_pdf(pdf_path, output_folder):
    print(f"[INFO] Processando arquivo: {os.path.basename(pdf_path)}")

    xml_path = converter_pdf_para_xml(pdf_path, output_folder)
    if not xml_path:
        print(f"[ERRO] Falha ao converter o PDF {os.path.basename(pdf_path)} para XML.")
        return [os.path.basename(pdf_path), "N/A", "N/A", "N/A", "N/A"]

    texto = extrair_texto_do_xml(xml_path)
    if not texto:
        print(f"[ERRO] Nenhum texto extraído do XML gerado para {os.path.basename(pdf_path)}.")
        return [os.path.basename(pdf_path), "N/A", "N/A", "N/A", "N/A"]

    info = extrair_relevantes_com_regex(texto)

    print(f"[INFO] Dados extraídos do arquivo {os.path.basename(pdf_path)}: {info}")

    return [os.path.basename(pdf_path), info["Numero da NF"], info["Valor da NF"], info["Data de Emissão"], info["CNPJ"]]

def limpar_pasta_xml(output_folder):
    """Remove todos os arquivos XML da pasta especificada."""
    try:
        for file in os.listdir(output_folder):
            if file.endswith(".xml"):
                os.remove(os.path.join(output_folder, file))
    except Exception as e:
        print(f"[ERRO] Falha ao limpar a pasta XML: {e}")

def limpar_pasta_nfs(folder_path):
    """Remove todos os arquivos da pasta especificada."""
    try:
        for file in os.listdir(folder_path):
            file_path = os.path.join(folder_path, file)
            if os.path.isfile(file_path):
                os.remove(file_path)
        print(f"[INFO] Todos os arquivos da pasta {folder_path} foram excluídos.")
    except Exception as e:
        print(f"[ERRO] Falha ao limpar a pasta NFs: {e}")

def process_pdfs_in_folder(folder_path, output_folder, output_excel_path):
    print("[INFO] Iniciando processamento dos PDFs...")
    files = [f for f in os.listdir(folder_path) if f.endswith(".pdf")]

    with ThreadPoolExecutor() as executor:
        data = list(executor.map(process_single_pdf, [os.path.join(folder_path, f) for f in files], [output_folder] * len(files)))

    df = pd.DataFrame(data, columns=["Arquivo", "Numero da NF", "Valor da NF", "Data de Emissão", "CNPJ"])
    df.to_excel(output_excel_path, index=False)
    print(f"[INFO] Processamento concluído. Dados salvos em: {output_excel_path}")

    # Limpar pastas
    limpar_pasta_xml(output_folder)
    limpar_pasta_nfs(folder_path)

if __name__ == "__main__":
    base_path = Path(__file__).parent

    folder_path = base_path / "NFs"
    output_folder = base_path / "XMLs"
    output_excel_path = base_path / "output.xlsx"

    print(folder_path)
    print(output_folder)
    print(output_excel_path)

    process_pdfs_in_folder(folder_path, output_folder, output_excel_path)
