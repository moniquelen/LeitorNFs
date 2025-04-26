import os
import re
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

def formatar_cte_number(cte_number):
    if cte_number:
        cte_number = re.sub(r'[^\d]', '', cte_number) 
        cte_number = cte_number.lstrip('0')
    return cte_number

def extrair_relevantes_com_regex(texto):
    texto = re.sub(r'\s+', ' ', texto)

    padrao_numero_nf = r'N[°º]\s?[.]?\s?\d{1,10}(?:\.\d{3})*|[0-9]{3}\.[0-9]{3}\.[0-9]{3}'
    padrao_data_emissao = r'(0[1-9]|[12][0-9]|3[01])\/(0[1-9]|1[0-2])\/[0-9]{4}'
    padrao_valor_cte = r'(?:VALOR TOTAL DA PRESTAÇÃO DO SERVIÇO|VALOR TOTAL DO SERVIÇO)[^\d\n]*(\d[\d.,]*)'
    padrao_valor_receber = r'VALOR A RECEBER[^\d\n]*(\d[\d.,]*)'

    cte_number = re.search(padrao_numero_nf, texto)
    emission_date = re.search(padrao_data_emissao, texto)
    cte_value_match = re.search(padrao_valor_cte, texto)
    valor_receber_match = re.search(padrao_valor_receber, texto)

    cte_value = float(cte_value_match.group(1).replace('.', '').replace(',', '.')) if cte_value_match else 0
    valor_receber = float(valor_receber_match.group(1).replace('.', '').replace(',', '.')) if valor_receber_match else 0

    cte_value_final = max(cte_value, valor_receber)

    return {
        "CTe Number": formatar_cte_number(cte_number.group(0)) if cte_number else "N/A",
        "CTe Value": f"{cte_value_final:,.2f}" if cte_value_final else "N/A",
        "Emission Date": emission_date.group(0) if emission_date else "N/A",
    }

def process_single_pdf(pdf_path, output_folder):
    print(f"[INFO] Processando arquivo: {os.path.basename(pdf_path)}")

    xml_path = converter_pdf_para_xml(pdf_path, output_folder)
    if not xml_path:
        print(f"[ERRO] Falha ao converter o PDF {os.path.basename(pdf_path)} para XML.")
        return [os.path.basename(pdf_path), "N/A", "N/A", "N/A"]

    texto = extrair_texto_do_xml(xml_path)
    if not texto:
        print(f"[ERRO] Nenhum texto extraído do XML gerado para {os.path.basename(pdf_path)}.")
        return [os.path.basename(pdf_path), "N/A", "N/A", "N/A"]

    info = extrair_relevantes_com_regex(texto)

    print(f"[INFO] Dados extraídos do arquivo {os.path.basename(pdf_path)}: {info}")

    return [os.path.basename(pdf_path), info["CTe Number"], info["CTe Value"], info["Emission Date"]]

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

    df = pd.DataFrame(data, columns=["Filename", "CTe Number", "CTe Value", "Emission Date"])
    df.to_excel(output_excel_path, index=False)
    print(f"[INFO] Processamento concluído. Dados salvos em: {output_excel_path}")

    # Limpar pastas
    limpar_pasta_xml(output_folder)
    limpar_pasta_nfs(folder_path)

if __name__ == "__main__":
    folder_path = r"C:\Users\S7U05260\OneDrive - Subsea7\Desktop\LeitorNFs\NFs"
    output_folder = r"C:\Users\S7U05260\OneDrive - Subsea7\Desktop\LeitorNFs\XMLs"
    output_excel_path = "output.xlsx"

    process_pdfs_in_folder(folder_path, output_folder, output_excel_path)
