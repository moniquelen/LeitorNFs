<div align="center">
<img alt="logo_leitor_nf" src="static\img\imgLogo.png" width="20%">
<h1>Leitor de Notas Fiscais</h1>
</div>

<div>
<h3 align="center">Projeto para extração automática de dados de Notas Fiscais em PDF.</h3>
<p>O Leitor de NFs automatiza a leitura de arquivos PDF de notas fiscais para extrair informações como Número da Nota, Valor, Data de Emissão e CNPJ. Os dados extraídos são organizados em uma planilha Excel. O projeto também realiza a conversão dos PDFs em XML para facilitar o tratamento dos dados durante o processo de extração.</p>
</div>

## ⚙️ Como usar

1. **Clone o repositório**:
   ```bash
   git clone https://github.com/seu-usuario/leitor-de-nfs.git
   cd leitor-de-nfs
   ````
2. **Instale as dependências**: 

    O programa requer as bibliotecas Python PyMuPDF e Pandas, então ainda caso não tenha, é necessário instalá-las:
    ```bash
    pip install pymupdf pandas
    ```
3. **Execute o script**:
    ```bash
    python main.py
    ```
   
## 💻 Interface e informações

<img src="static\media\imgInterface1.png" width="100%">

Após anexar os PDFs, o botão "Processar NFs" se tornará clicável e através dele é executada a leitura da notas fiscais.

- Quando a leitura das notas for concluída, aparecerá o pop-up informativo abaixo:

<img src="static\media\imgInterface2.png" width="100%">

Após isso, é possível salvar o arquivo Excel que foi gerado, por meio do botão "Exportar Excel".

- Quando o Excel for salvo, também aparecerá um pop-up informativo, e o Excel poderá ser aberto clicando no botão "Abrir Relatório"

<img src="static\media\imgInterface3.png" width="100%">

---
<p align="center">Feito com ♥ by Monique.</p>