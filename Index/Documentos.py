import pandas as pd
import re
from docx import Document
from tkinter import messagebox
import datetime

estado_nome_completo = {
    'AC': 'Acre', 'AL': 'Alagoas', 'AP': 'Amapá', 'AM': 'Amazonas', 'BA': 'Bahia',
    'CE': 'Ceará', 'DF': 'Distrito Federal', 'ES': 'Espírito Santo', 'GO': 'Goiás',
    'MA': 'Maranhão', 'MT': 'Mato Grosso', 'MS': 'Mato Grosso do Sul', 'MG': 'Minas Gerais',
    'PA': 'Pará', 'PB': 'Paraíba', 'PR': 'Paraná', 'PE': 'Pernambuco', 'PI': 'Piauí',
    'RJ': 'Rio de Janeiro', 'RN': 'Rio Grande do Norte', 'RS': 'Rio Grande do Sul',
    'RO': 'Rondônia', 'RR': 'Roraima', 'SC': 'Santa Catarina', 'SP': 'São Paulo',
    'SE': 'Sergipe', 'TO': 'Tocantins'
}

# Exemplo de chamada da função com o caminho da planilha correto
caminho_planilha = r"C:\Users\EduardaGoes\OneDrive - LAQUS\Documentos\LaqusProjetos\Base de Dados Comercial\Base de Dados Bancos.xlsx"


# Função para extrair o domínio do e-mail
def extrair_dominio_empresa(email):
    try:
        dominio = re.search(r'@([a-zA-Z0-9.-]+)', email).group(1)
        return dominio.lower()  # Força o domínio para minúsculas
    except AttributeError:
        messagebox.showerror("Erro", f"O e-mail {email} não é válido.")
        return None

# Função para ler a planilha e obter os dados do banco baseado no domínio do e-mail
def obter_dados_banco(dominio_email, caminho_planilha):
    """
    Lê a planilha e busca os dados do banco cujo domínio de e-mail corresponde ao domínio fornecido.
    """
    try:
        # Carregar a planilha Excel
        df = pd.read_excel(caminho_planilha, engine='openpyxl')
        
        # Normalizar os nomes das colunas para remover espaços e converter para minúsculas
        df.columns = df.columns.str.strip().str.lower()

        # Exibir todas as colunas lidas para diagnóstico
        print(f"Colunas lidas: {df.columns.tolist()}")

        # Verifica se a coluna 'e-mails' (em minúsculas) está na planilha
        if 'e-mails' not in df.columns:
            messagebox.showerror("Erro", "A coluna 'e-mails' não foi encontrada na planilha.")
            return None

        # Substituir valores NaN por string vazia
        df['e-mails'] = df['e-mails'].fillna('').str.lower()

        # Filtrar os dados da planilha para encontrar o banco com o domínio correspondente
        dados_banco = df[df['e-mails'].str.contains(dominio_email)]

        if dados_banco.empty:
            messagebox.showerror("Erro", f"Nenhum banco encontrado para o domínio {dominio_email}")
            return None
        
        # Retorna a primeira linha correspondente
        return dados_banco.iloc[0]

    except FileNotFoundError:
        messagebox.showerror("Erro", "A planilha de dados não foi encontrada.")
        return None
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao ler a planilha: {e}")
        return None

# Função para substituir as chaves no documento DOCX com os dados do banco
def substituir_chaves_no_documento(arquivo_docx, dados_banco):
    """
    Substitui as chaves no documento DOCX com base nos dados do banco e outras informações.
    """
    try:
        # Carregar o documento Word
        doc = Document(arquivo_docx)

        # Obter a data atual (DataProposta)
        data_atual = datetime.datetime.now().strftime("%d/%m/%Y")

        # Verificar se as chaves estão presentes no DataFrame e substituir valores NaN por uma string vazia
        substituicoes = {
            'CNPJ': str(dados_banco.get('cnpj', '')).strip(),
            'ENDEREÇO': str(dados_banco.get('endereço', '')).strip(),
            'BAIRRO': str(dados_banco.get('bairro', '')).strip(),
            'CEP': str(dados_banco.get('cep', '')).strip(),
            'MUNICÍPIO': str(dados_banco.get('municipio', '')).strip(),
            'UF': str(dados_banco.get('uf', '')).strip(),
            'NOME INSTITUIÇÃO': str(dados_banco.get('nome instituição', '')).strip(),
            '<<DataProposta>>': data_atual,  # Adiciona a data atual
            '<<NomeCliente>>': str(dados_banco.get('nome instituição', '')).strip(),
            '<<EndBanco>>': str(dados_banco.get('endereço', '')).strip(),
            '<<CEPBanco>>': str(dados_banco.get('cep', '')).strip(),
            '<<SiglaEstado>>': estado_nome_completo.get(str(dados_banco.get('uf', '')).strip().upper(), '')
        }

        # Iterar por todos os parágrafos e fazer as substituições
        for paragrafo in doc.paragraphs:
            for chave, valor in substituicoes.items():
                if chave in paragrafo.text:
                    paragrafo.text = paragrafo.text.replace(chave, valor)

        # Salvar o documento modificado com um novo nome
        novo_nome_arquivo = arquivo_docx.replace('.docx', f'_transformado.docx')
        doc.save(novo_nome_arquivo)
        return novo_nome_arquivo

    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao processar o documento: {str(e)}")
        return None

# Função para processar o arquivo DOCX com base no e-mail
def processar_documento(emails, arquivo_docx, caminho_planilha):
    """
    Processa o documento com base no domínio do e-mail fornecido e os dados da planilha.
    """
    email = emails.split(',')[0].strip()  # Pega o primeiro e-mail da lista
    dominio = extrair_dominio_empresa(email)
    if not dominio:
        return

    # Obter os dados do banco com base no domínio do e-mail
    dados_banco = obter_dados_banco(dominio, caminho_planilha)
    
    # Aqui usamos `.empty` para verificar se o DataFrame está vazio
    if dados_banco is None or dados_banco.empty:
        return

    if not arquivo_docx:
        messagebox.showerror("Erro", "Nenhum arquivo .docx foi selecionado.")
        return

    # Substituir as chaves no documento DOCX
    novo_arquivo = substituir_chaves_no_documento(arquivo_docx, dados_banco)

    # Mostrar mensagem de sucesso
    if novo_arquivo:
        messagebox.showinfo("Sucesso", f"O documento foi transformado e salvo como: {novo_arquivo}")
    else:
        messagebox.showerror("Erro", "Não foi possível transformar o documento.")

