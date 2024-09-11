import os
import re
import win32com.client as win32
from Front import iniciar_interface, obter_campos
import ttkbootstrap as ttk
from ttkbootstrap.constants import INFO, SUCCESS, PRIMARY, DANGER
from tkinter import filedialog, messagebox

# Expressão regular para validar e-mails
EMAIL_REGEX = r"^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$"

# Função para validar e-mails
def validar_emails(emails_raw, label_status):
    emails = [email.strip() for email in emails_raw.split(',') if email.strip()]
    all_valid = all(re.match(EMAIL_REGEX, email) for email in emails)

    if all_valid:
        label_status.config(text="Todos os e-mails são válidos!", bootstyle=SUCCESS)
    else:
        label_status.config(text="Um ou mais e-mails são inválidos!", bootstyle=DANGER)

# Função para converter o arquivo em PDF (usando o Word)
def converter_para_pdf(documento_path):
    word = win32.Dispatch('Word.Application')
    doc = word.Documents.Open(documento_path)
    pdf_path = documento_path.replace(".docx", ".pdf")
    doc.SaveAs(pdf_path, FileFormat=17)  # 17 é o formato PDF no Word
    doc.Close()
    word.Quit()
    return pdf_path

# Função para enviar o e-mail
def enviar_email(visao=True):
    # Obter valores dos campos de entrada da interface
    emails_raw, assunto, corpo, file_path, converter_pdf = obter_campos()

    # Verifica se o caminho do arquivo é válido
    if file_path and not os.path.isfile(file_path):
        raise FileNotFoundError("O caminho do arquivo não é válido. Por favor, selecione um arquivo existente.")

    # Verificar se o arquivo é um PDF e a caixa de conversão está marcada
    if converter_pdf and file_path.endswith('.pdf'):
        messagebox.showerror("Erro de Conversão", "O arquivo já está em formato PDF. A conversão não é necessária.")
        return  # Interrompe a execução da função

    # Enviar e-mails individualmente
    emails = [email.strip() for email in emails_raw.split(',') if email.strip()]

    # Validar e-mails antes de enviar
    if not all(re.match(EMAIL_REGEX, email) for email in emails):
        messagebox.showerror("Erro de Validação", "Um ou mais e-mails são inválidos. Corrija antes de enviar.")
        return  # Interrompe a execução da função

    # Converter para PDF se a opção estiver marcada e o arquivo for DOCX
    if converter_pdf and file_path.endswith('.docx'):
        file_path = converter_para_pdf(file_path)

    # Inicia a aplicação do Outlook
    outlook = win32.Dispatch('outlook.application')

    for email_to in emails:
        # Cria um novo e-mail
        mail = outlook.CreateItem(0)
        mail.To = email_to
        mail.Subject = assunto
        mail.Body = corpo

        # Anexa o arquivo se o caminho for válido
        if file_path:
            mail.Attachments.Add(file_path)

        if visao:
            # Exibe o e-mail para revisão
            mail.Display()
        else:
            # Envia o e-mail diretamente
            mail.Send()

        print(f"E-mail {'preparado para' if visao else 'enviado para'}: {email_to}")

    print("Todos os e-mails foram processados com sucesso!")

# Função para escolher o arquivo (PDF ou DOCX)
def escolher_arquivo():
    file_path = filedialog.askopenfilename(filetypes=[("Todos os arquivos", "*.*"), ("PDF Files", "*.pdf"), ("Word Documents", "*.docx")])
    if file_path:
        entry_pdf.delete(0, ttk.END)
        entry_pdf.insert(0, file_path)

# Função para iniciar a interface gráfica
def iniciar_interface(enviar_email_callback):
    global entry_emails, entry_assunto, entry_corpo, entry_pdf, check_convert_pdf

    app = ttk.Window(themename="superhero")  # Usando um tema moderno
    app.title("Envio Automático de E-mails")
    
    # Definindo o tamanho da janela (ajustado para 750x600)
    app.geometry("750x600")
    app.resizable(True, True)

    # Título principal
    ttk.Label(app, text="Envio Automático de E-mails", font="Helvetica 16 bold", bootstyle=INFO).pack(pady=20)

    # Campo de e-mails
    ttk.Label(app, text="Endereço de e-mails (separe por vírgula):", font="Helvetica 12 bold").pack(anchor="w", padx=20)
    entry_emails = ttk.Text(app, height=3, width=60)
    entry_emails.pack(pady=10, padx=20, fill='x')

    # Label para mostrar o status de validação dos e-mails
    label_status = ttk.Label(app, text="", font="Helvetica 10", bootstyle=INFO)
    label_status.pack(anchor="w", padx=20, pady=10)

    # Vincular a função de validação ao evento de digitação no campo de e-mails
    entry_emails.bind("<KeyRelease>", lambda event: validar_emails(entry_emails.get("1.0", "end-1c"), label_status))

    # Campo de assunto
    ttk.Label(app, text="Assunto do E-mail:", font="Helvetica 12 bold").pack(anchor="w", padx=20)
    entry_assunto = ttk.Entry(app, width=62)
    entry_assunto.pack(pady=10, padx=20, fill='x')

    # Campo do corpo do e-mail
    ttk.Label(app, text="Corpo do E-mail:", font="Helvetica 12 bold").pack(anchor="w", padx=20)
    entry_corpo = ttk.Text(app, height=6, width=60)
    entry_corpo.pack(pady=10, padx=20, fill='x')

    # Campo para selecionar o arquivo
    ttk.Label(app, text="Anexo (PDF ou DOCX):", font="Helvetica 12 bold").pack(anchor="w", padx=20)
    entry_pdf = ttk.Entry(app, width=53)
    entry_pdf.pack(anchor="w", padx=20, fill='x')
    button_pdf = ttk.Button(app, text="Selecionar Arquivo", command=escolher_arquivo, bootstyle=INFO)
    button_pdf.pack(anchor="w", padx=20, pady=10)

    # Caixa de seleção para converter o arquivo em PDF (opcional)
    check_convert_pdf = ttk.IntVar()
    ttk.Checkbutton(app, text="Converter para PDF", variable=check_convert_pdf, bootstyle=PRIMARY).pack(anchor="w", padx=20, pady=10)

    # Criando um frame para os botões com grid layout
    button_frame = ttk.Frame(app)
    button_frame.pack(pady=30, fill='x')

    # Botão "Enviar com Visualização"
    button_enviar_com_visualizacao = ttk.Button(button_frame, text="Enviar com Visualização", command=lambda: enviar_email_callback(visao=True), bootstyle=SUCCESS, width=25)
    button_enviar_com_visualizacao.grid(row=0, column=0, padx=20, pady=10)

    # Botão "Enviar sem Visualização"
    button_enviar_sem_visualizacao = ttk.Button(button_frame, text="Enviar sem Visualização", command=lambda: enviar_email_callback(visao=False), bootstyle=DANGER, width=25)
    button_enviar_sem_visualizacao.grid(row=0, column=1, padx=20, pady=10)

    app.mainloop()

# Função principal para iniciar a aplicação
if __name__ == "__main__":
    iniciar_interface(enviar_email)
