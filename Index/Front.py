import ttkbootstrap as ttk
from ttkbootstrap.constants import INFO, SUCCESS, PRIMARY, DANGER
from tkinter import filedialog, messagebox
import win32com.client as win32  # Biblioteca para integração com o Outlook
from Documentos import processar_documento

# Função para criar a aba de envio de e-mails
def criar_aba_email(notebook):
    global entry_emails, entry_enviar_de, entry_assunto, entry_corpo  # Declarando variáveis globais
    frame_email = ttk.Frame(notebook)
    notebook.add(frame_email, text="Envio de E-mails")

    ttk.Label(frame_email, text="Envio Automático de E-mails", font="Helvetica 16 bold", bootstyle=INFO).pack(pady=20)

    ttk.Label(frame_email, text="Endereço de e-mails (separe por vírgula):", font="Helvetica 12 bold").pack(anchor="w", padx=20)
    entry_emails = ttk.Text(frame_email, height=3, width=60)
    entry_emails.pack(pady=10, padx=20, fill='x')

    ttk.Label(frame_email, text="Enviar De (opcional):", font="Helvetica 12 bold").pack(anchor="w", padx=20)
    entry_enviar_de = ttk.Entry(frame_email, width=62)
    entry_enviar_de.pack(pady=10, padx=20, fill='x')

    ttk.Label(frame_email, text="Assunto do E-mail:", font="Helvetica 12 bold").pack(anchor="w", padx=20)
    entry_assunto = ttk.Entry(frame_email, width=62)
    entry_assunto.pack(pady=10, padx=20, fill='x')

    ttk.Label(frame_email, text="Corpo do E-mail:", font="Helvetica 12 bold").pack(anchor="w", padx=20)
    entry_corpo = ttk.Text(frame_email, height=6, width=60)
    entry_corpo.pack(pady=10, padx=20, fill='x')

    ttk.Label(frame_email, text="Assinatura (opcional):", font="Helvetica 12 bold").pack(anchor="w", padx=20)
    combo_assinaturas = ttk.Combobox(frame_email, values=["Nenhuma"], width=60)
    combo_assinaturas.set("Nenhuma")
    combo_assinaturas.pack(pady=10, padx=20, fill='x')

    check_importante = ttk.IntVar()
    ttk.Checkbutton(frame_email, text="Marcar como importante", variable=check_importante, bootstyle=PRIMARY).pack(anchor="w", padx=20, pady=10)

    check_confirmacao_entrega = ttk.IntVar()
    ttk.Checkbutton(frame_email, text="Solicitar confirmação de entrega", variable=check_confirmacao_entrega, bootstyle=PRIMARY).pack(anchor="w", padx=20, pady=10)

    button_frame = ttk.Frame(frame_email)
    button_frame.pack(pady=30, fill='x')

    button_enviar_com_visualizacao = ttk.Button(button_frame, text="Enviar com Visualização", 
                                                command=lambda: enviar_email(visao=True),  # Enviar com visualização
                                                bootstyle=SUCCESS, width=25)
    button_enviar_com_visualizacao.grid(row=0, column=0, padx=20, pady=10)

    button_enviar_sem_visualizacao = ttk.Button(button_frame, text="Enviar sem Visualização", 
                                                command=lambda: enviar_email(visao=False),  # Enviar sem visualização
                                                bootstyle=DANGER, width=25)
    button_enviar_sem_visualizacao.grid(row=0, column=1, padx=20, pady=10)

# Função para enviar e-mails via Outlook
def enviar_email(visao=True):
    # Obter valores dos campos da interface
    emails_raw = obter_emails()  # Chama a função obter_emails()
    assunto = entry_assunto.get().strip()
    corpo = entry_corpo.get("1.0", "end-1c").strip()
    email_enviar_de = entry_enviar_de.get().strip()

    # Verificar se os campos essenciais estão preenchidos
    if not emails_raw:
        messagebox.showerror("Erro", "O campo de e-mails não pode estar vazio.")
        return

    # Processar os e-mails separados por vírgula
    emails = [email.strip() for email in emails_raw.split(',') if email.strip()]

    # Validar se pelo menos um e-mail foi inserido
    if not emails:
        messagebox.showerror("Erro", "Por favor, insira pelo menos um e-mail válido.")
        return

    try:
        # Iniciar o Outlook
        outlook = win32.Dispatch('outlook.application')

        for email_to in emails:
            # Criar um novo e-mail
            mail = outlook.CreateItem(0)
            mail.To = email_to
            mail.Subject = assunto
            mail.Body = corpo

            # Definir o remetente alternativo, se fornecido
            if email_enviar_de:
                mail.SentOnBehalfOfName = email_enviar_de

            if visao:
                # Exibe o e-mail no Outlook para visualização
                mail.Display()
            else:
                # Envia diretamente o e-mail
                mail.Send()

            print(f"E-mail {'preparado para' if visao else 'enviado para'}: {email_to}")

        messagebox.showinfo("Sucesso", "Todos os e-mails foram processados com sucesso!")

    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao enviar e-mails: {e}")

# Função para obter os e-mails digitados na aba de envio
def obter_emails():
    return entry_emails.get("1.0", "end-1c").strip()

# Função para criar a aba de transformação de documento
def criar_aba_transformar_documento(notebook):
    frame_transformar = ttk.Frame(notebook)
    notebook.add(frame_transformar, text="Transformar Documento")

    ttk.Label(frame_transformar, text="Transformar Documento", font="Helvetica 16 bold", bootstyle=PRIMARY).pack(pady=20)

    ttk.Label(frame_transformar, text="Anexo (DOCX) - Opcional:", font="Helvetica 12 bold").pack(anchor="w", padx=20)
    entry_pdf = ttk.Entry(frame_transformar, width=50)
    entry_pdf.pack(pady=10, padx=20, fill='x')

    button_pdf = ttk.Button(frame_transformar, text="Selecionar Arquivo", command=lambda: selecionar_arquivo(entry_pdf), bootstyle=PRIMARY)
    button_pdf.pack(pady=10, padx=20)

    check_convert_pdf = ttk.IntVar()
    ttk.Checkbutton(frame_transformar, text="Converter para PDF", variable=check_convert_pdf, bootstyle=PRIMARY).pack(anchor="w", padx=20, pady=10)

    # Adicionar o caminho da planilha para o argumento de processar_documento
    caminho_planilha = r"C:\Users\EduardaGoes\OneDrive - LAQUS\Documentos\LaqusProjetos\Base de Dados Comercial\Base de Dados Bancos.xlsx"

    # Botão para processar o documento (anexo opcional)
    button_processar = ttk.Button(frame_transformar, text="Processar Documento", 
                                  command=lambda: processar_documento(obter_emails(), entry_pdf.get() if entry_pdf.get() else None, caminho_planilha), 
                                  bootstyle=PRIMARY)
    button_processar.pack(pady=10)

# Função para selecionar o arquivo
def selecionar_arquivo(entry_pdf):
    file_path = filedialog.askopenfilename(filetypes=[("Word Documents", "*.docx")])
    if file_path:
        entry_pdf.delete(0, ttk.END)
        entry_pdf.insert(0, file_path)

# Função principal para iniciar a interface gráfica
def iniciar_interface():
    app = ttk.Window(themename="superhero")  # Usando um tema moderno
    app.title("Envio de E-mails e Transformação de Documento")
    app.geometry("800x600")
    app.resizable(True, True)

    # Criar o Notebook (para as abas)
    notebook = ttk.Notebook(app)
    notebook.pack(pady=10, expand=True, fill='both')

    # Criação da aba de envio de e-mails
    criar_aba_email(notebook)

    # Criação da aba de transformação de documento
    criar_aba_transformar_documento(notebook)

    app.mainloop()
