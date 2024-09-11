import ttkbootstrap as ttk
from ttkbootstrap.constants import INFO, SUCCESS, PRIMARY, DANGER
from tkinter import filedialog, messagebox

# Variáveis globais para os campos de entrada
entry_emails = None
entry_assunto = None
entry_corpo = None
entry_pdf = None
check_convert_pdf = None

# Função para obter os campos de entrada da interface
def obter_campos():
    emails_raw = entry_emails.get("1.0", "end-1c")
    assunto = entry_assunto.get()
    corpo = entry_corpo.get("1.0", "end-1c")
    file_path = entry_pdf.get()
    converter_pdf = check_convert_pdf.get()
    return emails_raw, assunto, corpo, file_path, converter_pdf

# Função para escolher o arquivo (PDF ou DOCX)
def escolher_arquivo():
    file_path = filedialog.askopenfilename(filetypes=[("Todos os arquivos", "*.*"), ("PDF Files", "*.pdf"), ("Word Documents", "*.docx")])
    if file_path:
        entry_pdf.delete(0, ttk.END)
        entry_pdf.insert(0, file_path)
        messagebox.showinfo("Arquivo Selecionado", f"Arquivo selecionado: {file_path}")

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
