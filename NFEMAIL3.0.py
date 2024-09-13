import tkinter as tk
from tkinter import ttk
import os
import shutil
import threading
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# Função banco de dados txt
def carregar_dados():
    dados = {}
    try:
        with open('databasenf.txt', 'r', encoding='utf-8') as f:
            for linha in f:
                linha = linha.strip()
                if linha and "=" in linha:
                    chave, valor = linha.split('=', 1)
                    dados[chave.strip()] = valor.strip()
    except FileNotFoundError:
        print("Erro: O arquivo 'databasenf.txt' não foi encontrado.")
    except Exception as e:
        print(f"Erro ao carregar dados: {str(e)}")
    
    return dados


# Função para enviar o e-mail
def enviar_email(arquivos, destinatario):
    dados = carregar_dados()
    servidor_smtp = dados.get("servidor_smtp")
    porta_smtp = int(dados.get("porta_smtp", 25))
    usuario = dados.get("usuario")
    senha = dados.get("senha")
    
    # Configuração do e-mail
    mensagem = MIMEMultipart()
    mensagem["From"] = usuario
    mensagem["To"] = destinatario
    mensagem["Subject"] = "REENVIO XML NFE"
    
    corpo = "Os arquivos XML estão anexados:\n\n"
    mensagem.attach(MIMEText(corpo, "plain"))
    
    # Anexar os arquivos
    for arquivo in arquivos:
        anexo = MIMEBase("application", "octet-stream")
        with open(arquivo, "rb") as f:
            anexo.set_payload(f.read())
        encoders.encode_base64(anexo)
        anexo.add_header("Content-Disposition", f"attachment; filename={os.path.basename(arquivo)}")
        mensagem.attach(anexo)
    
    try:
        with smtplib.SMTP(servidor_smtp, porta_smtp) as server:
            server.login(usuario, senha)
            server.sendmail(usuario, destinatario, mensagem.as_string())
        status_var.set("Arquivos enviados por e-mail com sucesso!")
    except Exception as e:
        status_var.set(f"Erro ao enviar e-mail: {str(e)}")


# Função para copiar arquivos
def copiar_arquivos(numeros_notas, destinatario):
    dados = carregar_dados()
    origem_autorizadas = dados.get("origem_autorizadas")
    origem_nao_autorizadas = dados.get("origem_nao_autorizadas")
    destino = dados.get("destino")
    
    notas = [nota.strip() for nota in numeros_notas.split(',')]
    arquivos_copiados = []

    try:
        if not os.path.exists(destino):
            raise Exception(f"Diretório de destino {destino} não existe!")

        def procura_arquivos(diretorio):
            arquivos_encontrados = 0
            for subdir, _, arquivos in os.walk(diretorio):
                for arquivo in arquivos:
                    if arquivo.lower().endswith('.xml'):
                        caminho_arquivo = os.path.join(subdir, arquivo)
                        with open(caminho_arquivo, 'r', encoding='utf-8') as file:
                            conteudo = file.read()
                            for nota in notas:
                                if f'<nNF>{nota}</nNF>' in conteudo:
                                    destino_arquivo = os.path.join(destino, arquivo)
                                    shutil.copy(caminho_arquivo, destino_arquivo)
                                    arquivos_copiados.append(destino_arquivo)
                                    arquivos_encontrados += 1
                                    status_var.set(f"Copiando nota {nota}...")
                                    app.update_idletasks()  # Atualiza a interface gráfica
            return arquivos_encontrados

        total_autorizadas = procura_arquivos(origem_autorizadas)
        total_nao_autorizadas = procura_arquivos(origem_nao_autorizadas)

        total_arquivos = total_autorizadas + total_nao_autorizadas

        if total_arquivos == 0:
            status_var.set("Nenhum arquivo foi encontrado!")
        else:
            status_var.set(f"Processo concluído! {total_arquivos} arquivos copiados.")
            enviar_email(arquivos_copiados, destinatario)  # Enviar e-mail com os arquivos copiados

    except Exception as e:
        status_var.set(f"Erro: {str(e)}")
    finally:
        botao_iniciar.config(state="normal")  # Reativar o botão
        entrada_nota.config(state="normal")  # Reativar o campo de entrada
        entrada_email.config(state="normal")  # Reativar o campo de entrada de e-mail


# Função para iniciar a cópia
def iniciar_copia():
    botao_iniciar.config(state="disabled")  # Desativar o botão
    entrada_nota.config(state="disabled")  # Desativar o campo de entrada
    entrada_email.config(state="disabled")  # Desativar o campo de entrada de e-mail
    status_var.set("Iniciando a cópia...")
    app.update_idletasks()  # Atualiza a interface gráfica
    thread = threading.Thread(target=copiar_arquivos, args=(entrada_nota.get(), entrada_email.get()))
    thread.start()

# Função para validar a entrada (somente números)
def validar_entrada(char):
    return char.isdigit() or char == ','

# Função para validar o e-mail
def validar_email(char):
    return char.isalnum() or char in ['.', '@', '_', '-']

# Configuração da interface gráfica
app = tk.Tk()
app.title("Cópia de XML")
app.geometry("600x460")  # Ajuste da altura
app.resizable(False, False)
app.configure(bg="#f0f0f0")  # Cor de fundo

# Título em negrito e grande com fundo azul
titulo_frame = tk.Frame(app, bg="#007BFF")
titulo_frame.pack(fill='x', pady=10)
titulo = tk.Label(titulo_frame, text="NFE CÓPIA DE NOTAS", font=("Arial", 24, "bold"), bg="#007BFF", fg="white")
titulo.pack(pady=10)

# Frame para centralizar elementos
frame = tk.Frame(app, bg="#f0f0f0")
frame.pack(expand=True, fill='both')

dados = carregar_dados()
label_origem = tk.Label(frame, text=f"Origem: {dados.get('origem_autorizadas')}\n{dados.get('origem_nao_autorizadas')}", font=("Arial", 10), wraplength=550, anchor="center", bg="#f0f0f0")
label_origem.pack(pady=10)

label_destino = tk.Label(frame, text=f"Destino: {dados.get('destino')}", font=("Arial", 10), wraplength=550, anchor="center", bg="#f0f0f0")
label_destino.pack(pady=5)


# Frame adicional para deslocar itens para baixo
frame_conteudo = tk.Frame(frame, bg="#f0f0f0")
frame_conteudo.pack(pady=(20, 10))  # Ajuste o pady para mover o conteúdo para baixo

label_nota = tk.Label(frame_conteudo, text="Digite os números das notas fiscais (separados por vírgula):", font=("Arial", 12), bg="#f0f0f0")
label_nota.pack(pady=10)

# Configurar validação da entrada
validar_cmd = app.register(validar_entrada)

entrada_nota = tk.Entry(frame_conteudo, font=("Arial", 12), width=50, validate="key", validatecommand=(validar_cmd, "%S"))
entrada_nota.pack(pady=5)

botao_iniciar = tk.Button(frame_conteudo, text="Iniciar Cópia", command=iniciar_copia, font=("Arial", 12), bg="#4CAF50", fg="white")
botao_iniciar.pack(pady=10)

label_email = tk.Label(frame_conteudo, text="Digite o e-mail do destinatário:", font=("Arial", 12), bg="#f0f0f0")
label_email.pack(pady=(10, 0))  # Ajuste o padding para controlar a altura

# Configurar validação do e-mail
validar_email_cmd = app.register(validar_email)

entrada_email = tk.Entry(frame_conteudo, font=("Arial", 12), width=28, validate="key", validatecommand=(validar_email_cmd, "%S"))
entrada_email.pack(pady=5)

# Mensagem de status
status_var = tk.StringVar()
status_var.set("")  # Status inicial vazio
status_label = tk.Label(frame_conteudo, textvariable=status_var, font=("Arial", 12), bg="#f0f0f0", fg="red")
status_label.pack(pady=(10, 5))  # Coloca a mensagem abaixo do botão

# Mensagem de desenvolvedor (ajustada para sobrepor todos os itens)
info_var = tk.StringVar()
info_var.set("Desenvolvido por Jairo Rossi")
info_label = tk.Label(app, textvariable=info_var, font=("Arial", 8), bg="#f0f0f0", fg="gray")
info_label.place(relx=0.5, rely=0.95, anchor="center")  # Ajuste a posição conforme necessário

app.mainloop()
