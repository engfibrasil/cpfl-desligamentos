import csv
import unicodedata
from datetime import datetime as dt
import smtplib
from email.message import EmailMessage
import mimetypes

def remover_acentos(txt):
    if isinstance(txt, str):
        return unicodedata.normalize('NFKD', txt).encode('ASCII', 'ignore').decode('utf-8')
    return txt

def calcular_diferenca_horas(hora_inicio, hora_fim):
    try:
        inicio = dt.strptime(hora_inicio, "%H:%M")
        fim = dt.strptime(hora_fim, "%H:%M")
        diferenca = (fim - inicio).total_seconds() / 3600
        return diferenca
    except Exception:
        return 0

def enviar_email_com_anexo(nome_arquivo, data_inicio, data_fim):
    remetente = "eng.fibrasil@zohomail.com"
    senha = "F1br@$1L"
    destinatario = "robson.santos@fibrasil.com.br"
    copia = "marcos.oliveira@fibrasil.com.br"

    assunto = f"📄 Desligamentos Programados CPFL – Semana {data_inicio} a {data_fim}"
    corpo = f"""
    Senhores,

    Segue em anexo o arquivo com os desligamentos programados da CPFL
    referentes à semana de {data_inicio} a {data_fim}.

    Qualquer dúvida estou à disposição.

    Atenciosamente,
    """

    msg = EmailMessage()
    msg["Subject"] = assunto
    msg["From"] = remetente
    msg["To"] = destinatario
    msg["Cc"] = copia
    msg.set_content(corpo)

    with open(nome_arquivo, "rb") as f:
        conteudo = f.read()
        tipo, _ = mimetypes.guess_type(nome_arquivo)
        tipo_principal, subtipo = tipo.split("/") if tipo else ("application", "octet-stream")
        msg.add_attachment(conteudo, maintype=tipo_principal, subtype=subtipo, filename=nome_arquivo)

    try:
        with smtplib.SMTP("smtp.zoho.com", 587) as smtp:
            smtp.starttls()
            smtp.login(remetente, senha)
            smtp.send_message(msg)
            print("📧 E-mail enviado com sucesso!")
    except Exception as e:
        print(f"❌ Erro ao enviar e-mail: {e}")

# Leitura e filtro
entrada = "desligamentos_cpfl_30-06-2025_a_06-07-2025.csv"
saida = "desligamentos_cpfl_filtrado_sem_acentos.csv"
resultados_filtrados = []

with open(entrada, newline="", encoding="utf-8") as f:
    leitor = csv.DictReader(f)
    for linha in leitor:
        linha_sem_acentos = {k: remover_acentos(v) for k, v in linha.items()}
        manutencao = linha_sem_acentos.get("Manutencao", "")

        if manutencao == "Obras":
            resultados_filtrados.append(linha_sem_acentos)
        elif manutencao == "Manutencao":
            inicio = linha_sem_acentos.get("Inicio", "00:00")
            fim = linha_sem_acentos.get("Fim", "00:00")
            if calcular_diferenca_horas(inicio, fim) > 1:
                resultados_filtrados.append(linha_sem_acentos)

# Exportar CSV
if resultados_filtrados:
    with open(saida, "w", newline="", encoding="utf-8") as f:
        escritor = csv.DictWriter(f, fieldnames=resultados_filtrados[0].keys())
        escritor.writeheader()
        escritor.writerows(resultados_filtrados)

    print(f"✅ Arquivo filtrado salvo como: {saida} com {len(resultados_filtrados)} registros")
    
    # Enviar por e-mail
    hoje = dt.today().date()
    inicio_semana = hoje - datetime.timedelta(days=hoje.weekday())
    fim_semana = inicio_semana + datetime.timedelta(days=6)
    enviar_email_com_anexo(saida, inicio_semana.strftime("%d/%m/%Y"), fim_semana.strftime("%d/%m/%Y"))
else:
    print("⚠️ Nenhum registro encontrado após os filtros.")
