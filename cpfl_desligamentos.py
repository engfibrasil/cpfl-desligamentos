import requests
import datetime
import csv
import time
import smtplib
from email.message import EmailMessage
import mimetypes
import unicodedata
import io
from datetime import datetime as dt

# === CONFIGURAÇÕES ===
REMETENTE = "eng.fibrasil@zohomail.com"
SENHA = "F1br@$1L"
DESTINATARIO = "nathalia.martines@fibrasil.com.br"
COPIA = "marcos.oliveira@fibrasil.com.br, robson.santos@fibrasil.com.br, fabiano.cezar@fibrasil.com.br"
CHAVE_GEOCODIFICACAO = "5b3ce3597851110001cf6248f0b838209a014d1489661eb1afacd92a"

municipios = {
    "Lajeado":761, "Canela":698, "Igrejinha":770, "Alegrete":894,
    "Estrela":732, "Lagoa Vermelha":836, "Palmeira das Missões":1030,
    "Santana do Livramento":889, "Sarandi":1027, "Uruguaiana":893,
    "Vacaria":717, "Venancio Aires":795, "Cachoeira do Sul":812,
    "Três de Maio":988, "Santo Angelo":1070, "Santa Rosa":984,
    "Nova Petropolis":706, "Taquara":767, "Ivoti":781,
    "Carlos Barbosa":729, "Dois Irmãos":778, "Erechim":869,
    "Flores da Cunha":708, "São Marcos":711, "Veranópolis":834
}

headers = {
    "User-Agent": "Mozilla/5.0"
}

# === FUNÇÕES ===

def remover_acentos(txt):
    if isinstance(txt, str):
        return unicodedata.normalize('NFKD', txt).encode('ASCII', 'ignore').decode('utf-8')
    return txt

def calcular_diferenca_horas(hora_inicio, hora_fim):
    try:
        inicio = dt.strptime(hora_inicio, "%H:%M")
        fim = dt.strptime(hora_fim, "%H:%M")
        return (fim - inicio).total_seconds() / 3600
    except Exception:
        return 0

def obter_periodo_semana():
    hoje = datetime.date.today()
    inicio = hoje - datetime.timedelta(days=hoje.weekday())
    fim = inicio + datetime.timedelta(days=6)
    return inicio, fim

def consultar_lat_lon(endereco, chave):
    try:
        url = "https://api.openrouteservice.org/geocode/search"
        params = {
            "api_key": chave,
            "text": endereco,
            "boundary.country": "BR",
            "size": 1
        }
        resposta = requests.get(url, params=params, timeout=10)
        dados = resposta.json()
        features = dados.get("features", [])
        if features:
            coordenadas = features[0]["geometry"]["coordinates"]
            return coordenadas[1], coordenadas[0]  # (lat, lon)
    except Exception as e:
        print(f"  ⚠️ Erro na geocodificação '{endereco}': {e}")
    return "", ""

def enviar_email_com_anexo_csv_em_memoria(conteudo_csv, nome_arquivo, data_inicio, data_fim):
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
    msg["From"] = REMETENTE
    msg["To"] = DESTINATARIO
    msg["Cc"] = COPIA
    msg.set_content(corpo)

    msg.add_attachment(conteudo_csv.encode("utf-8"), maintype="text", subtype="csv", filename=nome_arquivo)

    try:
        with smtplib.SMTP("smtp.zoho.com", 587) as smtp:
            smtp.starttls()
            smtp.login(REMETENTE, SENHA)
            smtp.send_message(msg)
            print("📧 E-mail enviado com sucesso!")
    except Exception as e:
        print(f"❌ Erro ao enviar e-mail: {e}")

# === EXECUÇÃO PRINCIPAL ===

data_inicio_dt, data_fim_dt = obter_periodo_semana()
data_inicio = data_inicio_dt.strftime("%d/%m/%Y")
data_fim = data_fim_dt.strftime("%d/%m/%Y")
nome_arquivo = f"desligamentos_cpfl_{data_inicio_dt.strftime('%d-%m-%Y')}_a_{data_fim_dt.strftime('%d-%m-%Y')}.csv"

print(f"Iniciando consulta de {len(municipios)} municípios entre {data_inicio} e {data_fim}...\n")
resultados = []

for nome, id_municipio in municipios.items():
    print(f"Consultando {nome}...")
    url = (
        f"https://spir.cpfl.com.br/api/ConsultaDesligamentoProgramado/Pesquisar"
        f"?PeriodoDesligamentoInicial={data_inicio}"
        f"&PeriodoDesligamentoFinal={data_fim}"
        f"&IdMunicipio={id_municipio}"
    )

    try:
        response = requests.get(url, headers=headers, timeout=10)
        data = response.json()

        if not data.get("Data"):
            print(f"  👉 Nenhum desligamento.")
            continue

        for cidade in data["Data"]:
            nome_municipio = cidade["NomeMunicipio"]
            for dia in cidade["Datas"]:
                data_municipio = dia["Data"][:10]
                for doc in dia["Documentos"]:
                    manutencao = remover_acentos(doc.get("NecessidadeDocumento", ""))
                    inicio = doc.get("PeriodoExecucaoInicial", "")[11:16]
                    fim = doc.get("PeriodoExecucaoPeriodoFinal", "")[11:16]
                    descricao = remover_acentos(doc.get("DescricaoDocumento", ""))
                    estado = remover_acentos(doc.get("Estado", ""))

                    incluir = False
                    if manutencao == "Obras":
                        incluir = True
                    elif manutencao == "Manutencao" and calcular_diferenca_horas(inicio, fim) > 1:
                        incluir = True

                    if incluir:
                        for bairro in doc.get("Bairros", []):
                            nome_bairro = remover_acentos(bairro["NomeBairro"])
                            for rua in bairro.get("Ruas", []):
                                nome_rua = remover_acentos(rua["NomeRua"])
                                endereco_completo = f"{nome_rua}, {nome_bairro}, {nome_municipio}"
                                lat, lon = consultar_lat_lon(endereco_completo, CHAVE_GEOCODIFICACAO)
                                time.sleep(1)  # Evitar excesso na API

                                resultados.append({
                                    "Cidade": remover_acentos(nome_municipio),
                                    "Data": data_municipio,
                                    "Inicio": inicio,
                                    "Fim": fim,
                                    "Execucao": estado,
                                    "Manutencao": manutencao,
                                    "Descricao": descricao,
                                    "Bairro": nome_bairro,
                                    "Rua": nome_rua,
                                    "Latitude": lat,
                                    "Longitude": lon
                                })
        time.sleep(0.5)
    except Exception as e:
        print(f"  ⚠️ Erro: {e}")

# Gerar CSV em memória
if resultados:
    csv_buffer = io.StringIO()
    writer = csv.DictWriter(csv_buffer, fieldnames=resultados[0].keys())
    writer.writeheader()
    writer.writerows(resultados)
    csv_conteudo = csv_buffer.getvalue()

    enviar_email_com_anexo_csv_em_memoria(csv_conteudo, nome_arquivo, data_inicio, data_fim)
else:
    print("⚠️ Nenhum registro encontrado.")
