import requests
import datetime
import csv
import time

def obter_periodo_semana():
    hoje = datetime.date.today()
    inicio = hoje - datetime.timedelta(days=hoje.weekday())
    fim = inicio + datetime.timedelta(days=6)
    return inicio, fim

municipios = {
    "Lajeado": 761,
    "Canela": 677,
    "Igrejinha": 745,
    "Alegrete": 645,
    "Estrela": 700,
    "Lagoa Vermelha": 760,
    "Palmeira das Missões": 783,
    "Santana do Livramento": 804,
    "Sarandi": 806,
    "Uruguaiana": 827,
    "Vacaria": 828,
    "Venancio Aires": 832,
    "Cachoeira do Sul": 670,
    "Três de Maio": 821,
    "Santo Angelo": 811,
    "Santa Rosa": 803,
    "Nova Petropolis": 778,
    "Taquara": 819,
    "Ivoti": 748,
    "Carlos Barbosa": 676,
    "Dois Irmãos": 690,
    "Erechim": 698,
    "Flores da Cunha": 707,
    "São Marcos": 817,
    "Veranópolis": 835,
}

headers = {
    "User-Agent": "Mozilla/5.0"
}

data_inicio_dt, data_fim_dt = obter_periodo_semana()
data_inicio = data_inicio_dt.strftime("%d/%m/%Y")
data_fim = data_fim_dt.strftime("%d/%m/%Y")
nome_arquivo = f"desligamentos_cpfl_{data_inicio_dt.strftime('%d-%m-%Y')}_a_{data_fim_dt.strftime('%d-%m-%Y')}.csv"

resultados = []

print(f"Iniciando consulta de {len(municipios)} municípios entre {data_inicio} e {data_fim}...\n")

for nome, id_municipio in municipios.items():
    print(f"Consultando {nome} (ID {id_municipio})...")

    url = (
        f"https://spir.cpfl.com.br/api/ConsultaDesligamentoProgramado/Pesquisar"
        f"?PeriodoDesligamentoInicial={data_inicio}"
        f"&PeriodoDesligamentoFinal={data_fim}"
        f"&IdMunicipio={id_municipio}"
        f"&NomeBairro=&NomeRua="
    )

    try:
        response = requests.get(url, headers=headers, timeout=10)
        data = response.json()

        if not data.get("Data"):
            print(f"  👉 Nenhum desligamento encontrado.")
            continue

        for cidade in data["Data"]:
            nome_municipio = cidade["NomeMunicipio"]
            for dia in cidade["Datas"]:
                data_municipio = dia["Data"][:10]
                for documento in dia["Documentos"]:
                    descricao = documento.get("DescricaoDocumento")
                    estado = documento.get("Estado")
                    necessidade = documento.get("NecessidadeDocumento")
                    inicio_execucao = documento.get("PeriodoExecucaoInicial")[11:16]
                    fim_execucao = documento.get("PeriodoExecucaoPeriodoFinal")[11:16]

                    for bairro in documento.get("Bairros", []):
                        nome_bairro = bairro["NomeBairro"]
                        for rua in bairro.get("Ruas", []):
                            nome_rua = rua["NomeRua"]

                            resultados.append({
                                "Cidade": nome_municipio,
                                "Data": data_municipio,
                                "Início": inicio_execucao,
                                "Fim": fim_execucao,
                                "Execução": estado,
                                "Manutenção": necessidade,
                                "Descrição": descricao,
                                "Bairro": nome_bairro,
                                "Rua": nome_rua
                            })

        time.sleep(0.5)

    except Exception as e:
        print(f"  ⚠️ Erro ao consultar {nome}: {e}")

# Exportar CSV com nome personalizado e encoding UTF-8 (corrigindo caracteres)
if resultados:
    with open(nome_arquivo, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=resultados[0].keys())
        writer.writeheader()
        writer.writerows(resultados)

    print(f"\n✅ Consulta finalizada! Total de registros salvos: {len(resultados)}")
    print(f"📄 Arquivo gerado: {nome_arquivo}")
else:
    print("\n⚠️ Nenhum desligamento encontrado para os municípios informados.")
