import csv
import requests
import oracledb
from datetime import datetime

# Configurações de conexão
username = ''
password = ''
host = ''
porta = ''
sid = ''

# Crie uma conexão
conn = oracledb.connect(user=username, password=password, host=host, port=porta, service_name=sid)
cursor = conn.cursor()

# Pesquisa Banco
cursor.execute("SELECT NM_PRESTADOR, DS_CODIGO_CONSELHO, DECODE(TP_SITUACAO, 'A', 'ATIVO', 'I', 'INATIVO') FROM PRESTADOR WHERE CD_TIP_PRESTA = 8")


data_e_hora_atuais = datetime.now()
dh_atual = data_e_hora_atuais.strftime("%d/%m/%Y %H:%M")

# Retorno das pesquisas
linhas = cursor.fetchall()
for linha in linhas:
    nome = linha[0]
    crm = linha[1]
    ativo_sistema = linha[2]

    # URL da API com o crm a ser pesquisado
    url = f'https://api.cremesp.org.br/guia-medico/medico-info/{crm}'

    # Realize uma solicitação GET à API
    response = requests.get(url)

    # Verifique se a solicitação foi bem-sucedida (código de status HTTP 200)
    if response.status_code == 200:
        # Os dados da resposta da API estão em formato JSON
        dados = response.json()

        # Recebimento dos dados necessários
        nome_cremesp = dados.get("nome")
        situacao = dados.get("situacao")
        data_inativo = dados.get("mensagemStatus")
    else:
        nome_cremesp = 'Não Encontrado'
        situacao = 'Não Encontrado'
        data_inativo = 'Não Encontrado'

    if nome.lower() != nome_cremesp.lower():
        nm_igual = False
    else:
        nm_igual = True
    
    if situacao == 'A':
        situacao = 'ATIVO'
    else:
        situacao = 'INATIVO'
      
    # Exportando dados alimentando em um arquivo CSV
    with open('crm.csv', 'a', newline='', encoding='utf-8') as arquivo_csv:
        escritor_csv = csv.writer(arquivo_csv)
        escritor_csv.writerow([nome, nome_cremesp, nm_igual, crm, situacao, ativo_sistema, data_inativo, dh_atual])
