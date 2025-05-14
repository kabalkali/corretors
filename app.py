import pandas as pd
import telebot
import os
import requests
from datetime import datetime
from telebot.types import InlineKeyboardButton, InlineKeyboardMarkup

# Chave da API do Telegram
CHAVE_API = "7795566868:AAG5jU1tDM4DNop6m8oymVs3c8XoK4_v6bk"
bot = telebot.TeleBot(CHAVE_API)

# Consulta o logradouro na API ViaCEP
def consultar_cep(cep):
    url = f"https://viacep.com.br/ws/{cep}/json/"
    try:
        response = requests.get(url, timeout=5)
        if response.status_code == 200:
            data = response.json()
            if "erro" not in data:
                return data.get("logradouro", "")
        return None
    except Exception as e:
        print(f"Erro ao consultar o CEP {cep}: {e}")
        return None

# Atualiza o arquivo ceps.csv com novos dados
def atualizar_banco_ceps(cep, logradouro):
    try:
        with open('ceps.csv', 'a', encoding='utf-8') as f:
            f.write(f"{cep},{logradouro}\n")
        print(f"CEP {cep} adicionado ao ceps.csv")
    except Exception as e:
        print(f"Erro ao atualizar ceps.csv: {e}")

# Função para carregar a planilha Excel a partir do arquivo recebido pelo Telegram
def carregar_planilha(caminho_arquivo):
    try:
        df = pd.read_excel(caminho_arquivo)
        print("Planilha carregada com sucesso!")
        return df
    except Exception as e:
        print(f"Erro ao carregar a planilha: {e}")
        return None

# Função para salvar a planilha modificada com nome personalizado
def salvar_planilha(df):
    data_atual = datetime.now().strftime("%d-%m-%Y")
    caminho_salvar = f"Corrigido {data_atual}.xlsx"
    try:
        df.to_excel(caminho_salvar, index=False)
        print(f"Planilha salva com sucesso em: {caminho_salvar}")
        return caminho_salvar
    except Exception as e:
        print(f"Erro ao salvar a planilha: {e}")
        return None

# Função para normalizar endereços
def normalizar_endereco(endereco):
    if pd.isnull(endereco):
        return endereco
    # Remover espaços ao redor e normalizar números
    partes = endereco.split(',')
    if len(partes) == 2:
        rua = partes[0].strip()
        numero = partes[1].strip().lstrip('0')  # Remove zeros à esquerda
        return f"{rua}, {numero}"
    return endereco.strip()

# Função para processar a planilha
def processar_planilha(caminho_arquivo):
    try:
        banco_dados = pd.read_csv('ceps.csv', header=None, names=['CEP', 'Logradouro'])
    except Exception as e:
        print(f"Erro ao carregar o arquivo CSV: {e}")
        return None

    planilha = carregar_planilha(caminho_arquivo)

    if planilha is not None:
        if planilha.shape[1] > 4:
            planilha['Address Line 2'] = planilha.iloc[:, 4].str.split(',').str[2].fillna('')
            planilha['number'] = planilha.iloc[:, 4]
            planilha.iloc[:, 4] = planilha.iloc[:, 4].str.split(',').str[0]
            planilha['number'] = planilha['number'].str.split(',').str[1]

            planilha = planilha.merge(banco_dados[['CEP', 'Logradouro']], left_on='Zipcode/Postal code', right_on='CEP', how='left')
            planilha.iloc[:, 4] = planilha['Logradouro']

            # Consultar CEPs ausentes
            faltantes = planilha[planilha['Logradouro'].isna()]['Zipcode/Postal code'].unique()
            for cep in faltantes:
                cep_str = str(cep).zfill(8)
                logradouro = consultar_cep(cep_str)
                if logradouro:
                    planilha.loc[planilha['Zipcode/Postal code'] == cep, planilha.columns[4]] = logradouro
                    atualizar_banco_ceps(cep_str, logradouro)
                else:
                    print(f"CEP {cep} não encontrado na API.")

            planilha.drop(columns=['CEP', 'Logradouro'], inplace=True)

            # Normaliza os endereços antes de agrupar
            planilha['Destination Address'] = planilha['Destination Address'].apply(normalizar_endereco)
            planilha['number'] = planilha['number'].apply(lambda x: str(x).strip().lstrip('0'))  # Normaliza o número

            planilha['Pacotes na Parada'] = planilha.groupby(['Destination Address', 'number'])['Sequence']\
                .transform(lambda x: ', '.join(map(str, x.unique())))
            planilha['Pacotes Contados'] = planilha.groupby(['Destination Address', 'number'])['Sequence']\
                .transform(lambda x: x.nunique())

            planilha['Pacotes na Parada'] = planilha.apply(
                lambda row: f"{row['Pacotes na Parada']} - Total: {row['Pacotes Contados']} pacotes"
                if row['Pacotes Contados'] > 4 else row['Pacotes na Parada'],
                axis=1
            )

            planilha.drop(columns=['Pacotes Contados'], inplace=True)

            planilha['Destination Address'] = planilha.apply(
                lambda row: f"{row['Destination Address']}, {row['number']}".strip() if pd.notnull(row['number']) else row['Destination Address'],
                axis=1
            )

            planilha.drop(columns=['number'], inplace=True)
            planilha.drop(columns=['Sequence', 'Stop', 'SPX TN'], inplace=True)
            planilha.drop_duplicates(subset=['Destination Address', 'Pacotes na Parada'], inplace=True)

            print("Colunas manipuladas, valores adicionados e duplicatas removidas com sucesso!")
            caminho_salvar = salvar_planilha(planilha)
            return caminho_salvar
        else:
            print("A planilha não possui colunas suficientes.")
            return None
    else:
        print("Erro ao carregar a planilha.")
        return None

# Funções do Telegram Bot - Modificadas para ter apenas "Corrigir"
def exibir_opcoes_iniciais(mensagem):
    keyboard = InlineKeyboardMarkup([
        [InlineKeyboardButton("Corrigir Rota", callback_data='corrigir')]
    ])
    bot.send_message(mensagem.chat.id, "🔧 Envie sua planilha para correção automática:", reply_markup=keyboard)

@bot.message_handler(commands=["start", "corrigir"])
def start(mensagem):
    exibir_opcoes_iniciais(mensagem)

@bot.message_handler(commands=["Corrigir"])
def opcao_corrigir(mensagem):
    bot.send_message(mensagem.chat.id, "📤 Envie o arquivo Excel da rota Shopee para correção.")

@bot.message_handler(content_types=['document'])
def handle_document(message):
    try:
        file_id = message.document.file_id
        file_info = bot.get_file(file_id)
        downloaded_file = bot.download_file(file_info.file_path)

        with open("received_file.xlsx", 'wb') as new_file:
            new_file.write(downloaded_file)

        caminho_modificado = processar_planilha("received_file.xlsx")

        if caminho_modificado:
            with open(caminho_modificado, 'rb') as doc:
                bot.send_message(message.chat.id, "✅ Planilha corrigida com sucesso!")
                bot.send_document(message.chat.id, doc)
            os.remove(caminho_modificado)
        else:
            bot.send_message(message.chat.id, "❌ Erro ao processar a planilha. Verifique o formato.")

        os.remove("received_file.xlsx")

    except Exception as e:
        bot.send_message(message.chat.id, f"⚠️ Erro: {e}")
    exibir_opcoes_iniciais(message)

@bot.callback_query_handler(func=lambda call: True)
def callback_query(call):
    if call.data == 'corrigir':
        opcao_corrigir(call.message)
    exibir_opcoes_iniciais(call.message)

# Iniciar o bot
bot.polling()