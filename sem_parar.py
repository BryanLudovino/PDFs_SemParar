import pandas as pd

def _limpar_id_rota(id_rota):
    """
    Limpa o ID da rota de forma robusta.
    Converte para string e remove ".0" se for um número inteiro.
    Trata valores vazios ou nulos como "Fora de Rota".
    """
    if pd.isna(id_rota) or str(id_rota).strip() == '':
        return "Fora de Rota"
    try:
        val = float(id_rota)
        if val.is_integer():
            return str(int(val))
    except (ValueError, TypeError):
        # Mantém como string se não for um número que possamos limpar.
        pass
    return str(id_rota)

#Importando os dados de Pedágios
#Nesse primeiro comando, eu estou excluindo as 2 primeiras linhas do meu dataframe, e utilizando a 3 como cabeçario

def carregar_pedagio(sem_parar):
    """Carrega e prepara os dados de pedágio."""
    lista = pd.read_excel(sem_parar, header=2)
    tabela = pd.DataFrame(lista)
    tabela['Datax'] = tabela['Data'] + ' ' + tabela['Horário']
    tabela['Tipo'] = "Pedágio"
    tabela['Tipo de uso'] == "PASSAGENS"
    tabela_pedagio = tabela.drop(
        ['Horário','Número da Fatura','Débito/Crédito','Viagem de Vale Pedágio',
         'Embarcador','Sentido da Praça','Tipo de uso','Tipo do veículo'], axis=1
    )
    return tabela_pedagio

def carregar_rotas(Vickos):
    """Carrega as rotas do modelo Vickos."""
    rotas = pd.read_excel(Vickos, usecols=['ID EXTERNO','DATA DA CARGA','PLACA'])
    return pd.DataFrame(rotas)

def conciliar_dados(tabela_pedagio, rotas):
    """Faz a conciliação dos dados entre as tabelas."""
    modelo = pd.merge(
        tabela_pedagio, rotas,
        left_on=['Data','Placa Veículo'],
        right_on=['DATA DA CARGA','PLACA'],
        how='left'
    )
    modelo = modelo.rename(columns={'ID EXTERNO': 'Id da Rota'})
    modelo['Id da Rota'] = modelo['Id da Rota'].apply(_limpar_id_rota)
    return modelo

def tratar_valor(modelo):
    """Trata a coluna de valor para float, considerando o formato monetário brasileiro."""
    # Garante que a coluna é do tipo string para manipulação, tratando valores nulos.
    valores = modelo['Valor(R$)'].astype(str)
    
    valores_tratados = (
        valores
        .str.replace('R\\$', '', regex=True)  # Remove o símbolo de real
        .str.replace('.', '', regex=False)   # Remove o separador de milhares
        .str.strip()                         # Remove espaços em branco no início e fim
    )
    
    # Converte para tipo numérico. Se houver erro na conversão, o valor se tornará Nulo (NaN)
    modelo['Valor(R$)'] = pd.to_numeric(valores_tratados, errors='coerce')
    return modelo

def fluxo_principal(path_pedagio, path_rotas, path_saida, progress_callback=None):
    """Executa todo o fluxo de processamento."""
    if progress_callback: progress_callback(10, "Carregando pedágios...")
    tabela_pedagio = carregar_pedagio(path_pedagio)

    if progress_callback: progress_callback(25, "Carregando rotas...")
    rotas = carregar_rotas(path_rotas)

    if progress_callback: progress_callback(50, "Conciliando dados...")
    modelo = conciliar_dados(tabela_pedagio, rotas)

    if progress_callback: progress_callback(75, "Tratando valores...")
    modelo = tratar_valor(modelo)

    if progress_callback: progress_callback(90, "Salvando relatório...")
    modelo.to_excel(path_saida, index=False)
    somente_rotas = modelo['Id da Rota']
    lista_rotas = somente_rotas.drop_duplicates()

    if progress_callback: progress_callback(100, "Concluído!")
    return modelo, lista_rotas

# Exemplo de uso:
if __name__ == "__main__":
    modelo, lista_rotas = fluxo_principal(
        "Maio 2025.xlsx",
        "Maio 2025 - vickos.xlsx",
        "Análise Relatório Sem Parar.xlsx"
    )
