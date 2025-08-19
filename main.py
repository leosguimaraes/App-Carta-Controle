import dash
import dash_bootstrap_components as dbc
from dash import dcc, html
from dash.dependencies import Input, Output, State
import plotly.graph_objs as go
import numpy as np
import pandas as pd
from datetime import datetime, timedelta
import os
from flask import send_file, request


######################################## DADOS PARA CADA EQUIPAMENTO ######################################################

dados_calibracao = pd.read_excel('Planilhas/Metionina.xlsx')
multitek_cb = pd.read_excel('Planilhas/multitek_cb.xlsx')
multitek_cm = pd.read_excel('Planilhas/multitek_cm.xlsx')
multitek_ca = pd.read_excel('Planilhas/multitek_ca.xlsx')
pananalytical = pd.read_excel('Planilhas/Pananalytical.xlsx')
oxford_frx = pd.read_excel('Planilhas/Arsenio.xlsx')


dados_teste = {
    "Carbono": [],
    "Nitrogênio": [],
    "Hidrogênio": [],
    "Enxofre": []
}

dados_flash = pd.DataFrame(columns=['Elemento', 'Mês', 'Data de Inserção', 'Valor'])
dados_multitek_cb = pd.DataFrame(columns=['Elemento', 'Mês', 'Data de Inserção', 'Valor'])
dados_multitek_cm = pd.DataFrame(columns=['Elemento', 'Mês', 'Data de Inserção', 'Valor'])
dados_multitek_ca = pd.DataFrame(columns=['Elemento', 'Mês', 'Data de Inserção', 'Valor'])
dados_pananalytical = pd.DataFrame(columns=['Elemento', 'Mês', 'Data de Inserção', 'Valor'])
dados_oxford_frx = pd.DataFrame(columns=['Elemento', 'Mês', 'Data de Inserção', 'Valor'])


# Estilos e tema Bootstrap
app = dash.Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP], suppress_callback_exceptions=True)


# Caminho do arquivo CSV
CSV_PATH = 'Planilhas/dados_flash_meses.xlsx'
caminho_multitek_cb = 'Planilhas/multitek_cb_meses.xlsx'
caminho_multitek_cm = 'Planilhas/multitek_cm_meses.xlsx'
caminho_multitek_ca = 'Planilhas/multitek_ca_meses.xlsx'
caminho_pananalytical = 'Planilhas/pananalytical_meses.xlsx'
caminho_oxford_frx = 'Planilhas/oxford_frx_meses.xlsx'

####################################################### CARREGAR PLANILHAS #########################################################

def carregar_dados_flash():
    try:
        df = pd.read_excel(CSV_PATH)
    except FileNotFoundError:
        # Se o arquivo não existir, cria um DataFrame vazio com as colunas necessárias
        df = pd.DataFrame(columns=['Elemento', 'Mês', 'Data de Inserção', 'Valor'])
    return df

def carregar_dados_multitek_cb():
    try:
        df = pd.read_excel(caminho_multitek_cb)
    except FileNotFoundError:
        # Se o arquivo não existir, cria um DataFrame vazio com as colunas necessárias
        df = pd.DataFrame(columns=['Elemento', 'Mês', 'Data de Inserção', 'Valor'])
    return df

def carregar_dados_multitek_cm():
    try:
        df = pd.read_excel(caminho_multitek_cm)
    except FileNotFoundError:
        # Se o arquivo não existir, cria um DataFrame vazio com as colunas necessárias
        df = pd.DataFrame(columns=['Elemento', 'Mês', 'Data de Inserção', 'Valor'])
    return df

def carregar_dados_multitek_ca():
    try:
        df = pd.read_excel(caminho_multitek_ca)
    except FileNotFoundError:
        # Se o arquivo não existir, cria um DataFrame vazio com as colunas necessárias
        df = pd.DataFrame(columns=['Elemento', 'Mês', 'Data de Inserção', 'Valor'])
    return df

def carregar_dados_pananalytical():
    try:
        df = pd.read_excel(caminho_pananalytical)
    except FileNotFoundError:
        # Se o arquivo não existir, cria um DataFrame vazio com as colunas necessárias
        df = pd.DataFrame(columns=['Elemento', 'Mês', 'Data de Inserção', 'Valor'])
    return df

def carregar_dados_oxford_frx():
    try:
        df = pd.read_excel(caminho_oxford_frx)
    except FileNotFoundError:
        # Se o arquivo não existir, cria um DataFrame vazio com as colunas necessárias
        df = pd.DataFrame(columns=['Elemento', 'Mês', 'Data de Inserção', 'Valor'])
    return df
##################################### SIDEBAR #############################################################################

sidebar = html.Div(
    [
    
        html.H4("Equipamentos", style={"color": "#f8f9fa", "text-align": "center"}),
        html.Hr(),
        dbc.Nav(
            [
                dbc.Button("Página Inicial", id="home-button", href="/", color="primary", className="mb-3", n_clicks=0),
                
                dbc.NavItem(
                    dbc.Button("Flash 2000", id="flash2000-button", color="primary", className="mb-3", n_clicks=0,
                           style={"width": "100%", "text-align": "center", "padding": "5px"}),      
                ),
                dbc.Collapse(
                    dbc.Nav(
                        [
                            dbc.NavLink("Carbono", href="/carbono", id="carbono-link", active="exact"),
                            dbc.NavLink("Nitrogênio", href="/nitrogenio", id="nitrogenio-link", active="exact"),
                            dbc.NavLink("Hidrogênio", href="/hidrogenio", id="hidrogenio-link", active="exact"),
                            dbc.NavLink("Enxofre", href="/enxofre", id="enxofre-link", active="exact"),
                        ],
                        vertical=True,
                        pills=True,
                    ),
                    id="submenu-flash2000",
                    is_open=False,
                ),
            ],
            vertical=True,            
        ),

        dbc.Nav(
            [   
                dbc.NavItem(
                    dbc.Button("Flash 2000 2", id="flash2000_2-button", color="primary", className="mb-3", n_clicks=0,
                           style={"width": "100%", "text-align": "center", "padding": "5px"}),      
                ),
                dbc.Collapse(
                    dbc.Nav(
                        [
                            dbc.NavLink("Carbono", href="/carbono_flash2", id="carbono_flash2-link", active="exact"),
                            dbc.NavLink("Hidrogênio", href="/hidrogenio_flash2", id="hidrogenio_flash2-link", active="exact"),
                            dbc.NavLink("Oxigênio", href="/oxigenio_flash2", id="oxigenio_flash2-link", active="exact"),
                        ],
                        vertical=True,
                        pills=True,
                    ),
                    id="submenu-flash2000_2",
                    is_open=False,
                ),
            ],
            vertical=True,            
        ),


        dbc.Nav(
            [
                
                dbc.NavItem(
                    dbc.Button("Multitek - CB", id="multitek_cb-button", color="primary", className="mb-3", n_clicks=0,
                           style={"width": "100%", "text-align": "center", "padding": "5px"}), 
                    
                    
                ),
                dbc.Collapse(
                    dbc.Nav(
                        [
                            dbc.NavLink("Nitrogênio", href="/nitrogenio_multitek_cb", id="nitrogenio_multitekcb-link", active="exact"),
                            dbc.NavLink("Enxofre", href="/enxofre_multitek_cb", id="enxofre_multitek_cb-link", active="exact"),
                            
                        ],
                        vertical=True,
                        pills=True,
                    ),
                    id="submenu-multitek_cb",
                    is_open=False,
                ),
            ],
            vertical=True,

            
        ),

         dbc.Nav(
            [
                
                dbc.NavItem(
                    dbc.Button("Multitek - CM", id="multitek_cm-button", color="primary", className="mb-3", n_clicks=0,
                           style={"width": "100%", "text-align": "center", "padding": "5px"}), 
                    
                    
                ),
                dbc.Collapse(
                    dbc.Nav(
                        [
                            dbc.NavLink("Nitrogênio", href="/nitrogenio_multitek_cm", id="nitrogenio_multitek_cm-link", active="exact"),
                            dbc.NavLink("Enxofre", href="/enxofre_multitek_cm", id="enxofre_multitek_cm-link", active="exact"),
                            
                        ],
                        vertical=True,
                        pills=True,
                    ),
                    id="submenu-multitek_cm",
                    is_open=False,
                ),
            ],
            vertical=True,

        ),

        dbc.Nav(
            [
                
                dbc.NavItem(
                    dbc.Button("Multitek - CA", id="multitek_ca-button", color="primary", className="mb-3", n_clicks=0,
                           style={"width": "100%", "text-align": "center", "padding": "5px"}), 
                    
                    
                ),
                dbc.Collapse(
                    dbc.Nav(
                        [
                            dbc.NavLink("Nitrogênio", href="/nitrogenio_multitek_ca", id="nitrogenio_multitek_ca-link", active="exact"),
                            dbc.NavLink("Enxofre", href="/enxofre_multitek_ca", id="enxofre_multitek_ca-link", active="exact"),
                            
                        ],
                        vertical=True,
                        pills=True,
                    ),
                    id="submenu-multitek_ca",
                    is_open=False,
                ),
            ],
            vertical=True,

        ),

        dbc.Nav(
            [
                dbc.Button("PanAnalytical - FRX", id="pananalytical-button", href="/pananalytical_frx", color="primary", className="mb-3", n_clicks=0,
                           style={"width": "100%", "text-align": "center", "padding": "5px"}),
                
            ]),  

        dbc.Nav(
            [
                dbc.Button("Oxford - FRX", id="oxford-button", href="/oxford_frx", color="primary", className="mb-3", n_clicks=0,
                          style={"width": "100%", "text-align": "center", "padding": "5px"}),
                
            ]),
        
    ],
    style={"position": "fixed", "top": 0, "left": 0, "bottom": 0, "width": "220px", "padding": "20px",
           "background-color":  "#343a40"},

        
        
)



################################################### LAYOUT PÁGINA INICIAL ##########################################################

def gerar_pagina_inicial():
    return html.Div(
        [
            html.H1("Laboratório de Análise Térmica e Elementar", style={"textAlign": "center", "font-size": "61px"}),
            html.Img(
                src="/assets/LABQ.png", 
                style={
                    "display": "block",
                    "margin-left": "auto",
                    "margin-right": "auto",
                    "width": "50%", 
                    "opacity": 0.6 
                }
            ),
        ],
        style={
            "display": "flex", 
            "flexDirection": "column", 
            "justify-content": "center", 
            "align-items": "center", 
            "height": "100vh" 
        }
    )



##################################################### CARTA CONTROLE FLASH 2000 ###################################################

def criar_carta_controle(dados_calibracao, dados_teste, elemento):
    media = np.mean(dados_calibracao)
    desvio_padrao = np.std(dados_calibracao)
    LCS = media + 3 * desvio_padrao
    LCI = media - 3 * desvio_padrao
    LCI_2S = media - 2 * desvio_padrao
    LCS_2S = media + 2 * desvio_padrao


    fig = go.Figure()

    # Adicionar os dados de calibração ao gráfico
    fig.add_trace(go.Scatter(x=list(range(len(dados_calibracao))), y=dados_calibracao,
                             mode='lines+markers', name=f'% {elemento}', marker=dict(color='blue')))
    

    # Adicionar a média e os limites de controle
    fig.add_trace(go.Scatter(x=[1, len(dados_calibracao) + 1], y=[media, media],
                             mode='lines', line=dict(color='green', dash='dash'), name='Média'))
    fig.add_trace(go.Scatter(x=[1, len(dados_calibracao) + 1], y=[LCS, LCS],
                             mode='lines', line=dict(color='red', dash='dash'), name='LCS 3s'))
    fig.add_trace(go.Scatter(x=[1, len(dados_calibracao) + 1], y=[LCI, LCI],
                             mode='lines', line=dict(color='red', dash='dash'), name='LCI 3s'))
    fig.add_trace(go.Scatter(x=[1, len(dados_calibracao) + 1], y=[LCS_2S, LCS_2S],
                             mode='lines', line=dict(color='#FFD700', dash='dash'), name='LCS 2s'))
    fig.add_trace(go.Scatter(x=[1, len(dados_calibracao) + 1], y=[LCI_2S, LCI_2S],
                             mode='lines', line=dict(color='#FFD700', dash='dash'), name='LCI 2s'))

    fig.update_layout(xaxis_title='Execução', yaxis_title=f'Resultado ({elemento})')
    return fig

def criar_carta_controle_meses(dados, elemento, mes='Mês Não Especificado'):

    dados = dados.copy()

    dados.loc[:, 'Data de Inserção'] = pd.to_datetime(dados['Data de Inserção'], errors='coerce')
    # Mapeamento de meses
    meses = {
        "Janeiro": 1, "Fevereiro": 2, "Março": 3, "Abril": 4, "Maio": 5, "Junho": 6,
        "Julho": 7, "Agosto": 8, "Setembro": 9, "Outubro": 10, "Novembro": 11, "Dezembro": 12
    }

    # Calcular estatísticas
    media = np.mean(dados_calibracao['Carbono'])
    desvio_padrao = np.std(dados_calibracao['Carbono'])
    LCS = media + 3 * desvio_padrao
    LCI = media - 3 * desvio_padrao
    LCI_2S = media - 2 * desvio_padrao
    LCS_2S = media + 2 * desvio_padrao

    # Criar figura
    fig = go.Figure()

    if mes in meses:
        mes_num = meses[mes]
        ano = datetime.now().year  # Ajustar para o ano atual ou permitir seleção de ano

        # Obter o intervalo de datas para o mês
        min_data = datetime(ano, mes_num, 1)
        max_data = (datetime(ano, mes_num + 1, 1) - timedelta(days=1)) if mes_num < 12 else datetime(ano, 12, 31)

        # Filtrar os dados para o mês selecionado
        dados_filtrado = dados[
            (dados['Data de Inserção'] >= min_data) &
            (dados['Data de Inserção'] <= max_data)
        ]

        # Gerar datas completas para o mês selecionado
        datas_mes = pd.date_range(start=min_data, end=max_data, freq='D')

        # Adicionar linhas de controle
        fig.add_trace(go.Scatter(x=datas_mes, y=[media] * len(datas_mes), mode='lines', line=dict(color='green', dash='dash'), name='Média'))
        fig.add_trace(go.Scatter(x=datas_mes, y=[LCS] * len(datas_mes), mode='lines', line=dict(color='red', dash='dash'), name='LCS 3s'))
        fig.add_trace(go.Scatter(x=datas_mes, y=[LCI] * len(datas_mes), mode='lines', line=dict(color='red', dash='dash'), name='LCI 3s'))
        fig.add_trace(go.Scatter(x=datas_mes, y=[LCI_2S] * len(datas_mes), mode='lines', line=dict(color='yellow', dash='dash'), name='LCI 2s'))
        fig.add_trace(go.Scatter(x=datas_mes, y=[LCS_2S] * len(datas_mes), mode='lines', line=dict(color='yellow', dash='dash'), name='LCS 2s'))

        # Adicionar dados do mês (se houver)
        if not dados_filtrado.empty:
            fig.add_trace(go.Scatter(
                x=dados_filtrado['Data de Inserção'],
                y=dados_filtrado['Valor'],
                mode='lines+markers',
                marker=dict(color='blue'),
                name=f'Dados do mês {mes}'
            ))
        else:
            # Nenhum dado disponível
            fig.add_annotation(
                text="Sem dados para o mês selecionado",
                xref="paper", yref="paper",
                x=0.5, y=0.5, showarrow=False,
                font=dict(size=16, color="red")
            )

    return fig

   

def criar_carta_controle_meses_flash_nitrogenio(dados, elemento, mes='Mês Não Especificado'):

    dados = dados.copy()

    dados.loc[:, 'Data de Inserção'] = pd.to_datetime(dados['Data de Inserção'], errors='coerce')
    # Mapeamento de meses
    meses = {
        "Janeiro": 1, "Fevereiro": 2, "Março": 3, "Abril": 4, "Maio": 5, "Junho": 6,
        "Julho": 7, "Agosto": 8, "Setembro": 9, "Outubro": 10, "Novembro": 11, "Dezembro": 12
    }

    # Calcular estatísticas
    media = np.mean(dados_calibracao['Nitrogênio'])
    desvio_padrao = np.std(dados_calibracao['Nitrogênio'])
    LCS = media + 3 * desvio_padrao
    LCI = media - 3 * desvio_padrao
    LCI_2S = media - 2 * desvio_padrao
    LCS_2S = media + 2 * desvio_padrao

    # Criar figura
    fig = go.Figure()

    if mes in meses:
        mes_num = meses[mes]
        ano = datetime.now().year  # Ajustar para o ano atual ou permitir seleção de ano

        # Obter o intervalo de datas para o mês
        min_data = datetime(ano, mes_num, 1)
        max_data = (datetime(ano, mes_num + 1, 1) - timedelta(days=1)) if mes_num < 12 else datetime(ano, 12, 31)

        # Filtrar os dados para o mês selecionado
        dados_filtrado = dados[
            (dados['Data de Inserção'] >= min_data) &
            (dados['Data de Inserção'] <= max_data)
        ]

        # Gerar datas completas para o mês selecionado
        datas_mes = pd.date_range(start=min_data, end=max_data, freq='D')

        # Adicionar linhas de controle
        fig.add_trace(go.Scatter(x=datas_mes, y=[media] * len(datas_mes), mode='lines', line=dict(color='green', dash='dash'), name='Média'))
        fig.add_trace(go.Scatter(x=datas_mes, y=[LCS] * len(datas_mes), mode='lines', line=dict(color='red', dash='dash'), name='LCS 3s'))
        fig.add_trace(go.Scatter(x=datas_mes, y=[LCI] * len(datas_mes), mode='lines', line=dict(color='red', dash='dash'), name='LCI 3s'))
        fig.add_trace(go.Scatter(x=datas_mes, y=[LCI_2S] * len(datas_mes), mode='lines', line=dict(color='yellow', dash='dash'), name='LCI 2s'))
        fig.add_trace(go.Scatter(x=datas_mes, y=[LCS_2S] * len(datas_mes), mode='lines', line=dict(color='yellow', dash='dash'), name='LCS 2s'))

        # Adicionar dados do mês (se houver)
        if not dados_filtrado.empty:
            fig.add_trace(go.Scatter(
                x=dados_filtrado['Data de Inserção'],
                y=dados_filtrado['Valor'],
                mode='lines+markers',
                marker=dict(color='blue'),
                name=f'Dados do mês {mes}'
            ))
        else:
            # Nenhum dado disponível
            fig.add_annotation(
                text="Sem dados para o mês selecionado",
                xref="paper", yref="paper",
                x=0.5, y=0.5, showarrow=False,
                font=dict(size=16, color="red")
            )

    return fig

    

def criar_carta_controle_meses_flash_hidrogenio(dados, elemento, mes='Mês Não Especificado'):

    dados = dados.copy()

    dados.loc[:, 'Data de Inserção'] = pd.to_datetime(dados['Data de Inserção'], errors='coerce')
    # Mapeamento de meses
    meses = {
        "Janeiro": 1, "Fevereiro": 2, "Março": 3, "Abril": 4, "Maio": 5, "Junho": 6,
        "Julho": 7, "Agosto": 8, "Setembro": 9, "Outubro": 10, "Novembro": 11, "Dezembro": 12
    }

    # Calcular estatísticas
    media = np.mean(dados_calibracao['Hidrogênio'])
    desvio_padrao = np.std(dados_calibracao['Hidrogênio'])
    LCS = media + 3 * desvio_padrao
    LCI = media - 3 * desvio_padrao
    LCI_2S = media - 2 * desvio_padrao
    LCS_2S = media + 2 * desvio_padrao

    # Criar figura
    fig = go.Figure()

    if mes in meses:
        mes_num = meses[mes]
        ano = datetime.now().year  # Ajustar para o ano atual ou permitir seleção de ano

        # Obter o intervalo de datas para o mês
        min_data = datetime(ano, mes_num, 1)
        max_data = (datetime(ano, mes_num + 1, 1) - timedelta(days=1)) if mes_num < 12 else datetime(ano, 12, 31)

        # Filtrar os dados para o mês selecionado
        dados_filtrado = dados[
            (dados['Data de Inserção'] >= min_data) &
            (dados['Data de Inserção'] <= max_data)
        ]

        # Gerar datas completas para o mês selecionado
        datas_mes = pd.date_range(start=min_data, end=max_data, freq='D')

        # Adicionar linhas de controle
        fig.add_trace(go.Scatter(x=datas_mes, y=[media] * len(datas_mes), mode='lines', line=dict(color='green', dash='dash'), name='Média'))
        fig.add_trace(go.Scatter(x=datas_mes, y=[LCS] * len(datas_mes), mode='lines', line=dict(color='red', dash='dash'), name='LCS 3s'))
        fig.add_trace(go.Scatter(x=datas_mes, y=[LCI] * len(datas_mes), mode='lines', line=dict(color='red', dash='dash'), name='LCI 3s'))
        fig.add_trace(go.Scatter(x=datas_mes, y=[LCI_2S] * len(datas_mes), mode='lines', line=dict(color='yellow', dash='dash'), name='LCI 2s'))
        fig.add_trace(go.Scatter(x=datas_mes, y=[LCS_2S] * len(datas_mes), mode='lines', line=dict(color='yellow', dash='dash'), name='LCS 2s'))

        # Adicionar dados do mês (se houver)
        if not dados_filtrado.empty:
            fig.add_trace(go.Scatter(
                x=dados_filtrado['Data de Inserção'],
                y=dados_filtrado['Valor'],
                mode='lines+markers',
                marker=dict(color='blue'),
                name=f'Dados do mês {mes}'
            ))
        else:
            # Nenhum dado disponível
            fig.add_annotation(
                text="Sem dados para o mês selecionado",
                xref="paper", yref="paper",
                x=0.5, y=0.5, showarrow=False,
                font=dict(size=16, color="red")
            )

    return fig
    

def criar_carta_controle_meses_flash_enxofre(dados, elemento, mes='Mês Não Especificado'):

    dados = dados.copy()

    dados.loc[:, 'Data de Inserção'] = pd.to_datetime(dados['Data de Inserção'], errors='coerce')
    # Mapeamento de meses
    meses = {
        "Janeiro": 1, "Fevereiro": 2, "Março": 3, "Abril": 4, "Maio": 5, "Junho": 6,
        "Julho": 7, "Agosto": 8, "Setembro": 9, "Outubro": 10, "Novembro": 11, "Dezembro": 12
    }

    # Calcular estatísticas
    media = np.mean(dados_calibracao['Enxofre'])
    desvio_padrao = np.std(dados_calibracao['Enxofre'])
    LCS = media + 3 * desvio_padrao
    LCI = media - 3 * desvio_padrao
    LCI_2S = media - 2 * desvio_padrao
    LCS_2S = media + 2 * desvio_padrao

    # Criar figura
    fig = go.Figure()

    if mes in meses:
        mes_num = meses[mes]
        ano = datetime.now().year  # Ajustar para o ano atual ou permitir seleção de ano

        # Obter o intervalo de datas para o mês
        min_data = datetime(ano, mes_num, 1)
        max_data = (datetime(ano, mes_num + 1, 1) - timedelta(days=1)) if mes_num < 12 else datetime(ano, 12, 31)

        # Filtrar os dados para o mês selecionado
        dados_filtrado = dados[
            (dados['Data de Inserção'] >= min_data) &
            (dados['Data de Inserção'] <= max_data)
        ]

        # Gerar datas completas para o mês selecionado
        datas_mes = pd.date_range(start=min_data, end=max_data, freq='D')

        # Adicionar linhas de controle
        fig.add_trace(go.Scatter(x=datas_mes, y=[media] * len(datas_mes), mode='lines', line=dict(color='green', dash='dash'), name='Média'))
        fig.add_trace(go.Scatter(x=datas_mes, y=[LCS] * len(datas_mes), mode='lines', line=dict(color='red', dash='dash'), name='LCS 3s'))
        fig.add_trace(go.Scatter(x=datas_mes, y=[LCI] * len(datas_mes), mode='lines', line=dict(color='red', dash='dash'), name='LCI 3s'))
        fig.add_trace(go.Scatter(x=datas_mes, y=[LCI_2S] * len(datas_mes), mode='lines', line=dict(color='yellow', dash='dash'), name='LCI 2s'))
        fig.add_trace(go.Scatter(x=datas_mes, y=[LCS_2S] * len(datas_mes), mode='lines', line=dict(color='yellow', dash='dash'), name='LCS 2s'))

        # Adicionar dados do mês (se houver)
        if not dados_filtrado.empty:
            fig.add_trace(go.Scatter(
                x=dados_filtrado['Data de Inserção'],
                y=dados_filtrado['Valor'],
                mode='lines+markers',
                marker=dict(color='blue'),
                name=f'Dados do mês {mes}'
            ))
        else:
            # Nenhum dado disponível
            fig.add_annotation(
                text="Sem dados para o mês selecionado",
                xref="paper", yref="paper",
                x=0.5, y=0.5, showarrow=False,
                font=dict(size=16, color="red")
            )

    return fig
    


###################################################### CARREGAR PLANILHAS ##########################################################

dados_flash = carregar_dados_flash()
dados_multitek_cb = carregar_dados_multitek_cb()
dados_multitek_cm = carregar_dados_multitek_cm()
dados_multitek_ca = carregar_dados_multitek_ca()
dados_pananalytical = carregar_dados_pananalytical()
dados_oxford_frx = carregar_dados_oxford_frx()

####################################################### CARTA CONTROLE MULTITEK CURVA BAIXA #########################################

def criar_carta_controle_multitek_cb(multitek_cb, dados_teste, elemento):
    media = np.mean(multitek_cb)
    desvio_padrao = np.std(multitek_cb)
    LCS = media + 3 * desvio_padrao
    LCI = media - 3 * desvio_padrao
    LCI_2S = media - 2 * desvio_padrao
    LCS_2S = media + 2 * desvio_padrao


    fig = go.Figure()

    # Adicionar os dados de calibração ao gráfico
    fig.add_trace(go.Scatter(x=list(range(len(multitek_cb))), y=multitek_cb,
                             mode='lines+markers', name=f'mg/kg {elemento}', marker=dict(color='blue')))
    

    # Adicionar a média e os limites de controle
    fig.add_trace(go.Scatter(x=[1, len(multitek_cb) + 1], y=[media, media],
                             mode='lines', line=dict(color='green', dash='dash'), name='Média'))
    fig.add_trace(go.Scatter(x=[1, len(multitek_cb) + 1], y=[LCS, LCS],
                             mode='lines', line=dict(color='red', dash='dash'), name='LCS 3s'))
    fig.add_trace(go.Scatter(x=[1, len(multitek_cb) + 1], y=[LCI, LCI],
                             mode='lines', line=dict(color='red', dash='dash'), name='LCI 3s'))
    fig.add_trace(go.Scatter(x=[1, len(multitek_cb) + 1], y=[LCS_2S, LCS_2S],
                             mode='lines', line=dict(color='#FFD700', dash='dash'), name='LCS 2s'))
    fig.add_trace(go.Scatter(x=[1, len(multitek_cb) + 1], y=[LCI_2S, LCI_2S],
                             mode='lines', line=dict(color='#FFD700', dash='dash'), name='LCI 2s'))

    fig.update_layout(xaxis_title='Execução', yaxis_title=f'Resultado ({elemento})')
    return fig


def criar_carta_controle_meses_multitek_cb_enxofre(dados, elemento, mes='Mês Não Especificado'):

    dados = dados.copy()

    dados.loc[:, 'Data de Inserção'] = pd.to_datetime(dados['Data de Inserção'], errors='coerce')
    # Mapeamento de meses
    meses = {
        "Janeiro": 1, "Fevereiro": 2, "Março": 3, "Abril": 4, "Maio": 5, "Junho": 6,
        "Julho": 7, "Agosto": 8, "Setembro": 9, "Outubro": 10, "Novembro": 11, "Dezembro": 12
    }

    # Calcular estatísticas
    media = np.mean(multitek_cb['Enxofre'])
    desvio_padrao = np.std(multitek_cb['Enxofre'])
    LCS = media + 3 * desvio_padrao
    LCI = media - 3 * desvio_padrao
    LCI_2S = media - 2 * desvio_padrao
    LCS_2S = media + 2 * desvio_padrao

    # Criar figura
    fig = go.Figure()

    if mes in meses:
        mes_num = meses[mes]
        ano = datetime.now().year  # Ajustar para o ano atual ou permitir seleção de ano

        # Obter o intervalo de datas para o mês
        min_data = datetime(ano, mes_num, 1)
        max_data = (datetime(ano, mes_num + 1, 1) - timedelta(days=1)) if mes_num < 12 else datetime(ano, 12, 31)

        # Filtrar os dados para o mês selecionado
        dados_filtrado = dados[
            (dados['Data de Inserção'] >= min_data) &
            (dados['Data de Inserção'] <= max_data)
        ]

        # Gerar datas completas para o mês selecionado
        datas_mes = pd.date_range(start=min_data, end=max_data, freq='D')

        # Adicionar linhas de controle
        fig.add_trace(go.Scatter(x=datas_mes, y=[media] * len(datas_mes), mode='lines', line=dict(color='green', dash='dash'), name='Média'))
        fig.add_trace(go.Scatter(x=datas_mes, y=[LCS] * len(datas_mes), mode='lines', line=dict(color='red', dash='dash'), name='LCS 3s'))
        fig.add_trace(go.Scatter(x=datas_mes, y=[LCI] * len(datas_mes), mode='lines', line=dict(color='red', dash='dash'), name='LCI 3s'))
        fig.add_trace(go.Scatter(x=datas_mes, y=[LCI_2S] * len(datas_mes), mode='lines', line=dict(color='yellow', dash='dash'), name='LCI 2s'))
        fig.add_trace(go.Scatter(x=datas_mes, y=[LCS_2S] * len(datas_mes), mode='lines', line=dict(color='yellow', dash='dash'), name='LCS 2s'))

        # Adicionar dados do mês (se houver)
        if not dados_filtrado.empty:
            fig.add_trace(go.Scatter(
                x=dados_filtrado['Data de Inserção'],
                y=dados_filtrado['Valor'],
                mode='lines+markers',
                marker=dict(color='blue'),
                name=f'Dados do mês {mes}'
            ))
        else:
            # Nenhum dado disponível
            fig.add_annotation(
                text="Sem dados para o mês selecionado",
                xref="paper", yref="paper",
                x=0.5, y=0.5, showarrow=False,
                font=dict(size=16, color="red")
            )

    return fig

def criar_carta_controle_meses_multitek_cb_nitrogenio(dados, elemento, mes='Mês Não Especificado'):

    dados = dados.copy()

    dados.loc[:, 'Data de Inserção'] = pd.to_datetime(dados['Data de Inserção'], errors='coerce')
    # Mapeamento de meses
    meses = {
        "Janeiro": 1, "Fevereiro": 2, "Março": 3, "Abril": 4, "Maio": 5, "Junho": 6,
        "Julho": 7, "Agosto": 8, "Setembro": 9, "Outubro": 10, "Novembro": 11, "Dezembro": 12
    }

    # Calcular estatísticas
    media = np.mean(multitek_cb['Nitrogênio'])
    desvio_padrao = np.std(multitek_cb['Nitrogênio'])
    LCS = media + 3 * desvio_padrao
    LCI = media - 3 * desvio_padrao
    LCI_2S = media - 2 * desvio_padrao
    LCS_2S = media + 2 * desvio_padrao

    # Criar figura
    fig = go.Figure()

    if mes in meses:
        mes_num = meses[mes]
        ano = datetime.now().year  # Ajustar para o ano atual ou permitir seleção de ano

        # Obter o intervalo de datas para o mês
        min_data = datetime(ano, mes_num, 1)
        max_data = (datetime(ano, mes_num + 1, 1) - timedelta(days=1)) if mes_num < 12 else datetime(ano, 12, 31)

        # Filtrar os dados para o mês selecionado
        dados_filtrado = dados[
            (dados['Data de Inserção'] >= min_data) &
            (dados['Data de Inserção'] <= max_data)
        ]

        # Gerar datas completas para o mês selecionado
        datas_mes = pd.date_range(start=min_data, end=max_data, freq='D')

        # Adicionar linhas de controle
        fig.add_trace(go.Scatter(x=datas_mes, y=[media] * len(datas_mes), mode='lines', line=dict(color='green', dash='dash'), name='Média'))
        fig.add_trace(go.Scatter(x=datas_mes, y=[LCS] * len(datas_mes), mode='lines', line=dict(color='red', dash='dash'), name='LCS 3s'))
        fig.add_trace(go.Scatter(x=datas_mes, y=[LCI] * len(datas_mes), mode='lines', line=dict(color='red', dash='dash'), name='LCI 3s'))
        fig.add_trace(go.Scatter(x=datas_mes, y=[LCI_2S] * len(datas_mes), mode='lines', line=dict(color='yellow', dash='dash'), name='LCI 2s'))
        fig.add_trace(go.Scatter(x=datas_mes, y=[LCS_2S] * len(datas_mes), mode='lines', line=dict(color='yellow', dash='dash'), name='LCS 2s'))

        # Adicionar dados do mês (se houver)
        if not dados_filtrado.empty:
            fig.add_trace(go.Scatter(
                x=dados_filtrado['Data de Inserção'],
                y=dados_filtrado['Valor'],
                mode='lines+markers',
                marker=dict(color='blue'),
                name=f'Dados do mês {mes}'
            ))
        else:
            # Nenhum dado disponível
            fig.add_annotation(
                text="Sem dados para o mês selecionado",
                xref="paper", yref="paper",
                x=0.5, y=0.5, showarrow=False,
                font=dict(size=16, color="red")
            )

    return fig
    


#################################################### CARTA MULTITEK CURVA MÉDIA ######################################################


def criar_carta_controle_multitek_cm(multitek_cm, dados_teste, elemento):
    media = np.mean(multitek_cm)
    desvio_padrao = np.std(multitek_cm)
    LCS = media + 3 * desvio_padrao
    LCI = media - 3 * desvio_padrao
    LCI_2S = media - 2 * desvio_padrao
    LCS_2S = media + 2 * desvio_padrao


    fig = go.Figure()

    # Adicionar os dados de calibração ao gráfico
    fig.add_trace(go.Scatter(x=list(range(len(multitek_cm))), y=multitek_cm,
                             mode='lines+markers', name=f'mg/kg {elemento}', marker=dict(color='blue')))
    

    # Adicionar a média e os limites de controle
    fig.add_trace(go.Scatter(x=[1, len(multitek_cm) + 1], y=[media, media],
                             mode='lines', line=dict(color='green', dash='dash'), name='Média'))
    fig.add_trace(go.Scatter(x=[1, len(multitek_cm) + 1], y=[LCS, LCS],
                             mode='lines', line=dict(color='red', dash='dash'), name='LCS 3s'))
    fig.add_trace(go.Scatter(x=[1, len(multitek_cm) + 1], y=[LCI, LCI],
                             mode='lines', line=dict(color='red', dash='dash'), name='LCI 3s'))
    fig.add_trace(go.Scatter(x=[1, len(multitek_cm) + 1], y=[LCS_2S, LCS_2S],
                             mode='lines', line=dict(color='#FFD700', dash='dash'), name='LCS 2s'))
    fig.add_trace(go.Scatter(x=[1, len(multitek_cm) + 1], y=[LCI_2S, LCI_2S],
                             mode='lines', line=dict(color='#FFD700', dash='dash'), name='LCI 2s'))

    fig.update_layout(xaxis_title='Execução', yaxis_title=f'mg/kg ({elemento})')
    return fig


def criar_carta_controle_meses_multitek_cm_enxofre(dados, elemento, mes='Mês Não Especificado'):
    # Calcular média e desvio padrão dos valores de calibração
    dados = dados.copy()

    dados.loc[:, 'Data de Inserção'] = pd.to_datetime(dados['Data de Inserção'], errors='coerce')
    # Mapeamento de meses
    meses = {
        "Janeiro": 1, "Fevereiro": 2, "Março": 3, "Abril": 4, "Maio": 5, "Junho": 6,
        "Julho": 7, "Agosto": 8, "Setembro": 9, "Outubro": 10, "Novembro": 11, "Dezembro": 12
    }

    # Calcular estatísticas
    media = np.mean(multitek_cm['Enxofre'])
    desvio_padrao = np.std(multitek_cm['Enxofre'])
    LCS = media + 3 * desvio_padrao
    LCI = media - 3 * desvio_padrao
    LCI_2S = media - 2 * desvio_padrao
    LCS_2S = media + 2 * desvio_padrao

    # Criar figura
    fig = go.Figure()

    if mes in meses:
        mes_num = meses[mes]
        ano = datetime.now().year  # Ajustar para o ano atual ou permitir seleção de ano

        # Obter o intervalo de datas para o mês
        min_data = datetime(ano, mes_num, 1)
        max_data = (datetime(ano, mes_num + 1, 1) - timedelta(days=1)) if mes_num < 12 else datetime(ano, 12, 31)

        # Filtrar os dados para o mês selecionado
        dados_filtrado = dados[
            (dados['Data de Inserção'] >= min_data) &
            (dados['Data de Inserção'] <= max_data)
        ]

        # Gerar datas completas para o mês selecionado
        datas_mes = pd.date_range(start=min_data, end=max_data, freq='D')

        # Adicionar linhas de controle
        fig.add_trace(go.Scatter(x=datas_mes, y=[media] * len(datas_mes), mode='lines', line=dict(color='green', dash='dash'), name='Média'))
        fig.add_trace(go.Scatter(x=datas_mes, y=[LCS] * len(datas_mes), mode='lines', line=dict(color='red', dash='dash'), name='LCS 3s'))
        fig.add_trace(go.Scatter(x=datas_mes, y=[LCI] * len(datas_mes), mode='lines', line=dict(color='red', dash='dash'), name='LCI 3s'))
        fig.add_trace(go.Scatter(x=datas_mes, y=[LCI_2S] * len(datas_mes), mode='lines', line=dict(color='yellow', dash='dash'), name='LCI 2s'))
        fig.add_trace(go.Scatter(x=datas_mes, y=[LCS_2S] * len(datas_mes), mode='lines', line=dict(color='yellow', dash='dash'), name='LCS 2s'))

        # Adicionar dados do mês (se houver)
        if not dados_filtrado.empty:
            fig.add_trace(go.Scatter(
                x=dados_filtrado['Data de Inserção'],
                y=dados_filtrado['Valor'],
                mode='lines+markers',
                marker=dict(color='blue'),
                name=f'Dados do mês {mes}'
            ))
        else:
            # Nenhum dado disponível
            fig.add_annotation(
                text="Sem dados para o mês selecionado",
                xref="paper", yref="paper",
                x=0.5, y=0.5, showarrow=False,
                font=dict(size=16, color="red")
            )

    return fig

def criar_carta_controle_meses_multitek_cm_nitrogenio(dados, elemento, mes='Mês Não Especificado'):
    dados = dados.copy()

    dados.loc[:, 'Data de Inserção'] = pd.to_datetime(dados['Data de Inserção'], errors='coerce')
    # Mapeamento de meses
    meses = {
        "Janeiro": 1, "Fevereiro": 2, "Março": 3, "Abril": 4, "Maio": 5, "Junho": 6,
        "Julho": 7, "Agosto": 8, "Setembro": 9, "Outubro": 10, "Novembro": 11, "Dezembro": 12
    }

    # Calcular estatísticas
    media = np.mean(multitek_cm['Nitrogênio'])
    desvio_padrao = np.std(multitek_cm['Nitrogênio'])
    LCS = media + 3 * desvio_padrao
    LCI = media - 3 * desvio_padrao
    LCI_2S = media - 2 * desvio_padrao
    LCS_2S = media + 2 * desvio_padrao

    # Criar figura
    fig = go.Figure()

    if mes in meses:
        mes_num = meses[mes]
        ano = datetime.now().year  # Ajustar para o ano atual ou permitir seleção de ano

        # Obter o intervalo de datas para o mês
        min_data = datetime(ano, mes_num, 1)
        max_data = (datetime(ano, mes_num + 1, 1) - timedelta(days=1)) if mes_num < 12 else datetime(ano, 12, 31)

        # Filtrar os dados para o mês selecionado
        dados_filtrado = dados[
            (dados['Data de Inserção'] >= min_data) &
            (dados['Data de Inserção'] <= max_data)
        ]

        # Gerar datas completas para o mês selecionado
        datas_mes = pd.date_range(start=min_data, end=max_data, freq='D')

        # Adicionar linhas de controle
        fig.add_trace(go.Scatter(x=datas_mes, y=[media] * len(datas_mes), mode='lines', line=dict(color='green', dash='dash'), name='Média'))
        fig.add_trace(go.Scatter(x=datas_mes, y=[LCS] * len(datas_mes), mode='lines', line=dict(color='red', dash='dash'), name='LCS 3s'))
        fig.add_trace(go.Scatter(x=datas_mes, y=[LCI] * len(datas_mes), mode='lines', line=dict(color='red', dash='dash'), name='LCI 3s'))
        fig.add_trace(go.Scatter(x=datas_mes, y=[LCI_2S] * len(datas_mes), mode='lines', line=dict(color='yellow', dash='dash'), name='LCI 2s'))
        fig.add_trace(go.Scatter(x=datas_mes, y=[LCS_2S] * len(datas_mes), mode='lines', line=dict(color='yellow', dash='dash'), name='LCS 2s'))

        # Adicionar dados do mês (se houver)
        if not dados_filtrado.empty:
            fig.add_trace(go.Scatter(
                x=dados_filtrado['Data de Inserção'],
                y=dados_filtrado['Valor'],
                mode='lines+markers',
                marker=dict(color='blue'),
                name=f'Dados do mês {mes}'
            ))
        else:
            # Nenhum dado disponível
            fig.add_annotation(
                text="Sem dados para o mês selecionado",
                xref="paper", yref="paper",
                x=0.5, y=0.5, showarrow=False,
                font=dict(size=16, color="red")
            )

    return fig




####################################################### MULTITEK CRUVA ALTA ######################################################

def criar_carta_controle_multitek_ca(multitek_ca, dados_teste, elemento):
    media = np.mean(multitek_ca)
    desvio_padrao = np.std(multitek_ca)
    LCS = media + 3 * desvio_padrao
    LCI = media - 3 * desvio_padrao
    LCI_2S = media - 2 * desvio_padrao
    LCS_2S = media + 2 * desvio_padrao


    fig = go.Figure()

    # Adicionar os dados de calibração ao gráfico
    fig.add_trace(go.Scatter(x=list(range(len(multitek_ca))), y=multitek_ca,
                             mode='lines+markers', name=f'mg/kg {elemento}', marker=dict(color='blue')))
    

    # Adicionar a média e os limites de controle
    fig.add_trace(go.Scatter(x=[1, len(multitek_ca) + 1], y=[media, media],
                             mode='lines', line=dict(color='green', dash='dash'), name='Média'))
    fig.add_trace(go.Scatter(x=[1, len(multitek_ca) + 1], y=[LCS, LCS],
                             mode='lines', line=dict(color='red', dash='dash'), name='LCS 3s'))
    fig.add_trace(go.Scatter(x=[1, len(multitek_ca) + 1], y=[LCI, LCI],
                             mode='lines', line=dict(color='red', dash='dash'), name='LCI 3s'))
    fig.add_trace(go.Scatter(x=[1, len(multitek_ca) + 1], y=[LCS_2S, LCS_2S],
                             mode='lines', line=dict(color='#FFD700', dash='dash'), name='LCS 2s'))
    fig.add_trace(go.Scatter(x=[1, len(multitek_ca) + 1], y=[LCI_2S, LCI_2S],
                             mode='lines', line=dict(color='#FFD700', dash='dash'), name='LCI 2s'))

    fig.update_layout(xaxis_title='Execução', yaxis_title=f'mg/kg ({elemento})')
    return fig


def criar_carta_controle_meses_multitek_ca_enxofre(dados, elemento, mes='Mês Não Especificado'):
    # Calcular média e desvio padrão dos valores de calibração
    dados = dados.copy()

    dados.loc[:, 'Data de Inserção'] = pd.to_datetime(dados['Data de Inserção'], errors='coerce')
    # Mapeamento de meses
    meses = {
        "Janeiro": 1, "Fevereiro": 2, "Março": 3, "Abril": 4, "Maio": 5, "Junho": 6,
        "Julho": 7, "Agosto": 8, "Setembro": 9, "Outubro": 10, "Novembro": 11, "Dezembro": 12
    }

    # Calcular estatísticas
    media = np.mean(multitek_ca['Enxofre'])
    desvio_padrao = np.std(multitek_ca['Enxofre'])
    LCS = media + 3 * desvio_padrao
    LCI = media - 3 * desvio_padrao
    LCI_2S = media - 2 * desvio_padrao
    LCS_2S = media + 2 * desvio_padrao

    # Criar figura
    fig = go.Figure()

    if mes in meses:
        mes_num = meses[mes]
        ano = datetime.now().year  # Ajustar para o ano atual ou permitir seleção de ano

        # Obter o intervalo de datas para o mês
        min_data = datetime(ano, mes_num, 1)
        max_data = (datetime(ano, mes_num + 1, 1) - timedelta(days=1)) if mes_num < 12 else datetime(ano, 12, 31)

        # Filtrar os dados para o mês selecionado
        dados_filtrado = dados[
            (dados['Data de Inserção'] >= min_data) &
            (dados['Data de Inserção'] <= max_data)
        ]

        # Gerar datas completas para o mês selecionado
        datas_mes = pd.date_range(start=min_data, end=max_data, freq='D')

        # Adicionar linhas de controle
        fig.add_trace(go.Scatter(x=datas_mes, y=[media] * len(datas_mes), mode='lines', line=dict(color='green', dash='dash'), name='Média'))
        fig.add_trace(go.Scatter(x=datas_mes, y=[LCS] * len(datas_mes), mode='lines', line=dict(color='red', dash='dash'), name='LCS 3s'))
        fig.add_trace(go.Scatter(x=datas_mes, y=[LCI] * len(datas_mes), mode='lines', line=dict(color='red', dash='dash'), name='LCI 3s'))
        fig.add_trace(go.Scatter(x=datas_mes, y=[LCI_2S] * len(datas_mes), mode='lines', line=dict(color='yellow', dash='dash'), name='LCI 2s'))
        fig.add_trace(go.Scatter(x=datas_mes, y=[LCS_2S] * len(datas_mes), mode='lines', line=dict(color='yellow', dash='dash'), name='LCS 2s'))

        # Adicionar dados do mês (se houver)
        if not dados_filtrado.empty:
            fig.add_trace(go.Scatter(
                x=dados_filtrado['Data de Inserção'],
                y=dados_filtrado['Valor'],
                mode='lines+markers',
                marker=dict(color='blue'),
                name=f'Dados do mês {mes}'
            ))
        else:
            # Nenhum dado disponível
            fig.add_annotation(
                text="Sem dados para o mês selecionado",
                xref="paper", yref="paper",
                x=0.5, y=0.5, showarrow=False,
                font=dict(size=16, color="red")
            )

    return fig

def criar_carta_controle_meses_multitek_ca_nitrogenio(dados, elemento, mes='Mês Não Especificado'):
    dados = dados.copy()

    dados.loc[:, 'Data de Inserção'] = pd.to_datetime(dados['Data de Inserção'], errors='coerce')
    # Mapeamento de meses
    meses = {
        "Janeiro": 1, "Fevereiro": 2, "Março": 3, "Abril": 4, "Maio": 5, "Junho": 6,
        "Julho": 7, "Agosto": 8, "Setembro": 9, "Outubro": 10, "Novembro": 11, "Dezembro": 12
    }

    # Calcular estatísticas
    media = np.mean(multitek_ca['Nitrogênio'])
    desvio_padrao = np.std(multitek_ca['Nitrogênio'])
    LCS = media + 3 * desvio_padrao
    LCI = media - 3 * desvio_padrao
    LCI_2S = media - 2 * desvio_padrao
    LCS_2S = media + 2 * desvio_padrao

    # Criar figura
    fig = go.Figure()

    if mes in meses:
        mes_num = meses[mes]
        ano = datetime.now().year  # Ajustar para o ano atual ou permitir seleção de ano

        # Obter o intervalo de datas para o mês
        min_data = datetime(ano, mes_num, 1)
        max_data = (datetime(ano, mes_num + 1, 1) - timedelta(days=1)) if mes_num < 12 else datetime(ano, 12, 31)

        # Filtrar os dados para o mês selecionado
        dados_filtrado = dados[
            (dados['Data de Inserção'] >= min_data) &
            (dados['Data de Inserção'] <= max_data)
        ]

        # Gerar datas completas para o mês selecionado
        datas_mes = pd.date_range(start=min_data, end=max_data, freq='D')

        # Adicionar linhas de controle
        fig.add_trace(go.Scatter(x=datas_mes, y=[media] * len(datas_mes), mode='lines', line=dict(color='green', dash='dash'), name='Média'))
        fig.add_trace(go.Scatter(x=datas_mes, y=[LCS] * len(datas_mes), mode='lines', line=dict(color='red', dash='dash'), name='LCS 3s'))
        fig.add_trace(go.Scatter(x=datas_mes, y=[LCI] * len(datas_mes), mode='lines', line=dict(color='red', dash='dash'), name='LCI 3s'))
        fig.add_trace(go.Scatter(x=datas_mes, y=[LCI_2S] * len(datas_mes), mode='lines', line=dict(color='yellow', dash='dash'), name='LCI 2s'))
        fig.add_trace(go.Scatter(x=datas_mes, y=[LCS_2S] * len(datas_mes), mode='lines', line=dict(color='yellow', dash='dash'), name='LCS 2s'))

        # Adicionar dados do mês (se houver)
        if not dados_filtrado.empty:
            fig.add_trace(go.Scatter(
                x=dados_filtrado['Data de Inserção'],
                y=dados_filtrado['Valor'],
                mode='lines+markers',
                marker=dict(color='blue'),
                name=f'Dados do mês {mes}'
            ))
        else:
            # Nenhum dado disponível
            fig.add_annotation(
                text="Sem dados para o mês selecionado",
                xref="paper", yref="paper",
                x=0.5, y=0.5, showarrow=False,
                font=dict(size=16, color="red")
            )

    return fig


####################################################### CARTA CONTROLE PANANALYTICAL #############################################
def criar_carta_controle_pananalytical(pananalytical, dados_teste, elemento):
    media = np.mean(pananalytical)
    desvio_padrao = np.std(pananalytical)
    LCS = media + 3 * desvio_padrao
    LCI = media - 3 * desvio_padrao
    LCI_2S = media - 2 * desvio_padrao
    LCS_2S = media + 2 * desvio_padrao


    fig = go.Figure()

    # Adicionar os dados de calibração ao gráfico
    fig.add_trace(go.Scatter(x=list(range(len(pananalytical))), y=pananalytical,
                             mode='lines+markers', name=f'% {elemento}', marker=dict(color='blue')))
    

    # Adicionar a média e os limites de controle
    fig.add_trace(go.Scatter(x=[1, len(pananalytical) + 1], y=[media, media],
                             mode='lines', line=dict(color='green', dash='dash'), name='Média'))
    fig.add_trace(go.Scatter(x=[1, len(pananalytical) + 1], y=[LCS, LCS],
                             mode='lines', line=dict(color='red', dash='dash'), name='LCS 3s'))
    fig.add_trace(go.Scatter(x=[1, len(pananalytical) + 1], y=[LCI, LCI],
                             mode='lines', line=dict(color='red', dash='dash'), name='LCI 3s'))
    fig.add_trace(go.Scatter(x=[1, len(pananalytical) + 1], y=[LCS_2S, LCS_2S],
                             mode='lines', line=dict(color='#FFD700', dash='dash'), name='LCS 2s'))
    fig.add_trace(go.Scatter(x=[1, len(pananalytical) + 1], y=[LCI_2S, LCI_2S],
                             mode='lines', line=dict(color='#FFD700', dash='dash'), name='LCI 2s'))

    fig.update_layout(xaxis_title='Execução', yaxis_title=f'% ({elemento})')
    return fig


def criar_carta_controle_meses_pananalytical(dados, elemento, mes='Mês Não Especificado'):
    # Calcular média e desvio padrão dos valores de calibração
    dados = dados.copy()

    dados.loc[:, 'Data de Inserção'] = pd.to_datetime(dados['Data de Inserção'], errors='coerce')
    # Mapeamento de meses
    meses = {
        "Janeiro": 1, "Fevereiro": 2, "Março": 3, "Abril": 4, "Maio": 5, "Junho": 6,
        "Julho": 7, "Agosto": 8, "Setembro": 9, "Outubro": 10, "Novembro": 11, "Dezembro": 12
    }

    # Calcular estatísticas
    media = np.mean(pananalytical['Enxofre'])
    desvio_padrao = np.std(pananalytical['Enxofre'])
    LCS = media + 3 * desvio_padrao
    LCI = media - 3 * desvio_padrao
    LCI_2S = media - 2 * desvio_padrao
    LCS_2S = media + 2 * desvio_padrao

    # Criar figura
    fig = go.Figure()

    if mes in meses:
        mes_num = meses[mes]
        ano = datetime.now().year  # Ajustar para o ano atual ou permitir seleção de ano

        # Obter o intervalo de datas para o mês
        min_data = datetime(ano, mes_num, 1)
        max_data = (datetime(ano, mes_num + 1, 1) - timedelta(days=1)) if mes_num < 12 else datetime(ano, 12, 31)

        # Filtrar os dados para o mês selecionado
        dados_filtrado = dados[
            (dados['Data de Inserção'] >= min_data) &
            (dados['Data de Inserção'] <= max_data)
        ]

        # Gerar datas completas para o mês selecionado
        datas_mes = pd.date_range(start=min_data, end=max_data, freq='D')

        # Adicionar linhas de controle
        fig.add_trace(go.Scatter(x=datas_mes, y=[media] * len(datas_mes), mode='lines', line=dict(color='green', dash='dash'), name='Média'))
        fig.add_trace(go.Scatter(x=datas_mes, y=[LCS] * len(datas_mes), mode='lines', line=dict(color='red', dash='dash'), name='LCS 3s'))
        fig.add_trace(go.Scatter(x=datas_mes, y=[LCI] * len(datas_mes), mode='lines', line=dict(color='red', dash='dash'), name='LCI 3s'))
        fig.add_trace(go.Scatter(x=datas_mes, y=[LCI_2S] * len(datas_mes), mode='lines', line=dict(color='yellow', dash='dash'), name='LCI 2s'))
        fig.add_trace(go.Scatter(x=datas_mes, y=[LCS_2S] * len(datas_mes), mode='lines', line=dict(color='yellow', dash='dash'), name='LCS 2s'))

        # Adicionar dados do mês (se houver)
        if not dados_filtrado.empty:
            fig.add_trace(go.Scatter(
                x=dados_filtrado['Data de Inserção'],
                y=dados_filtrado['Valor'],
                mode='lines+markers',
                marker=dict(color='blue'),
                name=f'Dados do mês {mes}'
            ))
        else:
            # Nenhum dado disponível
            fig.add_annotation(
                text="Sem dados para o mês selecionado",
                xref="paper", yref="paper",
                x=0.5, y=0.5, showarrow=False,
                font=dict(size=16, color="red")
            )

    return fig

######################################################## OXFORD FRX #################################################################


def criar_carta_controle_oxford(oxford_frx, dados_teste, elemento):
    media = np.mean(oxford_frx)
    desvio_padrao = np.std(oxford_frx)
    LCS = media + 3 * desvio_padrao
    LCI = media - 3 * desvio_padrao
    LCI_2S = media - 2 * desvio_padrao
    LCS_2S = media + 2 * desvio_padrao


    fig = go.Figure()

    # Adicionar os dados de calibração ao gráfico
    fig.add_trace(go.Scatter(x=list(range(len(oxford_frx))), y=oxford_frx,
                             mode='lines+markers', name=f'mg/kg {elemento}', marker=dict(color='blue')))
    

    # Adicionar a média e os limites de controle
    fig.add_trace(go.Scatter(x=[1, len(oxford_frx) + 1], y=[media, media],
                             mode='lines', line=dict(color='green', dash='dash'), name='Média'))
    fig.add_trace(go.Scatter(x=[1, len(oxford_frx) + 1], y=[LCS, LCS],
                             mode='lines', line=dict(color='red', dash='dash'), name='LCS 3s'))
    fig.add_trace(go.Scatter(x=[1, len(oxford_frx) + 1], y=[LCI, LCI],
                             mode='lines', line=dict(color='red', dash='dash'), name='LCI 3s'))
    fig.add_trace(go.Scatter(x=[1, len(oxford_frx) + 1], y=[LCS_2S, LCS_2S],
                             mode='lines', line=dict(color='#FFD700', dash='dash'), name='LCS 2s'))
    fig.add_trace(go.Scatter(x=[1, len(oxford_frx) + 1], y=[LCI_2S, LCI_2S],
                             mode='lines', line=dict(color='#FFD700', dash='dash'), name='LCI 2s'))

    fig.update_layout(xaxis_title='Execução', yaxis_title=f'mg/kg ({elemento})')
    return fig


def criar_carta_controle_meses_oxford(dados, elemento, mes='Mês Não Especificado'):
    # Calcular média e desvio padrão dos valores de calibração
    dados = dados.copy()

    dados.loc[:, 'Data de Inserção'] = pd.to_datetime(dados['Data de Inserção'], errors='coerce')
    # Mapeamento de meses
    meses = {
        "Janeiro": 1, "Fevereiro": 2, "Março": 3, "Abril": 4, "Maio": 5, "Junho": 6,
        "Julho": 7, "Agosto": 8, "Setembro": 9, "Outubro": 10, "Novembro": 11, "Dezembro": 12
    }

    # Calcular estatísticas
    media = np.mean(oxford_frx['Arsênio'])
    desvio_padrao = np.std(oxford_frx['Arsênio'])
    LCS = media + 3 * desvio_padrao
    LCI = media - 3 * desvio_padrao
    LCI_2S = media - 2 * desvio_padrao
    LCS_2S = media + 2 * desvio_padrao

    # Criar figura
    fig = go.Figure()

    if mes in meses:
        mes_num = meses[mes]
        ano = datetime.now().year  # Ajustar para o ano atual ou permitir seleção de ano

        # Obter o intervalo de datas para o mês
        min_data = datetime(ano, mes_num, 1)
        max_data = (datetime(ano, mes_num + 1, 1) - timedelta(days=1)) if mes_num < 12 else datetime(ano, 12, 31)

        # Filtrar os dados para o mês selecionado
        dados_filtrado = dados[
            (dados['Data de Inserção'] >= min_data) &
            (dados['Data de Inserção'] <= max_data)
        ]

        # Gerar datas completas para o mês selecionado
        datas_mes = pd.date_range(start=min_data, end=max_data, freq='D')

        # Adicionar linhas de controle
        fig.add_trace(go.Scatter(x=datas_mes, y=[media] * len(datas_mes), mode='lines', line=dict(color='green', dash='dash'), name='Média'))
        fig.add_trace(go.Scatter(x=datas_mes, y=[LCS] * len(datas_mes), mode='lines', line=dict(color='red', dash='dash'), name='LCS 3s'))
        fig.add_trace(go.Scatter(x=datas_mes, y=[LCI] * len(datas_mes), mode='lines', line=dict(color='red', dash='dash'), name='LCI 3s'))
        fig.add_trace(go.Scatter(x=datas_mes, y=[LCI_2S] * len(datas_mes), mode='lines', line=dict(color='yellow', dash='dash'), name='LCI 2s'))
        fig.add_trace(go.Scatter(x=datas_mes, y=[LCS_2S] * len(datas_mes), mode='lines', line=dict(color='yellow', dash='dash'), name='LCS 2s'))

        # Adicionar dados do mês (se houver)
        if not dados_filtrado.empty:
            fig.add_trace(go.Scatter(
                x=dados_filtrado['Data de Inserção'],
                y=dados_filtrado['Valor'],
                mode='lines+markers',
                marker=dict(color='blue'),
                name=f'Dados do mês {mes}'
            ))
        else:
            # Nenhum dado disponível
            fig.add_annotation(
                text="Sem dados para o mês selecionado",
                xref="paper", yref="paper",
                x=0.5, y=0.5, showarrow=False,
                font=dict(size=16, color="red")
            )

    return fig



######################################################### LAYOUT DA PÁGINA INICIAL ###################################################

content = html.Div(id="page-content", style={"margin-left": "220px", "padding": "20px"})


####################################################DROPDOWN E PÁGINA DE CADA GRÁFICO FLASH 2000 #####################################


def gerar_layout_grafico_flash2000(elemento):
    meses = {
        "Janeiro": 1, "Fevereiro": 2, "Março": 3, "Abril": 4, "Maio": 5, "Junho": 6,
        "Julho": 7, "Agosto": 8, "Setembro": 9, "Outubro": 10, "Novembro": 11, "Dezembro": 12
    }

    # Valor inicial do dropdown definido para o mês atual
    mes_atual = datetime.now().month
    nome_mes_atual = [nome for nome, num in meses.items() if num == mes_atual][0]

    return html.Div([
        html.H1(f'Dashboard de Carta Controle - {elemento}'),
        html.Br(),
        
        # Div com texto "Selecione o mês" e o dropdown para selecionar o mês
        html.Div([
            html.Span("Selecione o mês:", style={"margin-right": "10px", "font-size": "16px", "font-weight": "bold"}),
            dcc.Dropdown(
                id=f'dropdown-mes-{elemento}',
                value=nome_mes_atual,  # Nome do mês atual
                options=[{"label": nome, "value": nome} for nome in meses.keys()],
                placeholder="Escolha um mês",
                style={"width": "200px"}
            ),
        ], style={"display": "flex", "align-items": "center"}),

        html.Br(),
        dcc.Input(id=f'input-valor-{elemento}', type='number', placeholder=f'% de {elemento}', debounce=True),
        html.Button('Adicionar Valor', id=f'btn-adicionar-{elemento}', n_clicks=0),
        
        dcc.Graph(id=f'grafico-controle-{elemento}', figure=criar_carta_controle(dados_calibracao[elemento], [], elemento)),
    ])


############################################################### DROPDOWN E PÁGINA GRÁFICO MULTITEK CURVA BAIXA #####################

def gerar_layout_grafico_multitek_cb(elemento):
    return html.Div([
        html.H1(f'Dashboard de Carta Controle - {elemento}'),
        html.Br(),
        
        # Div com texto "Selecione o mês" e o dropdown para selecionar o mês
        html.Div([
            html.Span("Selecione o mês:", style={"margin-right": "10px", "font-size": "16px", "font-weight": "bold"}),
            dcc.Dropdown(
                id=f'dropdown-mes-multitek_cb-{elemento}',
                options=[
                    {"label": "Janeiro", "value": "Janeiro"},
                    {"label": "Fevereiro", "value": "Fevereiro"},
                    {"label": "Março", "value": "Março"},
                    {"label": "Abril", "value": "Abril"},
                    {"label": "Maio", "value": "Maio"},
                    {"label": "Junho", "value": "Junho"},
                    {"label": "Julho", "value": "Julho"},
                    {"label": "Agosto", "value": "Agosto"},
                    {"label": "Setembro", "value": "Setembro"},
                    {"label": "Outubro", "value": "Outubro"},
                    {"label": "Novembro", "value": "Novembro"},
                    {"label": "Dezembro", "value": "Dezembro"},
                ],
                placeholder="Escolha um mês",
                style={"width": "200px"}
            ),
        ], style={"display": "flex", "align-items": "center"}),

        html.Br(),
        dcc.Input(id=f'input-valor-multitek_cb-{elemento}', type='number', placeholder=f'mg/kg de {elemento}', debounce=True),
        html.Button('Adicionar Valor', id=f'btn-adicionar-multitek_cb-{elemento}', n_clicks=0),
        
        dcc.Graph(id=f'grafico-multitek_cb-{elemento}', 
                  figure=criar_carta_controle_multitek_cb(multitek_cb[elemento], [], elemento)),
    ])




############################################DROPDOWN E PÁGINA GRÁFICO MULTITEK CURVA MÉDIA ###########################################

def gerar_layout_grafico_multitek_cm(elemento):
    return html.Div([
        html.H1(f'Dashboard de Carta Controle - {elemento}'),
        html.Br(),
        
        # Div com texto "Selecione o mês" e o dropdown para selecionar o mês
        html.Div([
            html.Span("Selecione o mês:", style={"margin-right": "10px", "font-size": "16px", "font-weight": "bold"}),
            dcc.Dropdown(
                id=f'dropdown-mes-multitek_cm-{elemento}',
                options=[
                    {"label": "Janeiro", "value": "Janeiro"},
                    {"label": "Fevereiro", "value": "Fevereiro"},
                    {"label": "Março", "value": "Março"},
                    {"label": "Abril", "value": "Abril"},
                    {"label": "Maio", "value": "Maio"},
                    {"label": "Junho", "value": "Junho"},
                    {"label": "Julho", "value": "Julho"},
                    {"label": "Agosto", "value": "Agosto"},
                    {"label": "Setembro", "value": "Setembro"},
                    {"label": "Outubro", "value": "Outubro"},
                    {"label": "Novembro", "value": "Novembro"},
                    {"label": "Dezembro", "value": "Dezembro"},
                ],
                placeholder="Escolha um mês",
                style={"width": "200px"}
            ),
        ], style={"display": "flex", "align-items": "center"}),

        html.Br(),
        dcc.Input(id=f'input-valor-multitek_cm-{elemento}', type='number', placeholder=f'mg/kg de {elemento}', debounce=True),
        html.Button('Adicionar Valor', id=f'btn-adicionar-multitek_cm-{elemento}', n_clicks=0),
        
        dcc.Graph(id=f'grafico-multitek_cm-{elemento}', 
                  figure=criar_carta_controle_multitek_cm(multitek_cm[elemento], [], elemento)),
    ])


############################################################ DROPDOWN E PÁGINA GRÁFICO MULTITEK CA ################################

def gerar_layout_grafico_multitek_ca(elemento):
    return html.Div([
        html.H1(f'Dashboard de Carta Controle - {elemento}'),
        html.Br(),
        
        # Div com texto "Selecione o mês" e o dropdown para selecionar o mês
        html.Div([
            html.Span("Selecione o mês:", style={"margin-right": "10px", "font-size": "16px", "font-weight": "bold"}),
            dcc.Dropdown(
                id=f'dropdown-mes-multitek_ca-{elemento}',
                options=[
                    {"label": "Janeiro", "value": "Janeiro"},
                    {"label": "Fevereiro", "value": "Fevereiro"},
                    {"label": "Março", "value": "Março"},
                    {"label": "Abril", "value": "Abril"},
                    {"label": "Maio", "value": "Maio"},
                    {"label": "Junho", "value": "Junho"},
                    {"label": "Julho", "value": "Julho"},
                    {"label": "Agosto", "value": "Agosto"},
                    {"label": "Setembro", "value": "Setembro"},
                    {"label": "Outubro", "value": "Outubro"},
                    {"label": "Novembro", "value": "Novembro"},
                    {"label": "Dezembro", "value": "Dezembro"},
                ],
                placeholder="Escolha um mês",
                style={"width": "200px"}
            ),
        ], style={"display": "flex", "align-items": "center"}),

        html.Br(),
        dcc.Input(id=f'input-valor-multitek_ca-{elemento}', type='number', placeholder=f'mg/kg de {elemento}', debounce=True),
        html.Button('Adicionar Valor', id=f'btn-adicionar-multitek_ca-{elemento}', n_clicks=0),
        
        dcc.Graph(id=f'grafico-multitek_ca-{elemento}', 
                  figure=criar_carta_controle_multitek_ca(multitek_ca[elemento], [], elemento)),
       
    ])




############################################################### DROPDOWN E PÁGINA GRÁFICO PANANALYTICAL ############################

def gerar_layout_grafico_pananalytical(elemento):
    return html.Div([
        html.H1(f'Dashboard de Carta Controle - {elemento}'),
        html.Br(),
        
        # Div com texto "Selecione o mês" e o dropdown para selecionar o mês
        html.Div([
            html.Span("Selecione o mês:", style={"margin-right": "10px", "font-size": "16px", "font-weight": "bold"}),
            dcc.Dropdown(
                id=f'dropdown-mes-pananalytical-{elemento}',
                options=[
                    {"label": "Janeiro", "value": "Janeiro"},
                    {"label": "Fevereiro", "value": "Fevereiro"},
                    {"label": "Março", "value": "Março"},
                    {"label": "Abril", "value": "Abril"},
                    {"label": "Maio", "value": "Maio"},
                    {"label": "Junho", "value": "Junho"},
                    {"label": "Julho", "value": "Julho"},
                    {"label": "Agosto", "value": "Agosto"},
                    {"label": "Setembro", "value": "Setembro"},
                    {"label": "Outubro", "value": "Outubro"},
                    {"label": "Novembro", "value": "Novembro"},
                    {"label": "Dezembro", "value": "Dezembro"},
                ],
                placeholder="Escolha um mês",
                style={"width": "200px"}
            ),
        ], style={"display": "flex", "align-items": "center"}),

        html.Br(),
        dcc.Input(id=f'input-valor-pananalytical-{elemento}', type='number', placeholder=f'% de {elemento}', debounce=True),
        html.Button('Adicionar Valor', id=f'btn-adicionar-pananalytical-{elemento}', n_clicks=0),
        
        dcc.Graph(id=f'grafico-pananalytical-{elemento}', 
                  figure=criar_carta_controle_pananalytical(pananalytical[elemento], [], elemento)),
    ])


############################################################## DROPDOWN E PÁGINA GRÁFICO OXFORD ####################################

def gerar_layout_grafico_oxford(elemento):
    return html.Div([
        html.H1(f'Dashboard de Carta Controle - {elemento}'),
        html.Br(),
        
        # Div com texto "Selecione o mês" e o dropdown para selecionar o mês
        html.Div([
            html.Span("Selecione o mês:", style={"margin-right": "10px", "font-size": "16px", "font-weight": "bold"}),
            dcc.Dropdown(
                id=f'dropdown-mes-oxford-{elemento}',
                options=[
                    {"label": "Janeiro", "value": "Janeiro"},
                    {"label": "Fevereiro", "value": "Fevereiro"},
                    {"label": "Março", "value": "Março"},
                    {"label": "Abril", "value": "Abril"},
                    {"label": "Maio", "value": "Maio"},
                    {"label": "Junho", "value": "Junho"},
                    {"label": "Julho", "value": "Julho"},
                    {"label": "Agosto", "value": "Agosto"},
                    {"label": "Setembro", "value": "Setembro"},
                    {"label": "Outubro", "value": "Outubro"},
                    {"label": "Novembro", "value": "Novembro"},
                    {"label": "Dezembro", "value": "Dezembro"},
                ],
                placeholder="Escolha um mês",
                style={"width": "200px"}
            ),
        ], style={"display": "flex", "align-items": "center"}),

        html.Br(),
        dcc.Input(id=f'input-valor-oxford-{elemento}', type='number', placeholder=f'mg/kg de {elemento}', debounce=True),
        html.Button('Adicionar Valor', id=f'btn-adicionar-oxford-{elemento}', n_clicks=0),
        
        dcc.Graph(id=f'grafico-oxford-{elemento}', 
                  figure=criar_carta_controle_oxford(oxford_frx[elemento], [], elemento)),
    ])




############################################################### CALLBACKS ##########################################################
@app.callback(
    Output("submenu-flash2000", "is_open"),
    [Input("flash2000-button", "n_clicks")],
    [State("submenu-flash2000", "is_open")],
)
def toggle_submenu(n_clicks, is_open):
    if n_clicks:
        return not is_open
    return is_open

@app.callback(
    Output("submenu-multitek_cb", "is_open"),
    [Input("multitek_cb-button", "n_clicks")],
    [State("submenu-multitek_cb", "is_open")],
)
def toggle_submenu2(n_clicks, is_open):
    if n_clicks:
        return not is_open
    return is_open

@app.callback(
    Output("submenu-multitek_cm", "is_open"),
    [Input("multitek_cm-button", "n_clicks")],
    [State("submenu-multitek_cm", "is_open")],
)
def toggle_submenu3(n_clicks, is_open):
    if n_clicks:
        return not is_open
    return is_open

@app.callback(
    Output("submenu-flash2000_2", "is_open"),
    [Input("flash2000_2-button", "n_clicks")],
    [State("submenu-flash2000_2", "is_open")],
)
def toggle_submenu4(n_clicks, is_open):
    if n_clicks:
        return not is_open
    return is_open

@app.callback(
    Output("submenu-multitek_ca", "is_open"),
    [Input("multitek_ca-button", "n_clicks")],
    [State("submenu-multitek_ca", "is_open")],
)
def toggle_submenu6(n_clicks, is_open):
    if n_clicks:
        return not is_open
    return is_open



########################################## CALLBACKS PARA ATUALIZAÇÃO DO GRÁFICO ##############################################

########################################## FLASH 2000 #########################################################################

@app.callback(
    Output('input-valor-Carbono', 'value'),  # Reseta o valor do input após adicionar
    [Input('btn-adicionar-Carbono', 'n_clicks')],
    [State('input-valor-Carbono', 'value'),
     State('dropdown-mes-Carbono', 'value')]
)
def adicionar_valor_flash_carbono(n_clicks, novo_valor, selected_month):
    elemento = 'Carbono'  # Definindo o elemento

    # Verifica se o botão foi clicado e o valor é válido
    if n_clicks and novo_valor is not None and selected_month is not None:
        # Adicionar o novo valor ao mês selecionado
        nova_linha = pd.DataFrame({
            'Elemento': [elemento],
            'Mês': [selected_month],  # Associando ao mês correto
            'Data de Inserção': [datetime.now().strftime('%Y-%m-%d')],
            'Valor': [novo_valor]
        })
        
        # Adicionar a nova linha ao dataframe global
        global dados_flash
        dados_flash = pd.concat([dados_flash, nova_linha], ignore_index=True)

        # Salva as atualizações no CSV
        dados_flash.to_excel(CSV_PATH, index=False)

        print(f"Valor {novo_valor} adicionado para o mês {selected_month}")

    # Reseta o valor do input após adicionar
    return None

# Callback para atualizar o gráfico (quando o mês é alterado ou um novo valor é adicionado)
@app.callback(
    Output('grafico-controle-Carbono', 'figure'),
    [Input('dropdown-mes-Carbono', 'value'), 
     Input('btn-adicionar-Carbono', 'n_clicks')]  # Adiciona o botão de inserção como disparador
)
def update_grafico_flash_carbono(selected_month, n_clicks):
    elemento = 'Carbono'  # Definindo o elemento

    # Se nenhum mês foi selecionado, retorna um gráfico vazio
    if selected_month is None:
        return criar_carta_controle(dados_calibracao[elemento], [], elemento)

    # Filtrar os dados pelo mês selecionado
    dados_mes = dados_flash[(dados_flash['Elemento'] == elemento) & (dados_flash['Mês'] == selected_month)]

    # Retornar o gráfico com os dados do mês selecionado, reutilizando os limites e a média
    return criar_carta_controle_meses(dados_mes, elemento, selected_month)


@app.callback(
    Output('input-valor-Nitrogênio', 'value'),  # Reseta o valor do input após adicionar
    [Input('btn-adicionar-Nitrogênio', 'n_clicks')],
    [State('input-valor-Nitrogênio', 'value'),
     State('dropdown-mes-Nitrogênio', 'value')]
)
def adicionar_valor_flash_nitrogenio(n_clicks, novo_valor, selected_month):
    elemento = 'Nitrogênio'  # Definindo o elemento

    # Verifica se o botão foi clicado e o valor é válido
    if n_clicks and novo_valor is not None and selected_month is not None:
        # Adicionar o novo valor ao mês selecionado
        nova_linha = pd.DataFrame({
            'Elemento': [elemento],
            'Mês': [selected_month],  # Associando ao mês correto
            'Data de Inserção': [datetime.now().strftime('%Y-%m-%d')],
            'Valor': [novo_valor]
        })

        
        
        # Adicionar a nova linha ao dataframe global
        global dados_flash
        dados_flash = pd.concat([dados_flash, nova_linha], ignore_index=True)


        # Salva as atualizações no CSV
        dados_flash.to_excel(CSV_PATH, index=False)

        print(f"Valor {novo_valor} adicionado para o mês {selected_month}")

    # Reseta o valor do input após adicionar
    return None

# Callback para atualizar o gráfico (quando o mês é alterado ou um novo valor é adicionado)
@app.callback(
    Output('grafico-controle-Nitrogênio', 'figure'),
    [Input('dropdown-mes-Nitrogênio', 'value'), 
     Input('btn-adicionar-Nitrogênio', 'n_clicks')]  # Adiciona o botão de inserção como disparador
)
def update_grafico_flash_nitrogenio(selected_month, n_clicks):
    elemento = 'Nitrogênio'  # Definindo o elemento

    # Se nenhum mês foi selecionado, retorna um gráfico vazio
    if selected_month is None:
        return criar_carta_controle(dados_calibracao[elemento], [], elemento)

    # Filtrar os dados pelo mês selecionado
    dados_mes = dados_flash[(dados_flash['Elemento'] == elemento) & (dados_flash['Mês'] == selected_month)]

    # Retornar o gráfico com os dados do mês selecionado, reutilizando os limites e a média
    return criar_carta_controle_meses_flash_nitrogenio(dados_mes, elemento, selected_month)


@app.callback(
    Output('input-valor-Hidrogênio', 'value'),  # Reseta o valor do input após adicionar
    [Input('btn-adicionar-Hidrogênio', 'n_clicks')],
    [State('input-valor-Hidrogênio', 'value'),
     State('dropdown-mes-Hidrogênio', 'value')]
)
def adicionar_valor_flash_hidrogenio(n_clicks, novo_valor, selected_month):
    elemento = 'Hidrogênio'  # Definindo o elemento

    # Verifica se o botão foi clicado e o valor é válido
    if n_clicks and novo_valor is not None and selected_month is not None:
        # Adicionar o novo valor ao mês selecionado
        nova_linha = pd.DataFrame({
            'Elemento': [elemento],
            'Mês': [selected_month],  # Associando ao mês correto
            'Data de Inserção': [datetime.now().strftime('%Y-%m-%d')],
            'Valor': [novo_valor]
        })
        
        # Adicionar a nova linha ao dataframe global
        global dados_flash
        dados_flash = pd.concat([dados_flash, nova_linha], ignore_index=True)

        # Salva as atualizações no CSV
        dados_flash.to_excel(CSV_PATH, index=False)

        print(f"Valor {novo_valor} adicionado para o mês {selected_month}")

    # Reseta o valor do input após adicionar
    return None

# Callback para atualizar o gráfico (quando o mês é alterado ou um novo valor é adicionado)
@app.callback(
    Output('grafico-controle-Hidrogênio', 'figure'),
    [Input('dropdown-mes-Hidrogênio', 'value'), 
     Input('btn-adicionar-Hidrogênio', 'n_clicks')]  # Adiciona o botão de inserção como disparador
)
def update_grafico_flash_hidrogenio(selected_month, n_clicks):
    elemento = 'Hidrogênio'  # Definindo o elemento

    # Se nenhum mês foi selecionado, retorna um gráfico vazio
    if selected_month is None:
        return criar_carta_controle(dados_calibracao[elemento], [], elemento)

    # Filtrar os dados pelo mês selecionado
    dados_mes = dados_flash[(dados_flash['Elemento'] == elemento) & (dados_flash['Mês'] == selected_month)]

    # Retornar o gráfico com os dados do mês selecionado, reutilizando os limites e a média
    return criar_carta_controle_meses_flash_hidrogenio(dados_mes, elemento, selected_month)

@app.callback(
    Output('input-valor-Enxofre', 'value'),  # Reseta o valor do input após adicionar
    [Input('btn-adicionar-Enxofre', 'n_clicks')],
    [State('input-valor-Enxofre', 'value'),
     State('dropdown-mes-Enxofre', 'value')]
)
def adicionar_valor_flash_enxofre(n_clicks, novo_valor, selected_month):
    elemento = 'Enxofre'  # Definindo o elemento

    # Verifica se o botão foi clicado e o valor é válido
    if n_clicks and novo_valor is not None and selected_month is not None:
        # Adicionar o novo valor ao mês selecionado
        nova_linha = pd.DataFrame({
            'Elemento': [elemento],
            'Mês': [selected_month],  # Associando ao mês correto
            'Data de Inserção': [datetime.now().strftime('%Y-%m-%d')],
            'Valor': [novo_valor]
        })
        
        # Adicionar a nova linha ao dataframe global
        global dados_flash
        dados_flash = pd.concat([dados_flash, nova_linha], ignore_index=True)

        # Salva as atualizações no CSV
        dados_flash.to_excel(CSV_PATH, index=False)

        print(f"Valor {novo_valor} adicionado para o mês {selected_month}")

    # Reseta o valor do input após adicionar
    return None

# Callback para atualizar o gráfico (quando o mês é alterado ou um novo valor é adicionado)
@app.callback(
    Output('grafico-controle-Enxofre', 'figure'),
    [Input('dropdown-mes-Enxofre', 'value'), 
     Input('btn-adicionar-Enxofre', 'n_clicks')]  # Adiciona o botão de inserção como disparador
)
def update_grafico_flash_enxofre(selected_month, n_clicks):
    elemento = 'Enxofre'  # Definindo o elemento

    # Se nenhum mês foi selecionado, retorna um gráfico vazio
    if selected_month is None:
        return criar_carta_controle(dados_calibracao[elemento], [], elemento)

    # Filtrar os dados pelo mês selecionado
    dados_mes = dados_flash[(dados_flash['Elemento'] == elemento) & (dados_flash['Mês'] == selected_month)]

    # Retornar o gráfico com os dados do mês selecionado, reutilizando os limites e a média
    return criar_carta_controle_meses_flash_enxofre(dados_mes, elemento, selected_month)




################################################### MULTITEK CRUVA BAIXA #############################################################

@app.callback(
    Output('input-valor-multitek_cb-Enxofre', 'value'),  # Reseta o valor do input após adicionar
    [Input('btn-adicionar-multitek_cb-Enxofre', 'n_clicks')],
    [State('input-valor-multitek_cb-Enxofre', 'value'),
     State('dropdown-mes-multitek_cb-Enxofre', 'value')]
)
def adicionar_valor_multitek_cb_enxofre(n_clicks, novo_valor, selected_month):
    elemento = 'Enxofre'  # Definindo o elemento

    # Verifica se o botão foi clicado e o valor é válido
    if n_clicks and novo_valor is not None and selected_month is not None:

        m_inicial_s_cb = 27.5059
        m_final_s_cb = 47.7022
        d_tolueno = 0.8650
        fd_s_cb = m_final_s_cb / m_inicial_s_cb
        densidade_s_cb = 0.8463
        d_final_s_cb = (((fd_s_cb - 1) * d_tolueno) + densidade_s_cb) / fd_s_cb
        k_s_cb = d_final_s_cb / d_tolueno

        correcao_s_cb = novo_valor / k_s_cb
        correcao_s_cb = round(correcao_s_cb, 4)

        
        nova_linha = pd.DataFrame({
            'Elemento': [elemento],
            'Mês': [selected_month],  # Associando ao mês correto
            'Data de Inserção': [datetime.now().strftime('%Y-%m-%d')],
            'Valor': [correcao_s_cb]
        })
        
        # Adicionar a nova linha ao dataframe global
        global dados_multitek_cb
        dados_multitek_cb = pd.concat([dados_multitek_cb, nova_linha], ignore_index=True)

        # Salva as atualizações no CSV
        dados_multitek_cb.to_excel(caminho_multitek_cb, index=False)

        print(f"Valor {correcao_s_cb} adicionado para o mês {selected_month}")

    # Reseta o valor do input após adicionar
    return None

# Callback para atualizar o gráfico (quando o mês é alterado ou um novo valor é adicionado)
@app.callback(
    Output('grafico-multitek_cb-Enxofre', 'figure'),
    [Input('dropdown-mes-multitek_cb-Enxofre', 'value'), 
     Input('btn-adicionar-multitek_cb-Enxofre', 'n_clicks')]  
)
def update_grafico_multitek_cb_enxofre(selected_month, n_clicks):
    elemento = 'Enxofre'  # Definindo o elemento

    # Se nenhum mês foi selecionado, retorna um gráfico vazio
    if selected_month is None:
        return criar_carta_controle_multitek_cb(multitek_cb[elemento], [], elemento)

    # Filtrar os dados pelo mês selecionado
    dados_mes = dados_multitek_cb[(dados_multitek_cb['Elemento'] == elemento) & (dados_multitek_cb['Mês'] == selected_month)]

    # Retornar o gráfico com os dados do mês selecionado, reutilizando os limites e a média
    return criar_carta_controle_meses_multitek_cb_enxofre(dados_mes, elemento, selected_month)

@app.callback(
    Output('input-valor-multitek_cb-Nitrogênio', 'value'),  # Reseta o valor do input após adicionar
    [Input('btn-adicionar-multitek_cb-Nitrogênio', 'n_clicks')],
    [State('input-valor-multitek_cb-Nitrogênio', 'value'),
     State('dropdown-mes-multitek_cb-Nitrogênio', 'value')]
)
def adicionar_valor_multitek_cb_nitrogenio(n_clicks, novo_valor, selected_month):
    elemento = 'Nitrogênio'  # Definindo o elemento

    # Verifica se o botão foi clicado e o valor é válido
    if n_clicks and novo_valor is not None and selected_month is not None:

        m_inicial_n_cb = 27.5059
        m_final_n_cb = 47.7022
        d_tolueno = 0.8650
        fd_n_cb = m_final_n_cb / m_inicial_n_cb
        densidade_n_cb = 0.8463
        d_final_n_cb = (((fd_n_cb - 1) * d_tolueno) + densidade_n_cb) / fd_n_cb
        k_n_cb = d_final_n_cb / d_tolueno

        correcao_n_cb = novo_valor / k_n_cb
        correcao_n_cb = round(correcao_n_cb, 4)
        
        nova_linha = pd.DataFrame({
            'Elemento': [elemento],
            'Mês': [selected_month],  # Associando ao mês correto
            'Data de Inserção': [datetime.now().strftime('%Y-%m-%d')],
            'Valor': [correcao_n_cb]
        })
        
        # Adicionar a nova linha ao dataframe global
        global dados_multitek_cb
        dados_multitek_cb = pd.concat([dados_multitek_cb, nova_linha], ignore_index=True)

        # Salva as atualizações no CSV
        dados_multitek_cb.to_excel(caminho_multitek_cb, index=False)

        print(f"Valor {correcao_n_cb} adicionado para o mês {selected_month}")

    # Reseta o valor do input após adicionar
    return None

# Callback para atualizar o gráfico (quando o mês é alterado ou um novo valor é adicionado)
@app.callback(
    Output('grafico-multitek_cb-Nitrogênio', 'figure'),
    [Input('dropdown-mes-multitek_cb-Nitrogênio', 'value'), 
     Input('btn-adicionar-multitek_cb-Nitrogênio', 'n_clicks')]  # Adiciona o botão de inserção como disparador
)
def update_grafico_multitek_cb_nitrogenio(selected_month, n_clicks):
    elemento = 'Nitrogênio'  # Definindo o elemento

    # Se nenhum mês foi selecionado, retorna um gráfico vazio
    if selected_month is None:
        return criar_carta_controle_multitek_cb(multitek_cb[elemento], [], elemento)

    # Filtrar os dados pelo mês selecionado
    dados_mes = dados_multitek_cb[(dados_multitek_cb['Elemento'] == elemento) & (dados_multitek_cb['Mês'] == selected_month)]

    # Retornar o gráfico com os dados do mês selecionado, reutilizando os limites e a média
    return criar_carta_controle_meses_multitek_cb_nitrogenio(dados_mes, elemento, selected_month)



################################################### MULTITEK CRUVA MÉDIA #############################################################

@app.callback(
    Output('input-valor-multitek_cm-Enxofre', 'value'),  # Reseta o valor do input após adicionar
    [Input('btn-adicionar-multitek_cm-Enxofre', 'n_clicks')],
    [State('input-valor-multitek_cm-Enxofre', 'value'),
     State('dropdown-mes-multitek_cm-Enxofre', 'value')]
)
def adicionar_valor_multitek_cm_enxofre(n_clicks, novo_valor, selected_month):
    elemento = 'Enxofre'  # Definindo o elemento

    # Verifica se o botão foi clicado e o valor é válido
    if n_clicks and novo_valor is not None and selected_month is not None:
        
        mi_s_cm = 5.0609
        mf_s_cm = 50.5416
        d_tolueno = 0.8650
        fd_s_cm = mf_s_cm / mi_s_cm
        densidade_s_cm = 0.8492
        d_final_s_cm = (((fd_s_cm - 1) * d_tolueno) + densidade_s_cm) / fd_s_cm
        k_s_cm = d_final_s_cm / d_tolueno

        correcao_s_cm = novo_valor / k_s_cm
        correcao_s_cm = round(correcao_s_cm, 4)


        nova_linha = pd.DataFrame({
            'Elemento': [elemento],
            'Mês': [selected_month],  # Associando ao mês correto
            'Data de Inserção': [datetime.now().strftime('%Y-%m-%d')],
            'Valor': [correcao_s_cm]
        })
        
        # Adicionar a nova linha ao dataframe global
        global dados_multitek_cm
        dados_multitek_cm = pd.concat([dados_multitek_cm, nova_linha], ignore_index=True)

        # Salva as atualizações no CSV
        dados_multitek_cm.to_excel(caminho_multitek_cm, index=False)

        print(f"Valor {correcao_s_cm} adicionado para o mês {selected_month}")

    # Reseta o valor do input após adicionar
    return None

# Callback para atualizar o gráfico (quando o mês é alterado ou um novo valor é adicionado)
@app.callback(
    Output('grafico-multitek_cm-Enxofre', 'figure'),
    [Input('dropdown-mes-multitek_cm-Enxofre', 'value'), 
     Input('btn-adicionar-multitek_cm-Enxofre', 'n_clicks')]  # Adiciona o botão de inserção como disparador
)
def update_grafico_multitek_cm_enxofre(selected_month, n_clicks):
    elemento = 'Enxofre'  # Definindo o elemento

    # Se nenhum mês foi selecionado, retorna um gráfico vazio
    if selected_month is None:
        return criar_carta_controle_multitek_cm(multitek_cm[elemento], [], elemento)

    # Filtrar os dados pelo mês selecionado
    dados_mes = dados_multitek_cm[(dados_multitek_cm['Elemento'] == elemento) & (dados_multitek_cm['Mês'] == selected_month)]

    # Retornar o gráfico com os dados do mês selecionado, reutilizando os limites e a média
    return criar_carta_controle_meses_multitek_cm_enxofre(dados_mes, elemento, selected_month)

@app.callback(
    Output('input-valor-multitek_cm-Nitrogênio', 'value'),  # Reseta o valor do input após adicionar
    [Input('btn-adicionar-multitek_cm-Nitrogênio', 'n_clicks')],
    [State('input-valor-multitek_cm-Nitrogênio', 'value'),
     State('dropdown-mes-multitek_cm-Nitrogênio', 'value')]
)
def adicionar_valor_multitek_cm_nitrogenio(n_clicks, novo_valor, selected_month):
    elemento = 'Nitrogênio'  # Definindo o elemento

    
    if n_clicks and novo_valor is not None and selected_month is not None:
        
        mi_n_cm = 5.1286
        mf_n_cm = 50.1362
        d_tolueno = 0.8650
        fd_n_cm = mf_n_cm / mi_n_cm
        densidade_n_cm = 0.8492
        d_final_n_cm = (((fd_n_cm - 1) * d_tolueno) + densidade_n_cm) / fd_n_cm
        k_n_cm = d_final_n_cm / d_tolueno
        

        correcao_n_cm = novo_valor / k_n_cm
        correcao_n_cm = round(correcao_n_cm, 4)


        nova_linha = pd.DataFrame({
            'Elemento': [elemento],
            'Mês': [selected_month],  # Associando ao mês correto
            'Data de Inserção': [datetime.now().strftime('%Y-%m-%d')],
            'Valor': [correcao_n_cm]
        })
        
        # Adicionar a nova linha ao dataframe global
        global dados_multitek_cm
        dados_multitek_cm = pd.concat([dados_multitek_cm, nova_linha], ignore_index=True)

        # Salva as atualizações no CSV
        dados_multitek_cm.to_excel(caminho_multitek_cm, index=False)

        print(f"Valor {correcao_n_cm} adicionado para o mês {selected_month}")
        

    # Reseta o valor do input após adicionar
    return None

# Callback para atualizar o gráfico (quando o mês é alterado ou um novo valor é adicionado)
@app.callback(
    Output('grafico-multitek_cm-Nitrogênio', 'figure'),
    [Input('dropdown-mes-multitek_cm-Nitrogênio', 'value'), 
     Input('btn-adicionar-multitek_cm-Nitrogênio', 'n_clicks')]  # Adiciona o botão de inserção como disparador
)
def update_grafico_multitek_cm_nitrogenio(selected_month, n_clicks):
    elemento = 'Nitrogênio'  # Definindo o elemento

    # Se nenhum mês foi selecionado, retorna um gráfico vazio
    if selected_month is None:
        return criar_carta_controle_multitek_cm(multitek_cm[elemento], [], elemento)

    # Filtrar os dados pelo mês selecionado
    dados_mes = dados_multitek_cm[(dados_multitek_cm['Elemento'] == elemento) & (dados_multitek_cm['Mês'] == selected_month)]

    # Retornar o gráfico com os dados do mês selecionado, reutilizando os limites e a média
    return criar_carta_controle_meses_multitek_cm_nitrogenio(dados_mes, elemento, selected_month)



############################################################ MULTITEK CURVA ALTA #################################################

@app.callback(
    Output('input-valor-multitek_ca-Enxofre', 'value'),  # Reseta o valor do input após adicionar
    [Input('btn-adicionar-multitek_ca-Enxofre', 'n_clicks')],
    [State('input-valor-multitek_ca-Enxofre', 'value'),
     State('dropdown-mes-multitek_ca-Enxofre', 'value')]
)
def adicionar_valor_multitek_ca_enxofre(n_clicks, novo_valor, selected_month):
    elemento = 'Enxofre'  # Definindo o elemento

    
    if n_clicks and novo_valor is not None and selected_month is not None:

        mi_s_ca = 20.0118
        mf_s_ca = 40.0263
        d_tolueno = 0.8650
        fd_s_ca = mf_s_ca / mi_s_ca
        densidade_s_ca = 0.8492
        d_final_s_ca = (((fd_s_ca - 1) * d_tolueno) + densidade_s_ca) / fd_s_ca
        k_s_ca = d_final_s_ca / d_tolueno

        correcao_s_ca = novo_valor / k_s_ca
        correcao_s_ca = round(correcao_s_ca, 4)
        
        nova_linha = pd.DataFrame({
            'Elemento': [elemento],
            'Mês': [selected_month],  # Associando ao mês correto
            'Data de Inserção': [datetime.now().strftime('%Y-%m-%d')],
            'Valor': [correcao_s_ca]
        })
        
        # Adicionar a nova linha ao dataframe global
        global dados_multitek_ca
        dados_multitek_ca = pd.concat([dados_multitek_ca, nova_linha], ignore_index=True)

        # Salva as atualizações no CSV
        dados_multitek_ca.to_excel(caminho_multitek_ca, index=False)

        print(f"Valor {correcao_s_ca} adicionado para o mês {selected_month}")

    # Reseta o valor do input após adicionar
    return None

# Callback para atualizar o gráfico (quando o mês é alterado ou um novo valor é adicionado)
@app.callback(
    Output('grafico-multitek_ca-Enxofre', 'figure'),
    [Input('dropdown-mes-multitek_ca-Enxofre', 'value'), 
     Input('btn-adicionar-multitek_ca-Enxofre', 'n_clicks')]  # Adiciona o botão de inserção como disparador
)
def update_grafico_multitek_ca_enxofre(selected_month, n_clicks):
    elemento = 'Enxofre'  # Definindo o elemento

    # Se nenhum mês foi selecionado, retorna um gráfico vazio
    if selected_month is None:
        return criar_carta_controle_multitek_ca(multitek_ca[elemento], [], elemento)

    # Filtrar os dados pelo mês selecionado
    dados_mes = dados_multitek_ca[(dados_multitek_ca['Elemento'] == elemento) & (dados_multitek_ca['Mês'] == selected_month)]

    # Retornar o gráfico com os dados do mês selecionado, reutilizando os limites e a média
    return criar_carta_controle_meses_multitek_ca_enxofre(dados_mes, elemento, selected_month)

@app.callback(
    Output('input-valor-multitek_ca-Nitrogênio', 'value'),  # Reseta o valor do input após adicionar
    [Input('btn-adicionar-multitek_ca-Nitrogênio', 'n_clicks')],
    [State('input-valor-multitek_ca-Nitrogênio', 'value'),
     State('dropdown-mes-multitek_ca-Nitrogênio', 'value')]
)
def adicionar_valor_multitek_ca_nitrogenio(n_clicks, novo_valor, selected_month):
    elemento = 'Nitrogênio'  # Definindo o elemento

    
    if n_clicks and novo_valor is not None and selected_month is not None:

        mi_n_ca = 20.0118
        mf_n_ca = 40.0263
        d_tolueno = 0.8650
        fd_n_ca = mf_n_ca / mi_n_ca
        densidade_n_ca = 0.8492
        d_final_n_ca = (((fd_n_ca - 1) * d_tolueno) + densidade_n_ca) / fd_n_ca
        k_n_ca = d_final_n_ca / d_tolueno

        correcao_n_ca = novo_valor / k_n_ca
        correcao_n_ca = round(correcao_n_ca, 4)
        
        
        nova_linha = pd.DataFrame({
            'Elemento': [elemento],
            'Mês': [selected_month],  # Associando ao mês correto
            'Data de Inserção': [datetime.now().strftime('%Y-%m-%d')],
            'Valor': [correcao_n_ca]
        })
        
        # Adicionar a nova linha ao dataframe global
        global dados_multitek_ca
        dados_multitek_ca = pd.concat([dados_multitek_ca, nova_linha], ignore_index=True)

        # Salva as atualizações no CSV
        dados_multitek_ca.to_excel(caminho_multitek_ca, index=False)

        print(f"Valor {correcao_n_ca} adicionado para o mês {selected_month}")

    # Reseta o valor do input após adicionar
    return None

# Callback para atualizar o gráfico (quando o mês é alterado ou um novo valor é adicionado)
@app.callback(
    Output('grafico-multitek_ca-Nitrogênio', 'figure'),
    [Input('dropdown-mes-multitek_ca-Nitrogênio', 'value'), 
     Input('btn-adicionar-multitek_ca-Nitrogênio', 'n_clicks')]  # Adiciona o botão de inserção como disparador
)
def update_grafico_multitek_cm_nitrogenio(selected_month, n_clicks):
    elemento = 'Nitrogênio'  # Definindo o elemento

    # Se nenhum mês foi selecionado, retorna um gráfico vazio
    if selected_month is None:
        return criar_carta_controle_multitek_ca(multitek_ca[elemento], [], elemento)

    # Filtrar os dados pelo mês selecionado
    dados_mes = dados_multitek_ca[(dados_multitek_ca['Elemento'] == elemento) & (dados_multitek_ca['Mês'] == selected_month)]

    # Retornar o gráfico com os dados do mês selecionado, reutilizando os limites e a média
    return criar_carta_controle_meses_multitek_ca_nitrogenio(dados_mes, elemento, selected_month)



############################################################# PANANALYTICAL #######################################################

@app.callback(
    Output('input-valor-pananalytical-Enxofre', 'value'),  # Reseta o valor do input após adicionar
    [Input('btn-adicionar-pananalytical-Enxofre', 'n_clicks')],
    [State('input-valor-pananalytical-Enxofre', 'value'),
     State('dropdown-mes-pananalytical-Enxofre', 'value')]
)
def adicionar_valor_pananalytical_enxofre(n_clicks, novo_valor, selected_month):
    elemento = 'Enxofre'  # Definindo o elemento

    # Verifica se o botão foi clicado e o valor é válido
    if n_clicks and novo_valor is not None and selected_month is not None:
        # Adicionar o novo valor ao mês selecionado
        nova_linha = pd.DataFrame({
            'Elemento': [elemento],
            'Mês': [selected_month],  # Associando ao mês correto
            'Data de Inserção': [datetime.now().strftime('%Y-%m-%d')],
            'Valor': [novo_valor]
        })
        
        # Adicionar a nova linha ao dataframe global
        global dados_pananalytical
        dados_pananalytical = pd.concat([dados_pananalytical, nova_linha], ignore_index=True)

        # Salva as atualizações no CSV
        dados_pananalytical.to_excel(caminho_pananalytical, index=False)

        print(f"Valor {novo_valor} adicionado para o mês {selected_month}")

    # Reseta o valor do input após adicionar
    return None

# Callback para atualizar o gráfico (quando o mês é alterado ou um novo valor é adicionado)
@app.callback(
    Output('grafico-pananalytical-Enxofre', 'figure'),
    [Input('dropdown-mes-pananalytical-Enxofre', 'value'), 
     Input('btn-adicionar-pananalytical-Enxofre', 'n_clicks')]  # Adiciona o botão de inserção como disparador
)
def update_grafico_pananalytical_enxofre(selected_month, n_clicks):

    elemento = 'Enxofre'  # Definindo o elemento

    # Se nenhum mês foi selecionado, retorna um gráfico vazio
    if selected_month is None:
        return criar_carta_controle_pananalytical(pananalytical[elemento], [], elemento)

    # Filtrar os dados pelo mês selecionado
    dados_mes = dados_pananalytical[(dados_pananalytical['Elemento'] == elemento) & (dados_pananalytical['Mês'] == selected_month)]

    # Retornar o gráfico com os dados do mês selecionado, reutilizando os limites e a média
    return criar_carta_controle_meses_pananalytical(dados_mes, elemento, selected_month)




############################################################# OXFORD ################################################################
@app.callback(
    Output('input-valor-oxford-Arsênio', 'value'),  # Reseta o valor do input após adicionar
    [Input('btn-adicionar-oxford-Arsênio', 'n_clicks')],
    [State('input-valor-oxford-Arsênio', 'value'),
     State('dropdown-mes-oxford-Arsênio', 'value')]

)
def adicionar_valor_oxford_arsenio(n_clicks, novo_valor, selected_month):
    elemento = 'Arsênio'  # Definindo o elemento

    # Verifica se o botão foi clicado e o valor é válido
    if n_clicks and novo_valor is not None and selected_month is not None:
        # Adicionar o novo valor ao mês selecionado
        nova_linha = pd.DataFrame({
            'Elemento': [elemento],
            'Mês': [selected_month],  # Associando ao mês correto
            'Data de Inserção': [datetime.now().strftime('%Y-%m-%d')],
            'Valor': [novo_valor]
        })
        
        # Adicionar a nova linha ao dataframe global
        global dados_oxford_frx
        dados_oxford_frx = pd.concat([dados_oxford_frx, nova_linha], ignore_index=True)

        # Salva as atualizações no CSV
        dados_oxford_frx.to_excel(caminho_oxford_frx, index=False)

        print(f"Valor {novo_valor} adicionado para o mês {selected_month}")

    # Reseta o valor do input após adicionar
    return None

# Callback para atualizar o gráfico (quando o mês é alterado ou um novo valor é adicionado)
@app.callback(
    Output('grafico-oxford-Arsênio', 'figure'),
    [Input('dropdown-mes-oxford-Arsênio', 'value'), 
     Input('btn-adicionar-oxford-Arsênio', 'n_clicks')]  # Adiciona o botão de inserção como disparador
)
def update_grafico_oxford_arsenio(selected_month, n_clicks):

    elemento = 'Arsênio'  # Definindo o elemento

    # Se nenhum mês foi selecionado, retorna um gráfico vazio
    if selected_month is None:
        return criar_carta_controle_oxford(oxford_frx[elemento], [], elemento)

    # Filtrar os dados pelo mês selecionado
    dados_mes = dados_oxford_frx[(dados_oxford_frx['Elemento'] == elemento) & (dados_oxford_frx['Mês'] == selected_month)]

    # Retornar o gráfico com os dados do mês selecionado, reutilizando os limites e a média
    return criar_carta_controle_meses_oxford(dados_mes, elemento, selected_month)


# Callback para trocar os gráficos de acordo com a navegação
@app.callback(
    Output("page-content", "children"),
    [Input("url", "pathname")]
)

def render_page_content(pathname):
    if pathname == "/carbono":
        return gerar_layout_grafico_flash2000("Carbono")
    
    elif pathname == "/nitrogenio":
        return gerar_layout_grafico_flash2000("Nitrogênio")
    
    elif pathname == "/hidrogenio":
        return gerar_layout_grafico_flash2000("Hidrogênio")
    
    elif pathname == "/enxofre":
        return gerar_layout_grafico_flash2000("Enxofre")
    
    elif pathname == "/enxofre_multitek_cb":
        return gerar_layout_grafico_multitek_cb("Enxofre")
    
    elif pathname == "/nitrogenio_multitek_cb":
        return gerar_layout_grafico_multitek_cb("Nitrogênio")
    
    elif pathname == "/pananalytical_frx":
        return gerar_layout_grafico_pananalytical("Enxofre")
    
    elif pathname == "/nitrogenio_multitek_cb":
        return gerar_layout_grafico_multitek_cb("Nitrogênio")
    
    elif pathname == "/enxofre_multitek_cb":
        return gerar_layout_grafico_multitek_cb("Enxofre")
    
    elif pathname == "/enxofre_multitek_cm":
        return gerar_layout_grafico_multitek_cm("Enxofre")
    
    elif pathname == "/nitrogenio_multitek_cm":
        return gerar_layout_grafico_multitek_cm("Nitrogênio")

    elif pathname == "/nitrogenio_multitek_ca":
        return gerar_layout_grafico_multitek_ca("Nitrogênio")

    elif pathname == "/enxofre_multitek_ca":
        return gerar_layout_grafico_multitek_ca("Enxofre")
    
    elif pathname == "/oxford_frx":
        return gerar_layout_grafico_oxford("Arsênio")
    
    else:
        return gerar_pagina_inicial()
    
# URL para navegação
app.layout = html.Div([dcc.Location(id='url'), sidebar, content])

server = app.server
@server.route('/baixar-planilha')
def baixar_planilha():
    nome = request.args.get('nome')  # ex: controle_carbono.xlsx
    caminho = os.path.join('Planilhas', nome)

    # Verifica se o arquivo existe
    if not os.path.exists(caminho):
        return f"Arquivo {nome} não encontrado.", 404

    return send_file(caminho, as_attachment=True)



if __name__ == '__main__':
    app.run_server()




