import pandas as pd
import re
import openpyxl
import streamlit as st
from io import BytesIO
from datetime import datetime

st.set_page_config(page_title="JFCE",
                   page_icon="chart",
                   layout="wide",
                   initial_sidebar_state="auto",
                   menu_items=None)

dados_adm = pd.read_csv('SERVIDORES_ADM_15_07_2025.csv')
dados_jud = pd.read_csv('SERVIDORES_JUD_15_07_2025.csv')
dados_adm = dados_adm.drop(columns=['LOTA_COD_LOTACAO'])

dados = pd.concat([dados_adm, dados_jud], ignore_index=True)

painel = st.sidebar.radio("Selecione o painel", ["Mapa da Corregedoria", "Dados", "Análises"])

if painel == "Mapa da Corregedoria":
    st.title("Mapa da Corregedoria")
    st.write("Este é o mapa da corregedoria do Tribunal Regional Federal da 5ª Região.")
    tabela = dados.groupby(["CARGO", "LOTACAO_PAI"]).size().unstack(fill_value=0)
    grupos_subsecao = {
        'TOTAL_FORTALEZA': [
            '1ª VARA - FORTALEZA-CE','2ª VARA - FORTALEZA-CE','3ª VARA - FORTALEZA-CE','4ª VARA - FORTALEZA-CE','5ª VARA - FORTALEZA-CE',
            '6ª VARA - FORTALEZA-CE','7ª VARA - FORTALEZA-CE','8ª VARA - FORTALEZA-CE','9ª VARA - FORTALEZA-CE','10ª VARA - FORTALEZA-CE',
            '11ª VARA - FORTALEZA-CE','12ª VARA - FORTALEZA-CE','13ª VARA - JEF - FORTALEZA-CE','14ª VARA - JEF - FORTALEZA-CE',
            '20ª VARA - FORTALEZA-CE','21ª VARA - JEF - FORTALEZA-CE','26ª VARA - JEF - FORTALEZA - CE','28ª VARA - JEF - FORTALEZA-CE',
            '32ª VARA - FORTALEZA-CE','33ª VARA - FORTALEZA-CE',
            '1ª TURMA RECURSAL','2ª TURMA RECURSAL','2ª TURMA RECURSAL/JEF/CE','3ª TURMA RECURSAL','3ª TURMA RECURSAL/JEF/CE',
            'DIRETORIA DO FORO','SECRETARIA ADMINISTRATIVA'
        ],
        'TOTAL_LIMOEIRO': ['15ª VARA - LIMOEIRO DO NORTE-CE','29ª VARA - JEF - LIMOEIRO DO NORTE - CE'],
        'TOTAL_JUAZEIRO': ['16ª VARA - JUAZEIRO DO NORTE-CE','17ª VARA - JEF - JUAZEIRO DO NORTE-CE','30ª VARA - JEF - JUAZEIRO DO NORTE - CE'],
        'TOTAL_SOBRAL': ['18ª VARA - SOBRAL-CE','19ª VARA - JEF - SOBRAL-CE','31ª VARA - JEF - SOBRAL - CE'],
        'TOTAL_CRATEUS': ['22ª VARA - CRATEÚS-CE'],
        'TOTAL_QUIXADA': ['23ª VARA - QUIXADÁ-CE'],
        'TOTAL_TAUA': ['24ª VARA - TAUÁ-CE'],
        'TOTAL_IGUATU': ['25ª VARA - IGUATU-CE'],
        'TOTAL_ITAPIPOCA': ['27ª VARA- ITAPIPOCA-CE'],
        'TOTAL_MARACANAU': ['34ª VARA - MARACANAÚ-CE','35ª VARA - JEF - MARACANAÚ-CE'],
        'TOTAL_SERVIDORES': ['TOTAL_FORTALEZA','TOTAL_LIMOEIRO''TOTAL_JUAZEIRO','TOTAL_SOBRAL','TOTAL_CRATEUS','TOTAL_QUIXADA','TOTAL_TAUA',
                            'TOTAL_IGUATU','TOTAL_ITAPIPOCA', 'TOTAL_MARACANAU']
    }
    for nome_total, colunas in grupos_subsecao.items():
        colunas_presentes = [col for col in colunas if col in tabela.columns]
        tabela[nome_total] = tabela[colunas_presentes].sum(axis=1)
    ordem_colunas  = ['1ª VARA - FORTALEZA-CE','2ª VARA - FORTALEZA-CE','3ª VARA - FORTALEZA-CE','4ª VARA - FORTALEZA-CE','5ª VARA - FORTALEZA-CE',
            '6ª VARA - FORTALEZA-CE','7ª VARA - FORTALEZA-CE','8ª VARA - FORTALEZA-CE','9ª VARA - FORTALEZA-CE','10ª VARA - FORTALEZA-CE',
            '11ª VARA - FORTALEZA-CE','12ª VARA - FORTALEZA-CE','13ª VARA - JEF - FORTALEZA-CE','14ª VARA - JEF - FORTALEZA-CE',
            '20ª VARA - FORTALEZA-CE','21ª VARA - JEF - FORTALEZA-CE','26ª VARA - JEF - FORTALEZA - CE','28ª VARA - JEF - FORTALEZA-CE',
            '32ª VARA - FORTALEZA-CE','33ª VARA - FORTALEZA-CE',
            '1ª TURMA RECURSAL','2ª TURMA RECURSAL','2ª TURMA RECURSAL/JEF/CE','3ª TURMA RECURSAL','3ª TURMA RECURSAL/JEF/CE',#Recursais
            'DIRETORIA DO FORO','SECRETARIA ADMINISTRATIVA',
            'TOTAL_FORTALEZA',
            '15ª VARA - LIMOEIRO DO NORTE-CE','29ª VARA - JEF - LIMOEIRO DO NORTE - CE',#Faltou subdiretoria do foro
            'TOTAL_LIMOEIRO',
            '16ª VARA - JUAZEIRO DO NORTE-CE','17ª VARA - JEF - JUAZEIRO DO NORTE-CE','30ª VARA - JEF - JUAZEIRO DO NORTE - CE',#Faltou Sub Foro
            'TOTAL_JUAZEIRO',
            '18ª VARA - SOBRAL-CE','19ª VARA - JEF - SOBRAL-CE','31ª VARA - JEF - SOBRAL - CE',#Faltou subdiretoria do foro
            'TOTAL_SOBRAL',
            '22ª VARA - CRATEÚS-CE',#Faltou subdiretoria do foro
            'TOTAL_CRATEUS',
            '23ª VARA - QUIXADÁ-CE',#Faltou subdiretoria do foro
            'TOTAL_QUIXADA',                 
            '24ª VARA - TAUÁ-CE',#Faltou subdiretoria do foro
            'TOTAL_TAUA',                 
            '25ª VARA - IGUATU-CE',#Faltou subdiretoria do foro
            'TOTAL_IGUATU',
            '27ª VARA- ITAPIPOCA-CE',#Faltou subdiretoria do foro
            'TOTAL_ITAPIPOCA',
            '34ª VARA - MARACANAÚ-CE','35ª VARA - JEF - MARACANAÚ-CE',
            'TOTAL_MARACANAU',
            'TOTAL_SERVIDORES'
            #Extras
            'NUCLEO DE GESTAO DE PESSOAS','NUCLEO DE TECNOLOGIA DA INFORMAÇAO E COMUNICAÇAO','NUCLEO DE INFRAESTRUTURA E ADMINISTRAÇAO PREDIAL',
            'NUCLEO DE ESTRATEGIA, GOVERNANÇA E INTEGRIDADE','NUCLEO DE AUDITORIA INTERNA']
    colunas_presentes_ordenadas = [col for col in ordem_colunas if col in tabela.columns]
    tabela_ordenada = tabela[colunas_presentes_ordenadas]
    st.dataframe(tabela_ordenada)
    hoje = datetime.today().strftime('%d-%m-%Y')
    nome_arquivo = f'Mapa da Corregedoria {hoje}.xlsx'
    tabela_ordenada.to_excel(nome_arquivo, index=True, engine='openpyxl')
    with open(nome_arquivo, "rb") as file:
        st.download_button(label="📥 Baixar Mapa da Corregedoria", data=file,
            file_name=nome_arquivo, mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        st.markdown("""
        <style>
        .centralizar-container {
            max-width: 850px;
            margin: 0 auto;
            padding-top: 1rem;
        }
        .centralizar-container .stSelectbox {
            width: 100% !important;
        }
        .tabela-subsecoes {
            width: 100%;
            border-collapse: collapse;
            font-family: Arial, sans-serif;
            margin-top: 20px;
        }
        .tabela-subsecoes th, .tabela-subsecoes td {
            border: 1px solid #ddd;
            padding: 8px;
            text-align: left;
            font-size: 14px;
        }
        .tabela-subsecoes th {
            background-color: #f2f2f2;
        }
        .tabela-subsecoes td:first-child {
            min-width: 350px;
        }
        </style>
        <div class="centralizar-container">
    """, unsafe_allow_html=True)

    st.title("Subseções")

    # Lista ordenada
    subsecoes_ordenadas = [
        'SECRETARIA ADMINISTRATIVA','DIRETORIA DO FORO',
        '1ª TURMA RECURSAL','2ª TURMA RECURSAL','3ª TURMA RECURSAL','2ª TURMA RECURSAL/JEF/CE','3ª TURMA RECURSAL/JEF/CE',
        '1ª VARA - FORTALEZA-CE','2ª VARA - FORTALEZA-CE','3ª VARA - FORTALEZA-CE',
        '4ª VARA - FORTALEZA-CE', '5ª VARA - FORTALEZA-CE', '6ª VARA - FORTALEZA-CE',
        '7ª VARA - FORTALEZA-CE', '8ª VARA - FORTALEZA-CE', '9ª VARA - FORTALEZA-CE',
        '10ª VARA - FORTALEZA-CE', '11ª VARA - FORTALEZA-CE', '12ª VARA - FORTALEZA-CE',
        '13ª VARA - JEF - FORTALEZA-CE', '14ª VARA - JEF - FORTALEZA-CE',
        '15ª VARA - LIMOEIRO DO NORTE-CE', '16ª VARA - JUAZEIRO DO NORTE-CE',
        '17ª VARA - JEF - JUAZEIRO DO NORTE-CE', '18ª VARA - SOBRAL-CE',
        '19ª VARA - JEF - SOBRAL-CE', '20ª VARA - FORTALEZA-CE',
        '21ª VARA - JEF - FORTALEZA-CE', '22ª VARA - CRATEÚS-CE',
        '23ª VARA - QUIXADÁ-CE', '24ª VARA - TAUÁ-CE', '25ª VARA - IGUATU-CE',
        '26ª VARA - JEF - FORTALEZA - CE', '27ª VARA- ITAPIPOCA-CE',
        '28ª VARA - JEF - FORTALEZA-CE', '29ª VARA - JEF - LIMOEIRO DO NORTE - CE',
        '30ª VARA - JEF - JUAZEIRO DO NORTE - CE', '31ª VARA - JEF - SOBRAL - CE',
        '32ª VARA - FORTALEZA-CE', '33ª VARA - FORTALEZA-CE',
        '34ª VARA - MARACANAÚ-CE', '35ª VARA - JEF - MARACANAÚ-CE',
        'NUCLEO DE AUDITORIA INTERNA',
        'NUCLEO DE ESTRATEGIA, GOVERNANÇA E INTEGRIDADE',
        'NUCLEO DE GESTAO DE PESSOAS',
        'NUCLEO DE INFRAESTRUTURA E ADMINISTRAÇAO PREDIAL',
        'NUCLEO DE TECNOLOGIA DA INFORMAÇAO E COMUNICAÇAO'
    ]

    subsecoes_validas = [s for s in subsecoes_ordenadas if s in dados['LOTACAO_PAI'].unique()]
    selecao_subsecoes = st.selectbox("Selecione a subseção para ver a quantidade de servidores:", subsecoes_validas)

    tabela_subsecoes = dados[dados['LOTACAO_PAI'] == selecao_subsecoes]
    tabela_subsecoes = (
        tabela_subsecoes.groupby("CARGO")
        .size()
        .reset_index(name='Quantidade de servidores')
        .sort_values(by='Quantidade de servidores', ascending=False)
    )

    tabela_html = tabela_subsecoes.to_html(index=False, classes='tabela-subsecoes', border=0)
    st.markdown(tabela_html, unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)

    hoje = datetime.today().strftime('%d-%m-%Y')
    nome_arquivo = f"Subseção - {selecao_subsecoes} - {hoje}.xlsx"

    # Salva arquivo Excel
    tabela_subsecoes.to_excel(nome_arquivo, index=False, engine='openpyxl')

    # Cria botão de download
    with open(nome_arquivo, "rb") as file:
        st.download_button(
            label=f"📥 Baixar dados da {selecao_subsecoes}",
            data=file,
            file_name=nome_arquivo,
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
elif painel == "Dados":
    st.title("Dados primários")
    ordem_subsecoes = [
        'SECRETARIA ADMINISTRATIVA','DIRETORIA DO FORO',
        '1ª TURMA RECURSAL','2ª TURMA RECURSAL','3ª TURMA RECURSAL','2ª TURMA RECURSAL/JEF/CE','3ª TURMA RECURSAL/JEF/CE',
        '1ª VARA - FORTALEZA-CE','2ª VARA - FORTALEZA-CE','3ª VARA - FORTALEZA-CE',
        '4ª VARA - FORTALEZA-CE', '5ª VARA - FORTALEZA-CE', '6ª VARA - FORTALEZA-CE',
        '7ª VARA - FORTALEZA-CE', '8ª VARA - FORTALEZA-CE', '9ª VARA - FORTALEZA-CE',
        '10ª VARA - FORTALEZA-CE', '11ª VARA - FORTALEZA-CE', '12ª VARA - FORTALEZA-CE',
        '13ª VARA - JEF - FORTALEZA-CE', '14ª VARA - JEF - FORTALEZA-CE',
        '15ª VARA - LIMOEIRO DO NORTE-CE', '16ª VARA - JUAZEIRO DO NORTE-CE',
        '17ª VARA - JEF - JUAZEIRO DO NORTE-CE', '18ª VARA - SOBRAL-CE',
        '19ª VARA - JEF - SOBRAL-CE', '20ª VARA - FORTALEZA-CE',
        '21ª VARA - JEF - FORTALEZA-CE', '22ª VARA - CRATEÚS-CE',
        '23ª VARA - QUIXADÁ-CE', '24ª VARA - TAUÁ-CE', '25ª VARA - IGUATU-CE',
        '26ª VARA - JEF - FORTALEZA - CE', '27ª VARA- ITAPIPOCA-CE',
        '28ª VARA - JEF - FORTALEZA-CE', '29ª VARA - JEF - LIMOEIRO DO NORTE - CE',
        '30ª VARA - JEF - JUAZEIRO DO NORTE - CE', '31ª VARA - JEF - SOBRAL - CE',
        '32ª VARA - FORTALEZA-CE', '33ª VARA - FORTALEZA-CE',
        '34ª VARA - MARACANAÚ-CE', '35ª VARA - JEF - MARACANAÚ-CE',
        'NUCLEO DE AUDITORIA INTERNA',
        'NUCLEO DE ESTRATEGIA, GOVERNANÇA E INTEGRIDADE',
        'NUCLEO DE GESTAO DE PESSOAS',
        'NUCLEO DE INFRAESTRUTURA E ADMINISTRAÇAO PREDIAL',
        'NUCLEO DE TECNOLOGIA DA INFORMAÇAO E COMUNICAÇAO'
    ]
    subsecoes = [s for s in ordem_subsecoes if s in dados['LOTACAO_PAI'].unique()]
    selecao_subsecoes = st.selectbox("Selecione a subseção para ver os dados:", subsecoes)
    tabela = dados[dados['LOTACAO_PAI'] == selecao_subsecoes]
    tabela_formatada = tabela.reset_index(drop=True)
    st.data_editor(tabela_formatada, use_container_width=True, hide_index=True, disabled=True)
    hoje = datetime.today().strftime('%d-%m-%Y')
    nome_arquivo = f'Dados da subseção - {selecao_subsecoes} {hoje}.xlsx'
    tabela.to_excel(nome_arquivo, index=False, engine='openpyxl')
    with open(nome_arquivo, "rb") as file:
        st.download_button(label=f"📥 Baixar dados da {selecao_subsecoes}", data=file,
            file_name=nome_arquivo, mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
elif painel == "Análises":
    st.title("Análises")

    import plotly.express as px
    import matplotlib.pyplot as plt

    subsecoes = dados['LOTACAO_PAI'].value_counts().sort_values(ascending=False)
    cargos = dados['CARGO'].value_counts().sort_values(ascending=False)

    # Preparar DataFrames para os gráficos
    subsecoes_df = subsecoes.reset_index()
    subsecoes_df.columns = ['Subseção', 'Quantidade de Servidores']
    cargos_df = cargos.reset_index()
    cargos_df.columns = ['Cargo', 'Quantidade de Servidores']

    col1, col2 = st.columns([2,1])

    with col1:
        fig1 = px.bar(subsecoes_df, y='Subseção', x='Quantidade de Servidores', orientation='h',
            labels={'Subseção': 'Subseção', 'Quantidade de Servidores': 'Quantidade de Servidores'},
            title="Quantidade de Servidores por Subseção")
        fig1.update_layout(
            yaxis={'categoryorder':'total ascending', 'tickfont': dict(size=16), 'titlefont': dict(size=18)},
            xaxis={'tickfont': dict(size=16), 'titlefont': dict(size=18)},
            title={'font': dict(size=22)}, height=700)
        for i, row in subsecoes_df.iterrows():
            fig1.add_annotation(
            x=row['Quantidade de Servidores'],
            y=row['Subseção'],
            text=str(row['Quantidade de Servidores']), showarrow=False, font=dict(size=14), 
            xanchor='left', yanchor='middle')
        fig1.update_traces(text=None, textposition='outside')
        st.plotly_chart(fig1, use_container_width=True)

        fig2 = px.bar(cargos_df, y='Cargo', x='Quantidade de Servidores', orientation='h',
            labels={'Cargo': 'Cargo', 'Quantidade de Servidores': 'Qtd de Servidores'},
            title="Quantidade de Servidores por Cargo", color='Quantidade de Servidores',
            color_continuous_scale='Blues')
        fig2.update_layout(
            yaxis={'categoryorder':'total ascending', 'tickfont': dict(size=16), 'titlefont': dict(size=18)},
            xaxis={'tickfont': dict(size=16), 'titlefont': dict(size=18)}, title={'font': dict(size=22)}, height=700,
            coloraxis_colorbar=dict(orientation='h', y=-0.25, x=0.5, xanchor='center', len=0.7, thickness=15, 
                                    title=None))
        fig2.update_traces(text=cargos_df['Quantidade de Servidores'], textposition='outside')
        st.plotly_chart(fig2, use_container_width=True)
        
        with col2:
            provimento_counts = dados['PROVIMENTO'].value_counts().reset_index()
            provimento_counts.columns = ['Provimento', 'Quantidade']
            fig3 = px.pie(provimento_counts, names='Provimento', values='Quantidade',
                title='Distribuição dos Servidores por Tipo de Provimento', hole=0.4)
            fig3.update_layout(legend_title_text='Tipo de Provimento',
                               legend=dict(orientation="h", y=-0.2, x=0.5, xanchor="center"))
            st.plotly_chart(fig3, use_container_width=True)
            
            provimento_treemap = dados['PROVIMENTO'].value_counts().reset_index()
            provimento_treemap.columns = ['Provimento', 'Quantidade']
            fig4 = px.treemap(
                provimento_treemap,
                path=['Provimento'],
                values='Quantidade',
                color='Quantidade',
                color_continuous_scale='Blues',
                hover_data={'Quantidade': True},
                title='Treemap: Distribuição por Provimento'
            )
            fig4.update_layout(margin=dict(t=50, l=25, r=25, b=25), height=700)
            st.plotly_chart(fig4, use_container_width=True)
