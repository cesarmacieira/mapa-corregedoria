import pandas as pd
import re
import openpyxl
import streamlit as st
from io import BytesIO
from datetime import datetime
import plotly.express as px
import matplotlib.pyplot as plt

st.set_page_config(page_title="JFCE",
                   page_icon="chart",
                   layout="wide",
                   initial_sidebar_state="auto",
                   menu_items=None)

dados_adm = pd.read_csv('SERVIDORES_ADM_15_07_2025.csv')
dados_jud = pd.read_csv('SERVIDORES_JUD_15_07_2025.csv')
dados_adm = dados_adm.drop(columns=['LOTA_COD_LOTACAO'])

def classificar_vinculo(situacao):
    if pd.isna(situacao):
        return 'SEM VÍNCULO'
    elif situacao in ['ATIVO']:
        return 'EFETIVO'
    elif situacao in ['ATIVO - EX. PROVISÓRIO']:
        return 'EXERCÍCIO PROVISÓRIO'
    elif situacao in ['REQUISITADO DE MUNICIPIOS - ESTATUTARIO', 'REQUISITADO DA UNIAO - CLT',
       'REQUISITADO DE MUNICIPIOS - CLT', 'REQUISITADO DE ESTADOS - ESTATUTARIO',
       'REQUISITADO DO JUDICIARIO FEDERAL', 'REQUISITADO DA UNIAO - ESTATUTARIO','REQUISITADO DE ESTADOS - CLT']:
        return 'REQUISITADO'
    elif situacao in ['ATIVO REMOVIDO (ACOMPANHAMENTO DE CONJUGE)', 'ATIVO REMOVIDO (ART. 41 RES. CJF Nº 03/2008)',
       'ATIVO REMOVIDO (MOTIVO DE SAUDE)', 'ATIVO REMOVIDO (SINAR)', 'ATIVO REMOVIDO (A PEDIDO, CRITERIO DA ADMINISTRACAO)',
       'DO JUDICIARIO FEDERAL - ATIVO REMOVIDO SINAR', 'ATIVO REMOVIDO (POR PERMUTA - RES. TRF5 Nº 07/2015)']:
        return 'REMOVIDO'
    else:
        return 'OUTRO'

def limpar_nome_aba(nome):
    return re.sub(r'[\[\]\:\*\?\/\\]', '', nome)[:31]
    
dados = pd.concat([dados_adm, dados_jud], ignore_index=True)
dados['STATUS_PROVIMENTO'] = dados['SITUACAO'].apply(classificar_vinculo)

painel = st.sidebar.radio("Selecione o painel", ["Mapa da Corregedoria", "Dados Brutos", "Análises"])

if painel == "Mapa da Corregedoria":
    st.title("Mapa da Corregedoria")
    st.write("Este é o mapa da corregedoria da Justiça Federal do Ceará.")

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
    
    ordem_abas = [
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
    'NUCLEO DE TECNOLOGIA DA INFORMAÇAO E COMUNICAÇAO']

    ct_cargo = pd.crosstab(dados['CARGO'], dados['LOTACAO_PAI'])
    ct_prov  = pd.crosstab(dados['STATUS_PROVIMENTO'], dados['LOTACAO_PAI'])

    colunas_validas = [col for col in ordem_colunas if col in ct_cargo.columns or col in ct_prov.columns]

    ct_cargo = ct_cargo.reindex(columns=colunas_validas, fill_value=0)
    ct_prov  = ct_prov.reindex(columns=colunas_validas, fill_value=0)

    st.write("Cargos por serventia")
    st.dataframe(ct_cargo)

    st.write("Provimentos por serventia")
    st.dataframe(ct_prov)

    # Exportar tabela completa
    ct_prov = ct_prov.reindex(['EFETIVO', 'REQUISITADO', 'EXERCÍCIO PROVISÓRIO', 'REMOVIDO', 'OUTRO'])
    cols = ct_cargo.columns.union(ct_prov.columns)
    ct_cargo = ct_cargo.reindex(columns=cols, fill_value=0)
    ct_prov  = ct_prov.reindex(columns=cols, fill_value=0)
    cargo_block = ct_cargo.copy()
    cargo_block.index = pd.MultiIndex.from_product([['Cargo'], cargo_block.index],
        names=['Variável', ct_cargo.index.name or 'Categoria'])
    prov_block = ct_prov.copy()
    prov_block.index = pd.MultiIndex.from_product([['Provimento'], prov_block.index], names=['Variável', ct_prov.index.name or 'Categoria'])
    tabela = pd.concat([cargo_block, prov_block])
    tabela_com_totais = tabela.copy()
    for nome_total, lotacoes in grupos_subsecao.items():
        colunas = []
        for lot in lotacoes:
            if lot in grupos_subsecao:
                colunas.extend(grupos_subsecao[lot])
            else:
                colunas.append(lot)
        colunas_validas = [c for c in colunas if c in tabela_com_totais.columns]
        tabela_com_totais[nome_total] = tabela_com_totais[colunas_validas].sum(axis=1)

    colunas_validas = [col for col in ordem_colunas if col in tabela_com_totais.columns]
    colunas_restantes = [col for col in tabela_com_totais.columns if col not in colunas_validas]
    tabela_ordenada = tabela_com_totais[colunas_validas + colunas_restantes]

    hoje = datetime.today().strftime('%d-%m-%Y')
    nome_arquivo = f'Mapa da Corregedoria {hoje}.xlsx'
    with pd.ExcelWriter(nome_arquivo, engine='openpyxl') as writer:
        tabela_ordenada.to_excel(writer, sheet_name="Mapa da Corregedoria")
        abas_usadas = set()
        for lotacao in ordem_abas:
            if lotacao in dados['LOTACAO_PAI'].unique():
                df_lotacao = dados[dados['LOTACAO_PAI'] == lotacao]
                aba = limpar_nome_aba(lotacao)
                while aba in abas_usadas:
                    aba += "_"
                    aba = aba[:31]
                abas_usadas.add(aba)
                df_lotacao.to_excel(writer, sheet_name=aba, index=False)

    with open(nome_arquivo, "rb") as file:
        st.download_button(
            label="📥 Baixar Mapa da Corregedoria",
            data=file,
            file_name=nome_arquivo,
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        
    st.title("Lotações")
    lotacoes_validas = [s for s in ordem_abas if s in dados['LOTACAO_PAI'].unique()]
    selecao_lotacoes = st.selectbox("Selecione a lotação para ver a quantidade de servidores:", lotacoes_validas)

    tabela_lotacoes = dados[dados['LOTACAO_PAI'] == selecao_lotacoes]
    tabela_lotacoes = (tabela_lotacoes.groupby("CARGO").size().reset_index(name='Quantidade de servidores')
        .sort_values(by='Quantidade de servidores', ascending=False))
    st.dataframe(tabela_lotacoes.reset_index(drop=True), use_container_width=False)
    
    hoje = datetime.today().strftime('%d-%m-%Y')
    nome_arquivo = f"Lotação - {selecao_lotacoes} - {hoje}.xlsx"
    tabela_lotacoes.to_excel(nome_arquivo, index=False, engine='openpyxl')

    # Cria botão de download
    with open(nome_arquivo, "rb") as file:
        st.download_button(
            label=f"📥 Baixar dados da {selecao_lotacoes}",
            data=file,
            file_name=nome_arquivo,
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            key="download_lotacoes"
        )
        
    st.title("Provimento")
    selecao_provimento = st.selectbox("Selecione a lotação para ver a quantidade de servidores:", 
                                      lotacoes_validas,key="select_lotacao_provimento")

    tabela_provimento = dados[dados['LOTACAO_PAI'] == selecao_provimento]
    tabela_provimento = (tabela_provimento.groupby("STATUS_PROVIMENTO").size().reset_index(name='Quantidade de servidores')
        .sort_values(by='Quantidade de servidores', ascending=False))
    st.dataframe(tabela_provimento.reset_index(drop=True), use_container_width=False)
    hoje = datetime.today().strftime('%d-%m-%Y')
    nome_arquivo = f"Lotação - {selecao_provimento} - {hoje}.xlsx"
    tabela_provimento.to_excel(nome_arquivo, index=False, engine='openpyxl')

    # Cria botão de download
    with open(nome_arquivo, "rb") as file:
        st.download_button(
            label=f"📥 Baixar dados da {selecao_provimento}",
            data=file,
            file_name=nome_arquivo,
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            key="download_provimento"
        )

elif painel == "Dados Brutos":
    st.title("Dados primários")
    ordem_lotacoes = [
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
    lotacoes = [s for s in ordem_lotacoes if s in dados['LOTACAO_PAI'].unique()]
    selecao_lotacoes = st.selectbox("Selecione a lotação para ver os dados:", lotacoes)
    tabela = dados[dados['LOTACAO_PAI'] == selecao_lotacoes]
    tabela_formatada = tabela.reset_index(drop=True)
    st.data_editor(tabela_formatada, use_container_width=True, hide_index=True, disabled=True)
    hoje = datetime.today().strftime('%d-%m-%Y')
    nome_arquivo = f'Dados da lotação - {selecao_lotacoes} {hoje}.xlsx'
    tabela.to_excel(nome_arquivo, index=False, engine='openpyxl')
    with open(nome_arquivo, "rb") as file:
        st.download_button(label=f"📥 Baixar dados da {selecao_lotacoes}", data=file,
            file_name=nome_arquivo, mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
elif painel == "Análises":
    st.title("Análises")
    lotacoes = dados['LOTACAO_PAI'].value_counts().sort_values(ascending=False)
    cargos = dados['CARGO'].value_counts().sort_values(ascending=False)
    lotacoes_df = lotacoes.reset_index()
    lotacoes_df.columns = ['Lotação', 'Quantidade de Servidores']
    cargos_df = cargos.reset_index()
    cargos_df.columns = ['Cargo', 'Quantidade de Servidores']

    col1, col2 = st.columns([2,1])
    with col1:
        fig1 = px.bar(lotacoes_df, y='Lotação', x='Quantidade de Servidores', orientation='h',
            labels={'Lotação': 'Lotação', 'Quantidade de Servidores': 'Quantidade de Servidores'},
            title="Quantidade de Servidores por Lotação")
        fig1.update_layout(
            yaxis={'categoryorder':'total ascending', 'tickfont': dict(size=16), 'titlefont': dict(size=18)},
            xaxis={'tickfont': dict(size=16), 'titlefont': dict(size=18)},
            title={'font': dict(size=22)}, height=700)
        for i, row in lotacoes_df.iterrows():
            fig1.add_annotation(
            x=row['Quantidade de Servidores'],
            y=row['Lotação'],
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
