import pandas as pd
import re
import openpyxl
dados = pd.read_excel('dados_exportados.xlsx')
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

def limpar_nome_aba(nome):
    return re.sub(r'[\[\]\:\*\?\/\\]', '', nome)[:31]

ordem_abas = ['SECRETARIA ADMINISTRATIVA','DIRETORIA DO FORO',
    '1ª TURMA RECURSAL','2ª TURMA RECURSAL','3ª TURMA RECURSAL','2ª TURMA RECURSAL/JEF/CE','3ª TURMA RECURSAL/JEF/CE',
    '1ª VARA - FORTALEZA-CE','2ª VARA - FORTALEZA-CE','3ª VARA - FORTALEZA-CE' 
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
    'NUCLEO DE TECNOLOGIA DA INFORMAÇAO E COMUNICAÇAO',
]

arquivo_saida = "tabela_contingencia_ordenada.xlsx"

with pd.ExcelWriter(arquivo_saida, engine='openpyxl') as writer:
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
            