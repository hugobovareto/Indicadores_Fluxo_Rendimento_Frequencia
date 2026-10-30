'''
3 fontes de dados:
- Relatório Geral de Matrículas;
- Relatório de Frequência dos Estudantes;
- Relatório de Notas dos Estudantes.

Excluir SITUAÇÃO = "TRANSFERIDO" e "FALECIDO" do relatório de matrículas.

Considerar somente as Séries dos Anos Finais do EF e Ensino Médio:
Limpeza nos 3 relatórios: Geral de matrículas, Notas e Frequência.
- 1ª Série;
- 2ª Série;
- 3ª Série;
- 6º Ano;
- 7º Ano;
- 8º Ano;
- 9º Ano.


O identificador único entre as bases é a Matrícula:
- Relatório Geral de Matrículas: 'MATRÍCULA';
- Notas: 'MATRÍCULA ESTUDANTE';
- Frequência: 'MATRÍCULA'.

Para notas, considerar somente as notas do 1º, 2º e 3º bimestres.
Caso o estudante não tenha nota em algum bimestre, considerar como NaN, logo não impacta na média.

Considerar somente os componentes da BNCC para o df_notas:
bncc = ['Arte',
        'Biologia',
        'Educação Física',
        'Filosofia',
        'Física',
        'Geografia',
        'História',
        'Língua Inglesa',
        'Língua Portuguesa',
        'Matemática',
        'Química',
        'Sociologia', 
        'Ciências']

Classificação de Rendimento e Frequência dos Estudantes para os grupos:
- Geral;
- Ensino Fundamental - Anos Finais (informação através da série);
- Ensino Médio (informação através da série);
- EPT (informação através do componente curricular);
- Ensino Médio Noturno (informação através do código da turma).


O Relatório de Notas está segmentado por componentes, então nesse, cria a variável:
- nota média.


A partir da base por componente, criar a base por estudante com as variáveis:
- idade;
- frequência;
- classificação frequência;
- classificação rendimento;

Para EPT, eu trato com os componentes da BNCC mesmo, o ponto é que é um grupo específico de estudantes que possuem componente de EPT.
Então lido com esse grupo a partir do grupo geral de componentes da BNCC e depois cruzo com a informação no relatório de notas para as matrículas que possuem algum componente de EPT.

Para noturnos, no Relatório Geral de Matrículas, são as turmas:
- EMPN1A;
- EMPN1B;
- EMPN1C;
- EMPN1D;
- EMPN2A;
- EMPN2B;
- EMPN2C;
- EMPN2D;
- EMPN3A;
- EMPN3B;
- EMPN3C;
- EMPN3D.

'''

# Importação das bibliotecas
import pandas as pd
import glob
import os
from tqdm import tqdm  # Para barra de progresso
import numpy as np
import warnings
warnings.filterwarnings('ignore')
from datetime import datetime
from dateutil.relativedelta import relativedelta


# Carregar as 3 bases de dados:
# RELATÓRIO GERAL DE MATRÍCULAS
# Ler os dados do relatório geral de matrículas do SigEduc:
df_matriculas = pd.read_excel(r"C:\Users\hugob\Downloads\20261029_Relatório Geral de Estudantes - Matrículas.xlsx", skiprows=2)


# RELATÓRIO DE NOTAS DOS ESTUDANTES
# caminho da pasta onde estão os arquivos
pasta = r"C:\Users\hugob\Downloads\Notas"

# lista todos os arquivos .xlsx da pasta
arquivos = glob.glob(os.path.join(pasta, "*.xlsx"))

# lista para armazenar os dataframes
dfs = []

for arquivo in tqdm(arquivos, desc="Processando arquivos"):
    # lê cada arquivo, pulando as 2 primeiras linhas
    df_unico = pd.read_excel(arquivo, skiprows=2)
    dfs.append(df_unico)

# concatena todos em um único dataframe
df_notas = pd.concat(dfs, ignore_index=True)


# RELATÓRIO DE FREQUÊNCIA DOS ESTUDANTES
# caminho da pasta onde estão os arquivos
pasta = r"C:\Users\hugob\Downloads\Frequencia - Estudantes"

# lista todos os arquivos .xlsx da pasta
arquivos = glob.glob(os.path.join(pasta, "*.xlsx"))

# lista para armazenar os dataframes
dfs = []

for arquivo in tqdm(arquivos, desc="Processando arquivos"):
    # lê cada arquivo, pulando as 2 primeiras linhas
    df_unico = pd.read_excel(arquivo, skiprows=4)
    dfs.append(df_unico)

# concatena todos em um único dataframe
df_frequencia = pd.concat(dfs, ignore_index=True)


# TRATAMENTO DOS DADOS
# Excluir SITUAÇÃO = "TRANSFERIDO" e "FALECIDO" do relatório de matrículas.
df_matriculas = df_matriculas[~df_matriculas['SITUAÇÃO'].isin(['TRANSFERIDO', 'FALECIDO'])]

# Manter somente as séries de interesse nas 3 bases de dados
# Manter só Anos Finais e Ensino Médio:
series_desejados = ['1ª SÉRIE',
                    '2ª SÉRIE',
                    '3ª SÉRIE',
                    '6º Ano',
                    '7º Ano',
                    '8º Ano',
                    '9º Ano',
                    '6º ANO',
                    '7º ANO',
                    '8º ANO',
                    '9º ANO']

df_matriculas = df_matriculas[df_matriculas['SÉRIE'].isin(series_desejados)]
df_notas = df_notas[df_notas['SÉRIE'].isin(series_desejados)]
df_frequencia = df_frequencia[df_frequencia['SÉRIE'].isin(series_desejados)]

# Substituição das séries e manter padronização
mapeamento = {
    '6º Ano': '6º ANO',
    '7º Ano': '7º ANO',
    '8º Ano': '8º ANO',
    '9º Ano': '9º ANO'
}

df_matriculas['SÉRIE'] = df_matriculas['SÉRIE'].replace(mapeamento)
df_notas['SÉRIE'] = df_notas['SÉRIE'].replace(mapeamento)
df_frequencia['SÉRIE'] = df_frequencia['SÉRIE'].replace(mapeamento)


# Criar nota média no df_notas
# Substituir vírgula por ponto para reconhecimento das notas como números:
colunas_para_converter = [
        "NOTA 1º BIMESTRE",
        "NOTA 2º BIMESTRE",
        "NOTA 3º BIMESTRE",
        "NOTA 4º BIMESTRE",
        "MÉDIA ANUAL",
        "EXAME FINAL",
        "AVALIAÇÃO ESPECIAL",
        "MÉDIA FINAL"] 

for col in colunas_para_converter:
    if col in df_notas.columns:  # só executa se a coluna estiver no DataFrame
        # Substitui vírgula por ponto
        df_notas[col] = df_notas[col].astype("string").str.replace(",", ".")
        # Converte para float, erros viram NaN
        df_notas[col] = pd.to_numeric(df_notas[col], errors="coerce")



# Nota média por componente
'''
Se as colunas têm valores → média das colunas com valores.
Se alguma não tem valor → retorna a média das colunas com valor.
Se todas são NaN → retorna NaN.
'''

df_notas['MEDIA_NOTAS'] = df_notas[['NOTA 1º BIMESTRE', 'NOTA 2º BIMESTRE', 'NOTA 3º BIMESTRE']].mean(axis=1, skipna=True)

# Criar uma coluna para Aprovado ou Reprovado por componente (reprovação caso a média seja menor que 6)
# (sem nota caso os três bimestres sejam NaN)
df_notas['STATUS'] = np.select(
    [
        df_notas['MEDIA_NOTAS'].isna(),
        df_notas['MEDIA_NOTAS'].ge(6).fillna(False)
    ],
    [
        'Sem lançamentos',
        'Aprovado'
    ],
    default='Reprovado'
)


# Selecionar somente os componentes da BNCC
bncc = ['Arte',
        'Biologia',
        'Educação Física',
        'Filosofia',
        'Física',
        'Geografia',
        'História',
        'Língua Inglesa',
        'Língua Portuguesa',
        'Matemática',
        'Química',
        'Sociologia', 
        'Ciências']


df_notas = df_notas[df_notas['COMPONENTE CURRICULAR'].isin(bncc)]

# No df_notas, garantir que cada estudante seja contado apenas uma vez, por componente
dedup_cols = [
    'MATRÍCULA ESTUDANTE',
    'SÉRIE',
    'COMPONENTE CURRICULAR']

df_notas = df_notas.drop_duplicates(subset=dedup_cols)

# No df_frequencia, se tiver valores duplicados de 'MATRÍCULA', somar os valores de aulas dadas e previstas para ter o percentual final
# Garantir que as colunas numéricas estejam como número
df_frequencia['AULAS DADAS'] = pd.to_numeric(df_frequencia['AULAS DADAS'], errors='coerce')
df_frequencia['PRESENÇAS'] = pd.to_numeric(df_frequencia['PRESENÇAS'], errors='coerce')

# Agrupar por matrícula e somar as aulas e presenças
df_frequencia = (
    df_frequencia
    .groupby('MATRÍCULA', as_index=False)
    .agg({
        'DIREC': 'first',
        'MUNICÍPIO': 'first',
        'INEP ESCOLA': 'first',
        'ESCOLA': 'first',
        'ALUNO': 'first',
        'MATRÍCULA': 'first',
        'ETAPA DE ENSINO': 'first',
        'SÉRIE': 'first',
        'AULAS PREVISTAS': 'sum',
        'AULAS DADAS': 'sum',
        'PRESENÇAS': 'sum'
    })
)

# Recalcular o percentual final de frequência
df_frequencia['% FREQUÊNCIA'] = (
    df_frequencia['PRESENÇAS'] / df_frequencia['AULAS DADAS']
) * 100

# Criar coluna de "ETAPA_RESUMIDA" para indicar Anos Finais ou Ensino Médio, de acordo com a série
mapeamento_etapa = {
    '1ª SÉRIE': 'Ensino Médio',
    '2ª SÉRIE': 'Ensino Médio',
    '3ª SÉRIE': 'Ensino Médio',
    '6º ANO': 'Ens. Fund. - Anos Finais',
    '7º ANO': 'Ens. Fund. - Anos Finais',
    '8º ANO': 'Ens. Fund. - Anos Finais',
    '9º ANO': 'Ens. Fund. - Anos Finais'
}

df_matriculas['ETAPA_RESUMIDA'] = df_matriculas['SÉRIE'].map(mapeamento_etapa)


# Excluir duplicatas de matrículas
# Manter os valores de Matrícula com data de operação mais recente, caso haja duplicidade de matrícula
# Converter 'DATA DA OPERAÇÃO' para datetime
df_matriculas['DATA DA OPERAÇÃO'] = pd.to_datetime(
    df_matriculas['DATA DA OPERAÇÃO'],
    dayfirst=True,
    errors='coerce'
)

# Ordenar pela data mais recente e exclui duplicatas de matrículas
df_matriculas = (
    df_matriculas
    .sort_values('DATA DA OPERAÇÃO', ascending=False)
    .drop_duplicates(subset='MATRÍCULA', keep='first'))



# Dataframes de Ensino Médio e Anos Finais
df_matriculas_em = df_matriculas[df_matriculas['ETAPA_RESUMIDA'] == 'Ensino Médio']
df_matriculas_ef = df_matriculas[df_matriculas['ETAPA_RESUMIDA'] == 'Ens. Fund. - Anos Finais']


# Selecionar as matrículas de EPT e Noturno 
# NOTURNO
turmas_noturnas = ['EMPN1A', 
                   'EMPN1B',
                   'EMPN1C',
                   'EMPN1D',
                   'EMPN2A',
                   'EMPN2B',
                   'EMPN2C',
                   'EMPN2D',
                   'EMPN3A',
                   'EMPN3B',
                   'EMPN3C',
                   'EMPN3D']

# Matrículas dos estudantes noturnos
df_matriculas_noturno = df_matriculas[df_matriculas['TURMA'].isin(turmas_noturnas)]

# EPT
# A partir do df_matriculas, na coluna 'ETAPA DE ENSINO', selecionar etapas que não são do regular, EJATEC, subsequente ou concomitante

etapas_excluir_ept = ['ENSINO FUNDAMENTAL',
                      'ENSINO MÉDIO POTIGUAR',
                      'CEJA ENSINO FUNDAMENTAL',
                      'EJA - ENSINO MÉDIO',
                      'EJA - ENS FUNDAMENTAL - 2º SEGMENTO ANUAL',
                      'CEJA ENSINO MEDIO',
                      'PROJETO CAMINHOS DE APRENDIZAGENS: AVEXADAS PARA APRENDER',
                      'EJA - ENS FUNDAMENTAL - 1º SEGMENTO ANUAL',
                      'ENSINO FUNDAMENTAL - MULTISSERIADA',
                      'ENSINO MÉDIO',
                      'ENSINO MÉDIO - EDUCAÇÃO DE JOVENS E ADULTOS',
                      'EDUCAÇÃO ESPECIAL',
                      'ENSINO FUNDAMENTAL - ANOS FINAIS',
                      'FUNDAMENTAL ANOS INICIAIS - EJA',
                      'ENSINO FUND. 1º SEGMENTO (ANOS INICIAIS) - EDUCAÇÃO DE JOVENS E ADULTOS',
                      'ENSINO FUND. 2º SEGMENTO (ANOS FINAIS) - EDUCAÇÃO DE JOVENS E ADULTOS',
                      'ENSINO MÉDIO - EJA',
                      'FUNDAMENTAL ANOS FINAIS - EJA'
                      ]

df_matriculas_ept = df_matriculas[~df_matriculas['ETAPA DE ENSINO'].isin(etapas_excluir_ept)]

# Retirar linhas que na coluna 'ETAPA DE ENSINO' tenham as palavras "SUBSEQUENTE", "CONCOMITANTE" ou "EJATEC"
df_matriculas_ept = df_matriculas_ept[
    ~df_matriculas_ept["ETAPA DE ENSINO"].str.contains(
        r"SUBSEQUENTE|CONCOMITANTE|EJATEC",
        case=False,
        na=False
    )
].copy()


##########################################################################
'''
A partir da base por componente, criar a base por estudante com as variáveis:
- idade;
- frequência;
- classificação frequência;
- classificação rendimento;
'''

# IDADE
# Criar a coluna 'IDADE' no df_matriculas, a partir da coluna 'DATA DE NASCIMENTO'
# Garantir que a coluna seja do tipo datetime
df_matriculas['DATA DE NASCIMENTO'] = pd.to_datetime(
    df_matriculas['DATA DE NASCIMENTO'],
    errors='coerce',
    dayfirst=True
)

# Calcular a idade em anos completos
hoje = datetime.now()

df_matriculas['IDADE'] = df_matriculas['DATA DE NASCIMENTO'].map(
    lambda x: relativedelta(hoje, x).years if pd.notna(x) else pd.NA
)


# FREQUÊNCIA
# Inserir a coluna de frequência no df_matriculas, a partir do df_frequencia
# Inserir somente a coluna % FREQUÊNCIA no df_matriculas
df_matriculas = df_matriculas.merge(
    df_frequencia[['MATRÍCULA', '% FREQUÊNCIA']],
    how='left',
    on='MATRÍCULA'
)



# CLASSIFICACAO DA FREQUENCIA
# Classificação de Frequência na coluna 'TERMOMETRO_FREQUENCIA' no df_matriculas, a partir da coluna '% FREQUÊNCIA'
condicoes = [
    (
        df_matriculas['% FREQUÊNCIA'].isna() |
        df_matriculas['% FREQUÊNCIA'].eq(0) |
        df_matriculas['SITUAÇÃO'].isin(['CANCELADO', 'DEIXOU DE FREQUENTAR'])
    ),

    df_matriculas['% FREQUÊNCIA'].gt(0) &
    df_matriculas['% FREQUÊNCIA'].le(50),

    df_matriculas['% FREQUÊNCIA'].gt(50) &
    df_matriculas['% FREQUÊNCIA'].le(75),

    df_matriculas['% FREQUÊNCIA'].gt(75)
]

classificacoes = [
    'Abandono',
    'Risco de abandono',
    'Risco de reprovação por falta',
    'Estudantes fora de risco'
]

df_matriculas['TERMOMETRO_FREQUENCIA'] = np.select(
    condicoes,
    classificacoes,
    default='Sem informação'
)


# CLASSIFICACAO DE RENDIMENTO
# Criar 3 colunas no df_matriculas: Aprovacoes; Reprovacoes; Sem_lancamentos

# Aprovacoes
# Contar aprovações por matrícula no df_notas
aprovacoes = (
    df_notas[df_notas['STATUS'] == 'Aprovado']
    .groupby('MATRÍCULA ESTUDANTE')
    .size()
)

# Inserir a quantidade de aprovações no df_matriculas
df_matriculas['Aprovacoes'] = (
    df_matriculas['MATRÍCULA']
    .map(aprovacoes)
    .fillna(0)
    .astype(int)
)

# Reprovacoes
reprovacoes = (
    df_notas[df_notas['STATUS'] == 'Reprovado']
    .groupby('MATRÍCULA ESTUDANTE')
    .size()
)

# Inserir a quantidade de reprovações no df_matriculas
df_matriculas['Reprovacoes'] = (
    df_matriculas['MATRÍCULA']
    .map(reprovacoes)
    .fillna(0)
    .astype(int)
)

# Sem_lancamentos
sem_lancamentos = (
    df_notas[df_notas['STATUS'] == 'Sem lançamentos']
    .groupby('MATRÍCULA ESTUDANTE')
    .size()
)

# Inserir a quantidade de sem lancamentos no df_matriculas
df_matriculas['Sem_lancamentos'] = (
    df_matriculas['MATRÍCULA']
    .map(sem_lancamentos)
    .fillna(0)
    .astype(int)    
)

# Classificação de Rendimento = coluna 'TERMOMETRO_RENDIMENTO' no df_matriculas
condicoes = [
    (
        (df_matriculas['Aprovacoes'] == 0) &
        (df_matriculas['Reprovacoes'] == 0)
    ),
    
    # Ensino Fundamental - Anos Finais
    (
        (df_matriculas['ETAPA_RESUMIDA'] == 'Ens. Fund. - Anos Finais') &
        (df_matriculas['Sem_lancamentos'] >= 4)
    ),

    (
        (df_matriculas['ETAPA_RESUMIDA'] == 'Ens. Fund. - Anos Finais') &
        (df_matriculas['Reprovacoes'] >= 4)
    ),

    (
        (df_matriculas['ETAPA_RESUMIDA'] == 'Ens. Fund. - Anos Finais') &
        (df_matriculas['Reprovacoes'].isin([2, 3]))
    ),

    (
        (df_matriculas['ETAPA_RESUMIDA'] == 'Ens. Fund. - Anos Finais') &
        (df_matriculas['Reprovacoes'] == 1)
    ),

    (
        (df_matriculas['ETAPA_RESUMIDA'] == 'Ens. Fund. - Anos Finais') &
        (df_matriculas['Reprovacoes'] == 0)
    ),

    # Ensino Médio
    (
        (df_matriculas['ETAPA_RESUMIDA'] == 'Ensino Médio') &
        (df_matriculas['Sem_lancamentos'] >= 6)
    ),

    (
        (df_matriculas['ETAPA_RESUMIDA'] == 'Ensino Médio') &
        (df_matriculas['Reprovacoes'] >= 7)
    ),

    (
        (df_matriculas['ETAPA_RESUMIDA'] == 'Ensino Médio') &
        (df_matriculas['Reprovacoes'].isin([3, 4, 5, 6]))
    ),

    (
        (df_matriculas['ETAPA_RESUMIDA'] == 'Ensino Médio') &
        (df_matriculas['Reprovacoes'].isin([1, 2]))
    ),

    (
        (df_matriculas['ETAPA_RESUMIDA'] == 'Ensino Médio') &
        (df_matriculas['Reprovacoes'] == 0)
    )
]

classificacoes = [
    'Baixo Lançamento', # Aprovacoes = 0 e Reprovacoes = 0
    'Baixo Lançamento', # Ensino Fundamental: Sem_lancamentos >= 4
    'Alto Risco',
    'Risco',
    'Atenção',
    'Adequado',
    'Baixo Lançamento', # Ensino Médio: Sem_lancamentos >= 6
    'Alto Risco',
    'Risco',
    'Atenção',
    'Adequado'
]

df_matriculas['TERMOMETRO_RENDIMENTO'] = np.select(
    condicoes,
    classificacoes,
    default='Sem classificação'
)


############################################################
# Grupos de acordo com os pedidos de segmentação
# Criar tabelas de resumo
def criar_tabela_resumo(df, coluna):
    contagem = df[coluna].value_counts(dropna=False)

    tabela = pd.DataFrame({
        'Classificação': contagem.index.astype('string').fillna('Sem classificação'),
        'Quantidade': contagem.values,
        'Percentual (%)': (contagem.values / len(df) * 100).round(2)
    })

    return tabela


# GERAL
# Tabela de frequência
tabela_frequencia_geral = criar_tabela_resumo(
    df_matriculas,
    'TERMOMETRO_FREQUENCIA'
)

# Tabela de rendimento
tabela_rendimento_geral = criar_tabela_resumo(
    df_matriculas,
    'TERMOMETRO_RENDIMENTO'
)


# ENSINO FUNDAMENTAL
# Selecionar somente os estudantes do Ensino Fundamental
df_matriculas_EF = df_matriculas[
    df_matriculas['MATRÍCULA'].isin(df_matriculas_ef['MATRÍCULA'])
].copy()


# Tabela de frequência
tabela_frequencia_EF = criar_tabela_resumo(
    df_matriculas_EF,
    'TERMOMETRO_FREQUENCIA'
)

# Tabela de rendimento
tabela_rendimento_EF = criar_tabela_resumo(
    df_matriculas_EF,
    'TERMOMETRO_RENDIMENTO'
)


# ENSINO MÉDIO
# Selecionar somente os estudantes do Ensino Médio
df_matriculas_EM = df_matriculas[
    df_matriculas['MATRÍCULA'].isin(df_matriculas_em['MATRÍCULA'])
].copy()


# Tabela de frequência
tabela_frequencia_EM = criar_tabela_resumo(
    df_matriculas_EM,
    'TERMOMETRO_FREQUENCIA'
)

# Tabela de rendimento
tabela_rendimento_EM = criar_tabela_resumo(
    df_matriculas_EM,
    'TERMOMETRO_RENDIMENTO'
)


# EPT
# Selecionar somente os estudantes de EPT
df_matriculas_EPT = df_matriculas[
    df_matriculas['MATRÍCULA'].isin(df_matriculas_ept['MATRÍCULA'])
].copy()


# Tabela de frequência
tabela_frequencia_EPT = criar_tabela_resumo(
    df_matriculas_EPT,
    'TERMOMETRO_FREQUENCIA'
)

# Tabela de rendimento
tabela_rendimento_EPT = criar_tabela_resumo(
    df_matriculas_EPT,
    'TERMOMETRO_RENDIMENTO'
)


# NOTURNO
# Selecionar somente os estudantes do Ensino Noturno
df_matriculas_NT = df_matriculas[
    df_matriculas['MATRÍCULA'].isin(df_matriculas_noturno['MATRÍCULA'])
].copy()


# Tabela de frequência
tabela_frequencia_NT = criar_tabela_resumo(
    df_matriculas_NT,
    'TERMOMETRO_FREQUENCIA'
)

# Tabela de rendimento
tabela_rendimento_NT = criar_tabela_resumo(
    df_matriculas_NT,
    'TERMOMETRO_RENDIMENTO'
)


# Estudantes com 15+ no EF
# Filtra estudantes com 15 anos ou mais nos Anos Finais do Ensino Fundamental
df_filtro_15 = df_matriculas[
    (df_matriculas['IDADE'] >= 15) &
    (df_matriculas['ETAPA_RESUMIDA'] == 'Ens. Fund. - Anos Finais')
].copy()

# Quantidade de estudantes por DIREC
tabela_DIREC_EF15 = (
    df_filtro_15
    .groupby('DIREC')
    .size()
    .reset_index(name='Quantidade de estudantes')
)

# Calcula o percentual sobre o total de estudantes filtrados
total_estudantes = len(df_filtro_15)

tabela_DIREC_EF15['Percentual (%)'] = (
    tabela_DIREC_EF15['Quantidade de estudantes'] / total_estudantes * 100
).round(2)



# Estudantes com 18+ no EM
# Filtra estudantes com 18 anos ou mais no Ensino Médio
df_filtro_18 = df_matriculas[
    (df_matriculas['IDADE'] >= 18) &
    (df_matriculas['ETAPA_RESUMIDA'] == 'Ensino Médio')
].copy()

# Quantidade de estudantes por DIREC
tabela_DIREC_EM18 = (
    df_filtro_18
    .groupby('DIREC')
    .size()
    .reset_index(name='Quantidade de estudantes')
)

# Calcula o percentual sobre o total de estudantes filtrados
total_estudantes = len(df_filtro_18)

tabela_DIREC_EM18['Percentual (%)'] = (
    tabela_DIREC_EM18['Quantidade de estudantes'] / total_estudantes * 100
).round(2)


# Estudantes com bom rendimento que reprovaram ou estão próximos de reprovar por falta
# Filtro para universo da análise: Estudantes com 'TERMOMETRO_FREQUENCIA' = 'Risco de reprovação por falta' ou 'Risco de abandono'
# Estudantes com 'TERMOMETRO_RENDIMENTO' = 'Adequado' e 'Atenção' (com divisão desses grupos)
df_reprov = df_matriculas[
    (df_matriculas['TERMOMETRO_FREQUENCIA'] == 'Risco de reprovação por falta') |
    (df_matriculas['TERMOMETRO_FREQUENCIA'] == 'Risco de abandono')
].copy()


# Filtra as duas classificações de rendimento
df_rendimento = df_reprov[
    df_reprov['TERMOMETRO_RENDIMENTO'].isin(['Adequado', 'Atenção'])
].copy()

# Agrupa por DIREC e classificação, contando os estudantes
tabela_final = (
    df_rendimento
    .groupby(['DIREC', 'TERMOMETRO_RENDIMENTO'])
    .size()
    .reset_index(name='Quantidade')
)

# Calcula o percentual sobre o total geral do df_reprov
total_geral = len(df_rendimento)

tabela_final['Percentual (%)'] = (
    tabela_final['Quantidade'] / total_geral * 100
).round(2)

# Ordena por DIREC e classificação
tabela_final = tabela_final.sort_values(
    ['DIREC', 'TERMOMETRO_RENDIMENTO']
).reset_index(drop=True)



# Estudantes que "deixaram de frequentar" e possuem com frequência suficiente para retornar
# Filtro para universo: SITUAÇÃO = 'CANCELADO' ou 'DEIXOU DE FREQUENTAR'
# TERMOMETRO_FREQUENCIA = 'Estudantes fora de risco'
# Filtra o universo desejado
df_cancelados = df_matriculas[
    (
        df_matriculas['SITUAÇÃO'].isin([
            'CANCELADO',
            'DEIXOU DE FREQUENTAR'
        ])
    ) &
    (
        df_matriculas['TERMOMETRO_FREQUENCIA'] == 'Estudantes fora de risco'
    )
].copy()

# Quantidade de estudantes por DIREC
tabela_cancelados = (
    df_cancelados
    .groupby('DIREC')
    .size()
    .reset_index(name='Quantidade')
)

# Percentual sobre o total geral do universo filtrado
total_geral = len(df_cancelados)

tabela_cancelados['Percentual (%)'] = (
    tabela_cancelados['Quantidade'] / total_geral * 100
).round(2)

# Ordena pela quantidade, da maior para a menor
tabela_cancelados = tabela_cancelados.sort_values(
    'Quantidade',
    ascending=False
).reset_index(drop=True)


### Não tem valores que atendem esse critério final aqui.
# Todos os estudantes com SITUAÇÃO = 'CANCELADO' ou 'DEIXOU DE FREQUENTAR', estão com 'TERMOMETRO_FREQUENCIA' = 'Abandono', por definição, essa cracterística fica na regra para criação desse grupo.




# Salvar em Excel o dataframe
with pd.ExcelWriter(r"D:\Scripts_Python\FGV\Indicadores_Fluxo_Rendimento_Frequencia\Indicadores_Fluxo_Rendimento_Frequencia.xlsx") as writer:
    df_matriculas.to_excel(writer, sheet_name='Base fonte', index=False)
    tabela_frequencia_geral.to_excel(writer, sheet_name='Geral - Frequência', index=False)
    tabela_rendimento_geral.to_excel(writer, sheet_name='Geral - Rendimentos', index=False)
    tabela_frequencia_EF.to_excel(writer, sheet_name='Ens Fund AF - Frequência', index=False)
    tabela_rendimento_EF.to_excel(writer, sheet_name='Ens Fund AF - Rendimentos', index=False)
    tabela_frequencia_EM.to_excel(writer, sheet_name='Ens Medio - Frequência', index=False)
    tabela_rendimento_EM.to_excel(writer, sheet_name='Ens Medio - Rendimentos', index=False)
    tabela_frequencia_EPT.to_excel(writer, sheet_name='EPT - Frequência', index=False)
    tabela_rendimento_EPT.to_excel(writer, sheet_name='EPT - Rendimentos', index=False)
    tabela_frequencia_NT.to_excel(writer, sheet_name='Noturno - Frequência', index=False)
    tabela_rendimento_NT.to_excel(writer, sheet_name='Noturno - Rendimentos', index=False)
    tabela_DIREC_EF15.to_excel(writer, sheet_name='Idade 15+ EF', index=False)
    tabela_DIREC_EM18.to_excel(writer, sheet_name='Idade 18+ EM', index=False)
    tabela_final.to_excel(writer, sheet_name='Bom Rendimento e Ausencia', index=False)
    tabela_cancelados.to_excel(writer, sheet_name='Deixaram de frequentar', index=False)



