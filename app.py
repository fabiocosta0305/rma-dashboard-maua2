# ------------------------------
# Imports
# ------------------------------
import pandas as pd
import panel as pn
import hvplot.pandas
import os
import traceback


# ------------------------------
# Extensão Panel
# ------------------------------
pn.extension()

# ------------------------------
# 1. Ler a planilha
# ------------------------------
#caminho = r"C:\Users\soareseas\Documents\Ezequiel\RMA projeto\CONSOLIDADO 2024 JAN-DEZ.xlsx"
caminho_arquivo = os.path.join(os.path.dirname(__file__), "CONSOLIDADO 2024 JAN-DEZ.xlsx")
aba = "CRAS GERAL"

# Lendo sem cabeçalho
df_raw = pd.read_excel(caminho_arquivo, sheet_name=aba, header=None)

# ------------------------------
# 2. Definir meses
# ------------------------------
meses = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
         "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]

col_inicio = 39  # coluna Janeiro
col_fim = 50     # coluna Dezembro

# ------------------------------
# Extrair dados da seção Recepção (1.1)
# ------------------------------
dados_recepcao = df_raw.iloc[13:16, col_inicio:col_fim+1]
dados_recepcao.index = ["Total", "Individual", "Coletivo"]
dados_recepcao.columns = meses

# ------------------------------
# Extrair dados da seção Recepção (1.2)
# ------------------------------
dados_recepcao_horarios = df_raw.iloc[17:22, col_inicio: col_fim +1]
dados_recepcao_horarios_lista = df_raw.iloc[17:22, 0].tolist()
dados_recepcao_horarios.index = dados_recepcao_horarios_lista
dados_recepcao_horarios.columns = meses

# ------------------------------
# Procedência (1.3)
# ------------------------------
procedencia = df_raw.iloc[23:44, col_inicio:col_fim+1]
tipos_procedencia = df_raw.iloc[23:44, 0].tolist()
procedencia.index = tipos_procedencia[:len(procedencia)]
totais = procedencia.sum(axis=1)
top_procedencia = (
    totais.sort_values(ascending=False)
    .head(5)
    .reset_index()
)
top_procedencia.columns = ["Procedência", "Total"]

# ------------------------------
# Demandas (1.4)
# ------------------------------
demandas = df_raw.iloc[45:76, col_inicio:col_fim+1]
nomes_demandas = df_raw.iloc[45:76, 0].tolist()
demandas.columns = meses
demandas.index = nomes_demandas
totais_demandas = demandas.sum(axis=1)
top_demandas = (
    totais_demandas.sort_values(ascending=False)
    .head(10)
    .reset_index()
)
top_demandas.columns = ["Demandas", "Total"]

# ------------------------------
# Tipo de Atendimento (1.6)
# ------------------------------
dados_recepcao_tipo_atend = df_raw.iloc[81:84, col_inicio: col_fim +1]
dados_recepcao_tipo_atend_lista = df_raw.iloc[81:84, 0].tolist()
dados_recepcao_tipo_atend.index = dados_recepcao_tipo_atend_lista
dados_recepcao_tipo_atend.columns = meses

# ------------------------------
# Nacionalidade (1.7)
# ------------------------------
dados_recepcao_nacionalidade = df_raw.iloc[85:92, col_inicio:col_fim+1]
dados_recepcao_nacionalidade_lista = df_raw.iloc[85:92, 0].tolist()
dados_recepcao_nacionalidade.index = dados_recepcao_nacionalidade_lista
dados_recepcao_nacionalidade.columns = meses
dados_recepcao_nacionalidade_sembrasileiros = dados_recepcao_nacionalidade.drop("Brasileiro", errors="ignore")
#adicionado para calcular o total
nacionalidade_total_anual = pd.DataFrame({
    "Nacionalidade": dados_recepcao_nacionalidade_sembrasileiros.index,
    "Total": dados_recepcao_nacionalidade_sembrasileiros.sum(axis=1).values
})
# Ordenar do maior para o menor
nacionalidade_total_anual = nacionalidade_total_anual.sort_values("Total", ascending=False)

# ------------------------------
# Idioma (1.8)
# ------------------------------
dados_recepcao_idioma = df_raw.iloc[93:98, col_inicio:col_fim+1].copy()
dados_recepcao_idioma.index = df_raw.iloc[93:98, 0].values
dados_recepcao_idioma = dados_recepcao_idioma.drop("Português", errors="ignore")
dados_recepcao_idioma_total = pd.DataFrame({
    "Idioma": dados_recepcao_idioma.index,
    "Total": dados_recepcao_idioma.sum(axis=1).values
})

# ------------------------------
# Deficiência (1.9)
# ------------------------------
dados_recepcao_deficiencia = df_raw.iloc[99:107, col_inicio:col_fim + 1].copy()
dados_recepcao_deficiencia.index = df_raw.iloc[99:107, 0].values
dados_recepcao_deficiencia = dados_recepcao_deficiencia.drop("Sem deficiência", errors="ignore")
dados_recepcao_deficiencia_total = pd.DataFrame({
    "Deficiência": dados_recepcao_deficiencia.index,
    "Total": dados_recepcao_deficiencia.sum(axis=1).values
})

# ------------------------------
# Famílias cadastradas por CRAS
# ------------------------------
abas = [
    "CRAS FALCHI", "CRAS FEITAL", "CRAS MACUCO", "CRAS ORATORIO",
    "CRAS PARQUE", "CRAS SAO JOAO", "CRAS VILA", "CRAS ZAIRA"
]
linha_familias_cadastradas = 131
dados_familias_cadastradas = {}
for aba in abas:
    df_raw_unique = pd.read_excel(caminho_arquivo, sheet_name=aba, header=None)
    valores = df_raw_unique.iloc[linha_familias_cadastradas, col_inicio:col_fim + 1].tolist()
    dados_familias_cadastradas[aba.replace("CRAS ", "")] = valores
dados_familias_cadastradas = pd.DataFrame(dados_familias_cadastradas, index=meses).T

# ------------------------------
# PAIF / SCFV / Perfis
# ------------------------------
acompanhamentos_paif_geral = df_raw.iloc[168:169, col_inicio: col_fim +1]
acompanhamentos_paif_lista = df_raw.iloc[168:169, 0].tolist()
acompanhamentos_paif_geral.index = acompanhamentos_paif_lista
acompanhamentos_paif_geral.columns = meses

# Linha onde está o total de famílias acompanhamento PAIF Total

acompanhamentos_paif_total = df_raw.iloc[374:375, col_inicio: col_fim +1]
# Pegar os nomes das demandas na coluna A (coluna 0)
acompanhamentos_paif_total_lista = df_raw.iloc[374:375, 0].tolist()
acompanhamentos_paif_total.index = ["Total de famílias em acompanhamento pelo PAIF A.1"]
acompanhamentos_paif_total.columns = meses

#calculo para porcentagem de familias que estão participando de grupos
# Garante que os valores estão em formato numérico
acompanhamentos_paif_geral = acompanhamentos_paif_geral.apply(pd.to_numeric, errors='coerce')
acompanhamentos_paif_total = acompanhamentos_paif_total.apply(pd.to_numeric, errors='coerce')

# Calcula o percentual (geral / total * 100)
percentual_participacao_paif = (acompanhamentos_paif_geral.values / acompanhamentos_paif_total.values)*100

# Cria um DataFrame com o resultado
percentual_participacao_paif = pd.DataFrame(
    percentual_participacao_paif,
    index=["% de Famílias com Participação Regular em Grupos (PAIF)"],
    columns=meses
)
#fim do trecho novo

scfv_totais = df_raw.iloc[187:191, col_inicio:col_fim + 1]
scfv_totais.index = ["Execução Direta Cras","Bombeiro Mirim","Execução Indireta","Totais"]
scfv_totais.columns = meses

atend_particularizados = df_raw.iloc[259:262, col_inicio:col_fim + 1]
atend_particularizados.index = ["Atendimentos particularizados Cras","Total de Visitas Domiciliares","Total"]
atend_particularizados.columns = meses

perfil_familias = df_raw.iloc[263:277, col_inicio:col_fim + 1]
perfil_familias.columns = meses
perfil_familias.index = [
    "Extrema pobreza","Bolsa Família","PBF - Descumprimento",
    "PBF - Suspensão Condicionalidades","PBF - Suspensão SICON",
    "Renda Cidadã","BPC","BPC Escola","PAI - Idoso","SCFV",
    "Trabalho Infantil","Serviço de Acolhimento","Ação Jovem","Cesta Básica"
]
top_perfis_df = (
    perfil_familias.sum(axis=1)
    .sort_values(ascending=False)
    .head(5)
    .reset_index()
)
top_perfis_df.columns = ["Perfil", "Total"]
perfil_familias_top = perfil_familias.loc[top_perfis_df["Perfil"]]

# ------------------------------
# 4. Criar gráfico interativo com hvPlot
# ------------------------------
grafico_recepcao = dados_recepcao.T.hvplot.bar(
    height=500,
    width=1400,
    title="Recepção - Quantidade de Pessoas - Cras Geral (2024)",
    xlabel="Meses",
    ylabel="Quantidade de Pessoas",
    rot=90,  # rotação dos meses
    tools=["hover"],
    stacked=False,  # barras lado a lado
    bar_width=0.7
)

grafico_recepca_horarios = dados_recepcao_horarios.T.hvplot.scatter(
    height=500,
    width=1400,
    title="Atendimentos por Horário - CRAS Geral 2024",
    xlabel="Meses",
    ylabel="Número de Atendimentos",
    #rot=90,  # rotação dos meses
    tools=["hover"],
    stacked=False  # barras lado a lado
) * dados_recepcao_horarios.T.hvplot.line(line_width=2)

# ------------------------------
# Procedência - Gráfico Interativo
# ------------------------------

grafico_procedencia = top_procedencia.hvplot.bar(
    x="Procedência",
    y="Total",
    height=500,
    width=1300,
    title="Procedência - Top 5 Atendimentos - CRAS Geral (2024)",
    xlabel="Procedência",
    ylabel="Número de Atendimentos",
    rot=0,
    tools=["hover"],
    # color="skyblue",
    line_color="black",
    cmap='Set3',
    color="Procedência",
    bar_width=0.6
)

grafico_demandas = top_demandas.hvplot.barh(
    x="Demandas",
    y="Total",
    height=500,
    width=1200,
    title="Top 10 Demandas - CRAS Geral (2024)",
    xlabel="Demandas",
    ylabel="Número de Atendimentos",
    rot=0,
    tools=["hover"],
    color="Demandas",
    cmap="Set3",
    line_color="black"
)

# Gráfico limpo e visualmente agradável
grafico_recepcao_nacionalidade = nacionalidade_total_anual.hvplot.bar(
    x="Nacionalidade",
    y="Total",
    height=500,
    width=1200,
    title="Total Anual de Atendimentos por Nacionalidade - CRAS Geral (2024)",
    xlabel="Nacionalidade",
    ylabel="Total de Atendimentos",
    color="Nacionalidade",
    cmap="Category20",
    line_color="black",
    tools=["hover"],
    legend=False,
    ylim=(0, nacionalidade_total_anual["Total"].max() + 5),
)

# Gráfico com ajustes visuais
grafico_recepcao_idioma_total = dados_recepcao_idioma_total.hvplot.bar(
    x="Idioma",
    y="Total",
    height=500,
    width=1200,
    title="Total de Atendimentos por Idioma - CRAS Geral (2024)",
    xlabel="Idioma",
    ylabel="Número de Atendimentos",
    #rot=90,
    tools=["hover"],
    color="Idioma",
    cmap='Set3',
    line_color="black",
    ylim=(0, dados_recepcao_idioma_total["Total"].max() + 5),  # escala justa
    shared_axes=False,
    bar_width=0.5  # controla a espessura das barras
)

grafico_recepcao_deficiencia_total = dados_recepcao_deficiencia_total.hvplot.bar(
    x="Deficiência",
    y="Total",
    height=500,
    width=1200,
    title="Total de Atendimentos por Deficiência - CRAS Geral (2024)",
    xlabel="Deficiência",
    ylabel="Número de Atendimentos",
    #rot=90,
    tools=["hover"],
    color="Deficiência",
    cmap='Set3',
    line_color="black",
    ylim=(0, dados_recepcao_deficiencia_total["Total"].max() + 5),  # escala justa
    shared_axes=False,
    #    responsive=True,
    bar_width=0.5  # controla a espessura das barras
)

grafico_familias_cadastradas = dados_familias_cadastradas.T.hvplot.bar(
    stacked=True,  # barras empilhadas
    height=500,
    width=1400,
    title="Total de Famílias Cadastradas por CRAS (2024)",
    xlabel="Meses",
    ylabel="Número de Famílias",
    #rot=90,
    tools=["hover"],
    legend="top_left",
    cmap="Set3",  # paleta neutra e suave
    shared_axes=False,
    bar_width=0.6,
)

grafico_percentual_paif = percentual_participacao_paif.T.hvplot.scatter(
    height=500,
    width=1400,
    title="% de Famílias com Participação Regular em Grupos do PAIF (2024)",
    xlabel="Meses",
    ylabel="Percentual (%)",
    #color="blue",
    marker="o",
    tools=["hover"],
    hover_format="%.1f%%"
) * percentual_participacao_paif.T.hvplot.line(line_width=2,hover_format="%.1f%%")


grafico_scfv_pontos = scfv_totais.T.hvplot.scatter(
    height=500,
    width=1200,
    title="Variação mensal de participações SCFV",
    xlabel="Meses",
    ylabel="Participações",
    #rot=90,
    tools=["hover"],
    size=8
) * scfv_totais.T.hvplot.line(line_width=2)

grafico_atend_particularizados = atend_particularizados.T.hvplot.scatter(
    height=500,
    width=1400,
    title="Variação mensal de Atendimentos Particularizados",
    xlabel="Meses",
    ylabel="Participações",
    #rot=90,
    tools=["hover"],
    size=8
) * atend_particularizados.T.hvplot.line(line_width=2)

grafico_top_perfis = top_perfis_df.hvplot.bar(
    x="Perfil",
    y="Total",
    height=500,
    width=1200,
    title="Top 5 Perfis de Famílias com Atendimentos Particularizados (2024)",
    xlabel="Perfil",
    ylabel="Total Anual de Atendimentos",
    color="Perfil",  # hvplot aplica uma cor diferente pra cada categoria
    cmap="Set3",  # paleta de cores agradável
    line_color="black",
    tools=["hover"],
    bar_width=0.5
    #rot=90
)
# Gráfico de tendência mensal
grafico_perfil_tendencia = perfil_familias_top.T.hvplot.line(
    height=500,
    width=1400,
    title="Evolução Mensal dos Principais Perfis de Famílias (2024)",
    xlabel="Meses",
    ylabel="Número de Famílias",
    #rot=90,
    legend="top_left",
    line_width=2,
    cmap="Category10",
    tools=["hover"]
) * perfil_familias_top.T.hvplot.scatter(
    size=6,
    cmap="Category10"
)

# ------------------------------
# 5. Layout e template
# ------------------------------
intro = pn.pane.Markdown("# Painel de Análise Socioassistencial - RMA 2024")

# Cria o template com barra lateral
template = pn.template.MaterialTemplate(
    title="Dashboard Socioassistencial RMA",
    sidebar=[
        "## ⚙️ Filtros",
        pn.widgets.Select(name='Ano', options=[2023, 2024, 2025]),
        pn.widgets.Select(name='Mês', options=meses),
        pn.pane.Markdown("---"),
        pn.pane.Markdown("**Desenvolvido por Ezequiel Soares**"),
        pn.pane.Markdown("Prefeitura de Mauá • Setor de Ciência de Dados"),
    ],
)

# Layout principal em formato de "jornal" (um gráfico abaixo do outro)
layout = pn.Column(
    intro,
    pn.layout.Divider(),
    pn.pane.Markdown("""
    ### Recepção – Atendimentos e Perfil do Público

    A recepção é a porta de entrada dos cidadãos nos CRAS, representando o primeiro contato com a 
    Assistência Social. É nesse momento que as pessoas buscam informações, orientações e encaminhamentos 
    sobre benefícios, serviços e programas.

    O acompanhamento dos dados de atendimentos na recepção permite identificar a demanda espontânea 
    e os principais motivos de procura, possibilitando uma melhor organização das equipes e o 
    planejamento das ações de acolhimento e atendimento.

    Essas informações são fundamentais para compreender o perfil das pessoas atendidas, avaliar 
    a eficiência do atendimento inicial e identificar tendências de aumento de demanda que podem 
    indicar vulnerabilidades emergentes na comunidade."""),
    dados_recepcao,  # tabela
    grafico_recepcao.opts(shared_axes=False),
pn.pane.Markdown("""
### Atendimentos por Horário – Recepção

A análise dos atendimentos por faixa de horário, ao longo dos meses, permite observar 
os períodos de maior movimento nos CRAS.  
Essas informações ajudam a otimizar o planejamento das equipes, reduzir filas de espera 
e ajustar a alocação de profissionais conforme a demanda real.
"""),
    grafico_recepca_horarios.opts(shared_axes=False),
pn.pane.Markdown("""
### Procedências dos Usuários – Recepção

A análise da procedência dos usuários revela como as pessoas chegam até o CRAS, 
se por demanda espontânea, agendamento, indicação de outros serviços 
ou por já serem usuários acompanhados.  

Essas informações ajudam a ntender os principais canais de acesso, 
avaliar a visibilidade do serviço e aprimorar as estratégias de divulgação e acolhimento.
""", width=900),
    grafico_procedencia.opts(shared_axes=False),
    grafico_demandas.opts(shared_axes=False),
    pn.pane.Markdown("### Nacionalidade e Diversidade"),
    pn.pane.Markdown(
        "O acompanhamento da nacionalidade dos atendidos contribui para a promoção da inclusão social, "
        "garantindo que o atendimento aos imigrantes e refugiados seja adequado e culturalmente sensível."
    ),
    grafico_recepcao_nacionalidade.opts(shared_axes=False),
    pn.pane.Markdown(
        "A identificação dos idiomas mais falados ajuda a **reduzir barreiras de comunicação** "
        "e planejar ações de mediação linguística e capacitação de servidores."
    ),
    grafico_recepcao_idioma_total.opts(shared_axes=False),
    pn.pane.Markdown(
        "Esses dados auxiliam na inclusão e acessibilidade, indicando a necessidade de "
        "adequações físicas, tecnológicas e comunicacionais nos serviços."
    ),
    grafico_recepcao_deficiencia_total.opts(shared_axes=False),
    grafico_familias_cadastradas.opts(shared_axes=False),

    pn.pane.Markdown("## PAIF – Proteção e Atendimento Integral à Família"),
    pn.pane.Markdown(
        "O PAIF acompanha as fámilias que vivem em situação de vulnerabilidade, oferecendo apoio,  "
        "orientação e acesso a direitos, ajudando-as a melhorar suas condições de vida e evitar que "
        "problemas sociais se agravem."
        "O Acompanhamento dos dados sobre famílias inseridas no PAIF e sua participação regular em grupos é enssencial"
        " para avaliar o alcance das ações socioassistenciais e a efetividade das atividades coletivas."
        "Esses indicadores ajudam a Assistência Social a monitorar a adesão das famílias, identificar necessidades "
        "de maior acompanhamento e planejar estratégias para ampliar a participação e o impacto do serviço."
    ),
    acompanhamentos_paif_total,
    acompanhamentos_paif_geral,
    pn.pane.Markdown(
        "## Análise dos Percentuais de Participação no PAIF\n\n"
        "No mês de julho, observa-se um valor acima de 100% na proporção de famílias com "
        "participação regular em grupos. Esse resultado pode indicar inconsistências nos dados de origem, "
        "possivelmente relacionadas a duplicidades, divergências entre sistemas ou falhas nas fórmulas "
        "utilizadas nas planilhas de apoio.\n\n"
        "É recomendada uma verificação dos registros nos sistemas de referência e das planilhas complementares "
        "que alimentam esses indicadores, garantindo maior precisão e confiabilidade nas análises futuras.",
    ),
    grafico_percentual_paif,

pn.pane.Markdown("""
    ### SCFV – Serviço de Convivência e Fortalecimento de Vínculos

    O SCFV é um serviço que promove atividades em grupo para fortalecer os vínculos familiares e comunitários,
    prevenindo situações de risco social.

    O acompanhamento do número de participações — como execução direta pelo CRAS, parcerias com OSCs
    e projetos como o Bombeiro Mirim — permite avaliar o alcance das ações e o envolvimento dos participantes
    ao longo do ano.
    """),


grafico_scfv_pontos.opts(shared_axes=False),
pn.pane.Markdown(
    "### Atendimentos Particularizados\n\n"
    "Os atendimentos particularizados envolvem ações individuais com as famílias, como visitas domiciliares "
    "e acompanhamentos específicos. A análise desses dados ajuda a identificar o volume de atendimentos "
    "realizados ao longo do tempo e a planejar melhor a atuação das equipes do CRAS."
),
    grafico_atend_particularizados.opts(shared_axes=False),
    pn.pane.Markdown(
    "### Perfil das Famílias em Atendimento Particularizado\n\n"
    "Esta seção apresenta os principais perfis das famílias que receberam atendimentos individualizados, "
    "como aquelas em situação de extrema pobreza, beneficiárias de programas de transferência de renda "
    "(Bolsa Família, BPC, Renda Cidadã) ou com demandas específicas, como cesta básica e trabalho infantil. "
    "O acompanhamento desses perfis permite identificar vulnerabilidades recorrentes e direcionar ações "
    "mais eficazes da Assistência Social."
    ),
    grafico_top_perfis.opts(shared_axes=False),
    pn.pane.Markdown(
    "### Evolução Mensal dos Principais Perfis\n\n"
    "O gráfico apresenta a variação dos principais perfis de famílias ao longo do ano, "
    "permitindo identificar possíveis sazonalidades nos atendimentos e nas situações de vulnerabilidade. "
    "Essas informações auxiliam o planejamento das ações do PAIF e demais serviços, "
    "ajustando estratégias conforme períodos de maior demanda."
    ),
    grafico_perfil_tendencia.opts(shared_axes=False),

    sizing_mode="stretch_width",  # largura ajustável, altura natural
    scroll=True                   # adiciona rolagem vertical automática
)

# Adiciona o layout principal ao template
template.main.append(layout)

# Corrige o scroll vertical invisível
custom_css = """
.bk.pn-main {
    overflow-y: auto !important;
    overflow-x: hidden !important;
}
"""
pn.config.raw_css.append(custom_css)

# Adiciona o layout ao template e o torna servável
template.main.append(layout)
template.servable()

# no fim do arquivo:
# if __name__ == "__main__":
#     pn.serve(template, port=7860, show=False)
# else:
#     app = template

