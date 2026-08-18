import pandas as pd
from funcoes_apurar_pis import formatar_moeda


# FUNÇÃO TRATAR CTE 
def tratar_cte_icms(planilha, cnpj): # FUNÇÃO USADA PARA REALIZAR O TRATAMENTO NA PLANILHA DOS CTES E RETORNAR OS VALORES QUE ESTÃO NO RETURN
    df_cte = pd.read_excel(planilha) # AQUI FOI LIDO O ARQUIVO XLSX DE PAGINA UNICA
    colunas_desejadas = ['Tomador', 'CPF/CNPJ Tomador', 'Valor', 'Valor ICMS', 'Status'] # AQUI SELECIONANDO APENAS AS COLUNAS DESEJADAS E CRIANDO UMA COPIA 
    df_cte_colunas_corretas = df_cte[colunas_desejadas].copy() # COPIADO O DF APENAS COM AS COLUNAS DESEJADAS

    # FILTRAR APENAS PARA STATUS AUTORIZADA
    df_cte_colunas_corretas = df_cte_colunas_corretas[df_cte_colunas_corretas['Status'] == 'Autorizada']

    # RENOMEANDO PARA VALOR DO FRETE
    df_cte_colunas_corretas = df_cte_colunas_corretas.rename(columns={'Valor': 'Valor do frete'})

    # FILTRAR PELO CNPJ FORNECIDO
    filtro_cnpj = df_cte_colunas_corretas[df_cte_colunas_corretas['CPF/CNPJ Tomador'] == cnpj]

    # CRIANDO UMA VARIAVEL DE TOTAL DE FRETE PELO CNPJ FORNECIDO
    total_frete_pelo_cnpj = filtro_cnpj['Valor do frete'].sum()

    # CRIANDO UMA VARIAVEL DE TOTAL DE ICMS PELO CNPJ FORNECIDO 
    total_icms_pelo_cnpj = filtro_cnpj['Valor ICMS'].sum()

    # A FUNÇÃO RETORNA A TABELA DINAMICA, VALOR TOTAL DE COFINS, DE PIS, TOTAL DE FRETE, TOTAL ICMS SP, TOTAL ICMS SC, TOTAL FRETE SP E TOTAL FRETE SC
    return total_icms_pelo_cnpj, total_frete_pelo_cnpj


# FUNÇÃO TRATAR ICMS
def tratar_icms(planilha, cnpj):

    cnpj = cnpj.replace('.', '').replace('-', '').replace('/', '')
    df_icms = pd.read_excel(planilha, sheet_name='relatorio') # AQUI É LIDO O ARQUIVO XLSX COM O NOME DA PAGINA ESPECIFICADO

    # AQUI SELECIONANDO APENAS AS COLUNAS DESEJADAS E CRIANDO UMA COPIA
    colunas_desejadas = ['Número', 'Status', 'CNPJ Destinatário', '[Item] CFOP', '[Item] Valor Total Bruto', 'Base ICMS', '[Item] Valor ICMS'] 
    df_icms_colunas_corretas = df_icms[colunas_desejadas].copy() 

    # ALTERANDO O TIPO DA COLUNA CNPJ Destinatário PARA STRING
    df_icms_colunas_corretas['CNPJ Destinatário'] = df_icms_colunas_corretas['CNPJ Destinatário'].astype(str)

    # ALTERANDO O TIPO DA COLUNA CFOP PARA STRING
    df_icms_colunas_corretas['[Item] CFOP'] = df_icms_colunas_corretas['[Item] CFOP'].astype(str).str.extract(r'([^.]*)')

    # FILTRANDO APENAS NOTAS COM STATUS DE AUTORIZADA
    df_icms_colunas_corretas = df_icms_colunas_corretas[df_icms_colunas_corretas['Status'] == 'Autorizadas']

    # FILTRAR PELO CNPJ FORNECIDO
    filtro_cnpj = df_icms_colunas_corretas[df_icms_colunas_corretas['CNPJ Destinatário'] == cnpj]

    # CRIANDO UMA VARIAVEL DE TOTAL BRUTO PELO CNPJ FORNECIDO
    total_bruto_pelo_cnpj = filtro_cnpj['[Item] Valor Total Bruto'].sum()

    # CRIANDO UMA VARIAVEL DE TOTAL DE ICMS PELO CNPJ FORNECIDO
    total_icms_pelo_cnpj = filtro_cnpj['[Item] Valor ICMS'].sum()
    
    # TIRANDO A DUPLICIDADE DE NUMEROS DE NF PARA SOMAR A BASE ICMS CORRETA
    filtro_cnpj = filtro_cnpj.drop_duplicates(subset=['Número'], keep = 'first')

    # Calculando a soma das colunas 'ValorICMS' e 'ValorBruto'
    soma_valores = df_icms_colunas_corretas.groupby('Número')[['[Item] Valor ICMS', '[Item] Valor Total Bruto']].sum().reset_index()

    # Combinando a soma dos valores com a coluna 'ValorBase' sem duplicados
    df_para_tabela = filtro_cnpj[['Número', 'Base ICMS', 'Status', 'CNPJ Destinatário', '[Item] CFOP']].merge(soma_valores, on='Número')

    # CRIANDO UMA VARIAVEL DE TOTAL BRUTO PELO CNPJ FORNECIDO
    total_baseicms_pelo_cnpj = filtro_cnpj['Base ICMS'].sum()

    # TABELA DINAMICA PELAS CFOP
    tabela_dinamica_icms = pd.pivot_table(df_para_tabela, values=['[Item] Valor Total Bruto','Base ICMS', '[Item] Valor ICMS'], index='[Item] CFOP', aggfunc='sum')

    # CRIANDO OS TOTAIS NA ULTIMA LINHA DA TABELA DINAMICA
    total_por_coluna = tabela_dinamica_icms.sum().to_frame().T
    total_por_coluna.index = ['Total']

    # Adicionando uma nova linha ao DataFrame da tabela dinâmica com os totais
    tabela_dinamica_icms = pd.concat([tabela_dinamica_icms, total_por_coluna])
    tabela_dinamica_icms = tabela_dinamica_icms.applymap(formatar_moeda)
    tabela_dinamica_icms.columns = ['BASE ICMS', 'VALOR ICMS', 'VALOR TOTAL BRUTO']

    return total_bruto_pelo_cnpj, total_icms_pelo_cnpj, total_baseicms_pelo_cnpj, tabela_dinamica_icms


# FUNÇÃO TRATAR DRE
def tratar_dre_icms(planilha, mes):
    
    df_dre = pd.read_excel(planilha) # AQUI É LIDO O ARQUIVO XLSX DE PAGINA UNICA 
    
    # FILTRANDO APENAS O MES SELECIONADO
    df_dre = df_dre[['Descrição', mes]]

    # RETIRANDO ESPAÇOS ANTES DAS LETRAS NA COLUNA DESCRIÇÃO
    df_dre['Descrição'] = df_dre['Descrição'].str.strip()

    # AQUI FOI RECORTADO O DF ATÉ A LINHA 19 QUE É O NECESSÁRIO
    df_dre = df_dre.iloc[:20]

    # ALTERANDO O TIPO PARA FLOAT
    df_dre[mes] = df_dre[mes].astype(float)

    # ALTERANDO TODOS VALORES PARA POSITIVOS
    df_dre[mes] = df_dre[mes].abs()

    # ALTERANDO O NOME DO PRIMEIRO VALOR ICMS QUE APARECE NA COLUNA DESCRIÇÃO
    for index, row in df_dre.iterrows():
    # Encontra a primeira ocorrência de 'ICMS' na descrição da linha
        first_icms_index = row['Descrição'].find('ICMS')
        if first_icms_index != -1:  # Se 'ICMS' for encontrado na linha
        # Substitui apenas o primeiro 'ICMS' encontrado por 'ICMS - imposto'
            new_description = row['Descrição'][:first_icms_index] + 'ICMS - imposto' + row['Descrição'][first_icms_index + len('ICMS'):]
            df_dre.at[index, 'Descrição'] = new_description
            break

    # CRIANDO VARIAVEL DO FATURAMENTO DO MES SELECIONADO
    mes_receita_bruta = df_dre.loc[df_dre['Descrição'] == 'RECEITA BRUTA', mes].values[0]

    # CRIANDO VARIAVEL TOTAL DEVOLUÇÕES DO MES SELECIONADO
    devolucoes_mes = df_dre.loc[df_dre['Descrição'] == 'Devolucoes', mes].values[0]

    # CRIANDO VARIAVEL ICMS TOTAL DO MES SELECIONADO
    icms_mes = df_dre.loc[df_dre['Descrição'] == 'ICMS - imposto', mes].values[0]

    # CRIANDO VARIAVEL ESTORNO ICMS DO MES SELECIONADO
    estorno_icms_mes = df_dre.loc[df_dre['Descrição'] == 'ICMS', mes].values[0]

    return mes_receita_bruta, devolucoes_mes, icms_mes, estorno_icms_mes


# FUNÇÃO APURAR ICMS
def apurar_icms(planilhacte, planilhaicms, planilhadre, cnpj, mes):

    # UNPACKING DOS RETURNS DA FUNÇÃO TRATAR ICMS CTE
    total_icmscte_pelo_cnpj, total_frete_pelo_cnpj = tratar_cte_icms(planilhacte, cnpj)

    # UNPACKING DOS RETURNS DA FUNÇÃO TRATAR ICMS 
    total_bruto_pelo_cnpj, total_icms_pelo_cnpj, total_baseicms_pelo_cnpj, tabela_dinamica_icms = tratar_icms(planilhaicms, cnpj)

    # UNPACKING DOS RETURNS DA FUNÇÃO TRATAR DRE
    mes_receita, devolucoes_mes, icms_mes, estorno_icms_mes = tratar_dre_icms(planilhadre, mes)

    #Iniciando apuração

    total_creditado_bruto = total_bruto_pelo_cnpj + total_frete_pelo_cnpj + devolucoes_mes 

    total_creditado_baseicms = total_baseicms_pelo_cnpj + total_frete_pelo_cnpj + devolucoes_mes 

    total_creditado_valoricms = total_icms_pelo_cnpj + total_icmscte_pelo_cnpj + estorno_icms_mes

    resultado_total_bruto = total_creditado_bruto - mes_receita

    resultado_total_baseicms = total_creditado_baseicms - mes_receita

    resultado_total_icms = total_creditado_valoricms - icms_mes


    #CRIANDO O DATAFRAME CREDITOS PARA RESULTADOS
    # Definindo os nomes das colunas
    colunas = ['VALOR TOTAL BRUTO', 'BASE ICMS', 'VALOR ICMS']

    # DEFININDO OS VALORES DA TABELA
    coluna_total_bruto = [total_bruto_pelo_cnpj, total_frete_pelo_cnpj, devolucoes_mes, total_creditado_bruto]
    coluna_base_icms = [total_baseicms_pelo_cnpj, total_frete_pelo_cnpj, devolucoes_mes, total_creditado_baseicms]
    coluna_valor_icms = [total_icms_pelo_cnpj, total_icmscte_pelo_cnpj, estorno_icms_mes, total_creditado_valoricms]

    #CRIANDO O DATAFRAME CREDITOS PARA RESULTADOS
    creditos = {
    'VALOR TOTAL BRUTO': coluna_total_bruto,
    'BASE ICMS': coluna_base_icms,
    'VALOR ICMS': coluna_valor_icms
    }

    tabela_creditos = pd.DataFrame(creditos, columns=colunas, index= ['Crédito: Entradas', 'Crédito: CTEs', 'Crédito: Devoluções', 'Total Creditado'])
    tabela_creditos = tabela_creditos.applymap(formatar_moeda)

    #CRIANDO O DATAFRAME DEBITOS PARA RESULTADOS
    debitos = {
    'VALOR TOTAL BRUTO': mes_receita,
    'BASE ICMS': mes_receita,
    'VALOR ICMS': icms_mes
    }

    tabela_debitos = pd.DataFrame(debitos, columns= ['VALOR TOTAL BRUTO', 'BASE ICMS', 'VALOR ICMS'], index= ['Débito: Nfe venda'])
    tabela_debitos = tabela_debitos.applymap(formatar_moeda)

    #CRIANDO O DATAFRAME CREDITOS - DEBITOS
    resultado = {
    'VALOR TOTAL BRUTO': resultado_total_bruto,
    'BASE ICMS': resultado_total_baseicms,
    'VALOR ICMS': resultado_total_icms
    }

    tabela_resultados = pd.DataFrame(resultado, columns= ['VALOR TOTAL BRUTO', 'BASE ICMS', 'VALOR ICMS'], index= ['Saldo Final'])
    tabela_resultados = tabela_resultados.applymap(formatar_moeda)

    #CRIANDO O DATAFRAME DETALHES CTE
    detalhes_cte = {
    'VALOR TOTAL - CTEs': total_frete_pelo_cnpj,
    'VALOR ICMS - CTEs': total_icmscte_pelo_cnpj,
    }


    tabela_detalhes_cte = pd.DataFrame(detalhes_cte, columns= ['VALOR TOTAL - CTEs', 'VALOR ICMS - CTEs'], index= [cnpj])
    tabela_detalhes_cte = tabela_detalhes_cte.applymap(formatar_moeda)

    #CRIANDO O DATAFRAME DETALHES DRE
    detalhes_dre = {
    'VALOR TOTAL': [devolucoes_mes, mes_receita],
    'VALOR ICMS': [estorno_icms_mes, icms_mes],
    }


    tabela_detalhes_dre = pd.DataFrame(detalhes_dre, columns= ['VALOR TOTAL', 'VALOR ICMS'], index= ['Crédito: Devoluções', 'Débito: Faturamento' ])
    tabela_detalhes_dre = tabela_detalhes_dre.applymap(formatar_moeda)

    return resultado_total_bruto, resultado_total_baseicms, resultado_total_icms, mes_receita, total_creditado_valoricms, tabela_dinamica_icms, total_bruto_pelo_cnpj, total_icms_pelo_cnpj, total_baseicms_pelo_cnpj, tabela_creditos, tabela_debitos, tabela_resultados, tabela_detalhes_cte, tabela_detalhes_dre


#planilhacte = 'relatorio cte geral abril.xlsx'
#planilhaicms = 'Relatorio apuracao credito icms abril.xlsx'
#planilhadre = 'dre_sp.xlsx'
#cnpj = '33.737.452/0001-00'
#mes= 'Abril'
#apurar_icms(planilhacte, planilhaicms, planilhadre, cnpj, mes)