import pandas as pd
from babel.numbers import format_currency

# FUNÇÃO PARA FORMATAR MOEDA BRASILEIRA
def formatar_moeda(valor):
    return format_currency(valor, 'BRL', locale='pt_BR')

# FUNÇÃO TRATAR CTE PARA APURAR PIS E COFINS
def tratar_cte_pis_cofins(planilhacte):
    
    df_cte = pd.read_excel(planilhacte) # AQUI FOI LIDO O ARQUIVO XLSX DE PAGINA UNICA
    colunas_desejadas = ['Tomador', 'CPF/CNPJ Tomador', 'Valor', 'Valor ICMS', 'Status'] # AQUI SELECIONANDO APENAS AS COLUNAS DESEJADAS E CRIANDO UMA COPIA 
    df_cte_colunas_corretas = df_cte[colunas_desejadas].copy() # COPIADO O DF APENAS COM AS COLUNAS DESEJADAS
    
    # FILTRAR O DF APENAS PARA STATUS AUTORIZADA
    df_cte_colunas_corretas = df_cte_colunas_corretas[df_cte_colunas_corretas['Status'] == 'Autorizada']

    # AQUI FORAM ADICIONADAS AS COLUNAS DE PIS E COFINS 
    df_cte_colunas_corretas['PIS'] = (df_cte_colunas_corretas['Valor'] - df_cte_colunas_corretas['Valor ICMS']) * 0.0165
    df_cte_colunas_corretas['COFINS'] = (df_cte_colunas_corretas['Valor'] - df_cte_colunas_corretas['Valor ICMS']) * 0.0760

    # RENOMEANDO PARA VALOR DO FRETE
    df_cte_colunas_corretas = df_cte_colunas_corretas.rename(columns={'Valor': 'Valor do frete'})

    # VARIAVEL TOTAL DE FRETE
    valor_total_cte = round(df_cte_colunas_corretas['Valor do frete'].sum(), 2)
    # VARIAVEL TOTAL DE PIS
    cte_total_pis = round(df_cte_colunas_corretas['PIS'].sum(), 2)
    # VARIAVEL TOTAL DE COFINS
    cte_total_cofins = round(df_cte_colunas_corretas['COFINS'].sum(), 2)


    return valor_total_cte, cte_total_pis, cte_total_cofins


# FUNÇÃO TRATAR PLANILHA ARQUIVEI PARA APURAR PIS E COFINS

def tratar_entrada_pis_cofins(planilha_pis_cofins):

    df_pis_cofins = pd.read_excel(planilha_pis_cofins, sheet_name='relatorio') # AQUI É LIDO O ARQUIVO XLSX COM O NOME DA PAGINA ESPECIFICADO
    
    # FILTRANDO APENAS COLUNAS DESEJADAS
    colunas_desejadas = ['Número', 'Status', 'CNPJ Destinatário', 'Natureza Operação', '[Item] NCM', '[Item] Valor Total Bruto', '[Item] Valor ICMS', '[Item] CST PIS', '[Item] Base de cálculo PIS', 'Valor PIS', '[Item] Base de cálculo COFINS', '[Item] Valor COFINS'] 
    df_pis_cofins = df_pis_cofins[colunas_desejadas].copy() # COPIADO O DF APENAS COM AS COLUNAS DESEJADAS
    
    # FILTRANDO APENAS AS NOTAS AUTORIZADAS NO RELATORIO 
    df_pis_cofins = df_pis_cofins[df_pis_cofins['Status'] == 'Autorizadas']

    # ALTERANDO VALORES NaN PARA 0.0 
    df_pis_cofins.fillna(0, inplace=True)

    # ALTERANDO O TIPO DAS COLUNAS NCM E CST DE FLOAT PARA STRING
    df_pis_cofins['[Item] NCM'] = df_pis_cofins['[Item] NCM'].astype(str)
    #df_pis_cofins['[Item] CST PIS'] = df_pis_cofins['[Item] CST PIS'].astype(str).str.zfill(2)
    df_pis_cofins['[Item] CST PIS'] = df_pis_cofins['[Item] CST PIS'].astype(str)

    # FILTRANDO NCM E CST APENAS ATÉ O DELIMITADOR PONTO
    df_pis_cofins['[Item] NCM'] = df_pis_cofins['[Item] NCM'].str.extract(r'([^.]*)')
    df_pis_cofins['[Item] CST PIS'] = df_pis_cofins['[Item] CST PIS'].str.extract(r'([^.]*)')
    df_pis_cofins['[Item] CST PIS'] = df_pis_cofins['[Item] CST PIS'].str.zfill(2)

    # CONVERTENDO A COLUNA NATUREZA OPERAÇÃO PARA MINUSCULAS
    df_pis_cofins['Natureza Operação'] = df_pis_cofins['Natureza Operação'].str.lower()

    # DEFININDO A BASE DE CALCULO PIS E COFINS PARA NATUREZAS DE VENDA 
    base_calculo = df_pis_cofins['[Item] Valor Total Bruto'] - df_pis_cofins['[Item] Valor ICMS']
    
    # NATUREZAS DE OPERAÇÃO VENDA
    naturezas = ['venda', 'vda']

    # Criar uma expressão regular para verificar a presença de qualquer string na lista naturezas
    pattern = '|'.join(naturezas)

    # Alterar valores nas colunas 'BASE PIS' e 'BASE COFINS' com base na presença das strings em 'descricao'
    df_pis_cofins.loc[df_pis_cofins['Natureza Operação'].str.contains(pattern), ['[Item] Base de cálculo PIS', '[Item] Base de cálculo COFINS']] = base_calculo

    # LISTA NCM MONOFASICAS PARA ZERAR BASES PIS E COFINS E ALTERAR CST PARA 04 
    ncm_monofasicas = ['85122019', '85122021', '83023000', '85122011', '85122022', '85122029', '85123000', '85129000', '85272100', '85443000', '87081000', '87082914', '87082919', '87082992', '87082999', '87089990', '90328929', '84148019', '84244100', '85272900', '85011019']

    # ALTERANDO AS COLUNAS ABAIXO CONFORME A TABELA DE NCM MONOFASICAS
    df_pis_cofins.loc[df_pis_cofins['[Item] NCM'].isin(ncm_monofasicas), '[Item] CST PIS'] = '04'
    df_pis_cofins.loc[df_pis_cofins['[Item] NCM'].isin(ncm_monofasicas), '[Item] Base de cálculo PIS'] = 0.0
    df_pis_cofins.loc[df_pis_cofins['[Item] NCM'].isin(ncm_monofasicas), '[Item] Base de cálculo COFINS'] = 0.0
    df_pis_cofins.loc[df_pis_cofins['[Item] NCM'].isin(ncm_monofasicas), 'Valor PIS'] = 0.0
    df_pis_cofins.loc[df_pis_cofins['[Item] NCM'].isin(ncm_monofasicas), '[Item] Valor COFINS'] = 0.0

    # ALTERANDO VALORES DE PIS E COFINS DE ACORDO COM AS NOVAS BASES
    df_pis_cofins['Valor PIS'] = df_pis_cofins['[Item] Base de cálculo PIS'] * 0.0165
    df_pis_cofins['[Item] Valor COFINS'] = df_pis_cofins['[Item] Base de cálculo COFINS'] * 0.0760

    total_bruto_consolidado = round(df_pis_cofins['[Item] Valor Total Bruto'].sum(), 2)
    total_pis = round(df_pis_cofins['Valor PIS'].sum(), 2)
    total_cofins = round(df_pis_cofins['[Item] Valor COFINS'].sum(), 2)

    # CRIANDO A TABELA DINAMICA
    tabela_dinamica = pd.pivot_table(df_pis_cofins, values=['[Item] Valor Total Bruto', 'Valor PIS', '[Item] Valor COFINS'], index='[Item] CST PIS', aggfunc='sum')
    # CRIANDO OS TOTAIS NA ULTIMA LINHA DA TABELA DINAMICA
    total_por_coluna = tabela_dinamica.sum().to_frame().T
    total_por_coluna.index = ['Total']

    # Adicionando uma nova linha ao DataFrame da tabela dinâmica com os totais
    tabela_dinamica = pd.concat([tabela_dinamica, total_por_coluna])
    tabela_dinamica.columns = ['PIS', 'COFINS', ' VALOR TOTAL BRUTO']
    tabela_dinamica = tabela_dinamica.map('{:.2f}'.format)
    tabela_dinamica = tabela_dinamica.applymap(formatar_moeda)
    #df_teste = df_pis_cofins[df_pis_cofins['Natureza Operação'] == 'retorno de mercadoria para conserto']
    #df_teste = pd.DataFrame(df_teste)

    return tabela_dinamica, total_bruto_consolidado, total_pis, total_cofins

#planilha_pis_cofins = r'Relatorio apuracao PIS e COFINS abril.xlsx'
#tratar_entrada_pis_cofins(planilha_pis_cofins)

# FUNÇÃO PARA TROCAR O PRIMEIRO NOME ENCONTRADO
def trocar_nome(df, nome):
    for index, row in df.iterrows():
    # Encontra a primeira ocorrência de nome na descrição da linha
        first_nome_index = row['Descrição'].find(nome)
        if first_nome_index != -1:  # Se nome for encontrado na linha
            # Substitui apenas o primeiro nome encontrado por 'nome - imposto'
            new_description = row['Descrição'][:first_nome_index] + nome + ' - imposto' + row['Descrição'][first_nome_index + len(nome):]
            df.at[index, 'Descrição'] = new_description
            break


# FUNÇÃO TRATAR DRE PARA APURAR PIS E COFINS
def tratar_dre_pis_cofins(planilha, mes):

    dre_pis_cofins = pd.read_excel(planilha) # AQUI É LIDO O ARQUIVO XLSX DE PAGINA UNICA 

    # FILTRANDO APENAS O MES SELECIONADO
    dre_pis_cofins = dre_pis_cofins[['Descrição', mes]]

    # RETIRANDO ESPAÇOS ANTES DAS LETRAS NA COLUNA DESCRIÇÃO
    dre_pis_cofins['Descrição'] = dre_pis_cofins['Descrição'].str.strip()

    # AQUI FOI RECORTADO O DF ATÉ A LINHA 23 QUE É O NECESSÁRIO
    dre_pis_cofins = dre_pis_cofins.iloc[:24]

    # ALTERANDO OS VALORES PARA FLOAT
    dre_pis_cofins[mes] = dre_pis_cofins[mes].astype(float)

    # ALTERANDO TODOS VALORES PARA POSITIVOS
    dre_pis_cofins[mes] = dre_pis_cofins[mes].abs()

    # ALTERANDO O NOME DO PRIMEIRO VALOR PIS QUE APARECE NA COLUNA DESCRIÇÃO
    trocar_nome(dre_pis_cofins, 'PIS')

    # ALTERANDO O NOME DO PRIMEIRO VALOR COFINS QUE APARECE NA COLUNA DESCRIÇÃO
    trocar_nome(dre_pis_cofins, 'COFINS')

    # CRIANDO VARIAVEL DO FATURAMENTO DO MES SELECIONADO
    mes_receita_bruta = dre_pis_cofins.loc[dre_pis_cofins['Descrição'] == 'RECEITA BRUTA', mes].values[0]

    # CRIANDO VARIAVEL TOTAL DEVOLUÇÕES DO MES SELECIONADO
    devolucoes_mes = dre_pis_cofins.loc[dre_pis_cofins['Descrição'] == 'Devolucoes', mes].values[0]

    # CRIANDO VARIAVEL ICMS TOTAL DO MES SELECIONADO
    pis_mes = dre_pis_cofins.loc[dre_pis_cofins['Descrição'] == 'PIS - imposto', mes].values[0]

    # CRIANDO VARIAVEL ICMS TOTAL DO MES SELECIONADO
    cofins_mes = dre_pis_cofins.loc[dre_pis_cofins['Descrição'] == 'COFINS - imposto', mes].values[0]

    # CRIANDO VARIAVEL ESTORNO ICMS DO MES SELECIONADO
    estorno_pis_mes = dre_pis_cofins.loc[dre_pis_cofins['Descrição'] == 'PIS', mes].values[0]

    # CRIANDO VARIAVEL ESTORNO ICMS DO MES SELECIONADO
    estorno_cofins_mes = dre_pis_cofins.loc[dre_pis_cofins['Descrição'] == 'COFINS', mes].values[0]

    return mes_receita_bruta, devolucoes_mes, pis_mes, cofins_mes, estorno_pis_mes, estorno_cofins_mes

#planilha = 'dre_geral.xlsx'
##mes = 'Abril'
#tratar_dre_pis_cofins(planilha, mes)


def apurar_pis(planilhacte, planilha_pis_cofins, planilha, mes):

    valor_total_cte, cte_total_pis, cte_total_cofins = tratar_cte_pis_cofins(planilhacte)
    tabela_dinamica, total_bruto_consolidado, total_pis, total_cofins = tratar_entrada_pis_cofins(planilha_pis_cofins)
    mes_receita_bruta, devolucoes_mes, pis_mes, cofins_mes, estorno_pis_mes, estorno_cofins_mes = tratar_dre_pis_cofins(planilha, mes)

    total_creditado_bruto = total_bruto_consolidado + valor_total_cte + devolucoes_mes
    total_creditado_pis = total_pis + cte_total_pis + estorno_pis_mes 
    total_creditado_cofins = total_cofins + cte_total_cofins + estorno_cofins_mes

    saldo_final_pis = total_creditado_pis - pis_mes
    saldo_final_cofins = total_creditado_cofins - cofins_mes
    saldo_final_bruto = total_creditado_bruto - mes_receita_bruta

    #CRIANDO O DATAFRAME CREDITOS PARA RESULTADOS
    # Definindo os nomes das colunas
    colunas = ['PIS', 'COFINS', 'VALOR TOTAL BRUTO']

    # DEFININDO OS VALORES DA TABELA
    coluna_total_bruto = [valor_total_cte, devolucoes_mes, total_creditado_bruto]
    coluna_pis = [cte_total_pis, estorno_pis_mes, total_creditado_pis]
    coluna_cofins = [cte_total_cofins, estorno_cofins_mes, total_creditado_cofins]


    #CRIANDO O DATAFRAME CREDITOS PARA RESULTADOS
    creditos_pis = { 
    'PIS': coluna_pis,
    'COFINS': coluna_cofins,
    'VALOR TOTAL BRUTO': coluna_total_bruto
    }

    tabela_pis_creditos = pd.DataFrame(creditos_pis, columns=colunas, index= ['Crédito: CTEs', 'Crédito: Devoluções', 'Total Creditado'])
    tabela_pis_creditos = tabela_pis_creditos.applymap(formatar_moeda)

    #CRIANDO O DATAFRAME DEBITOS PARA RESULTADOS
    debito_pis = {
    'PIS': pis_mes,
    'COFINS': cofins_mes,
    'VALOR TOTAL BRUTO': mes_receita_bruta
    }

    tabela_debito_pis = pd.DataFrame(debito_pis, columns= ['PIS', 'COFINS', 'VALOR TOTAL BRUTO'], index= ['Débito: Nfe venda'])
    tabela_debito_pis = tabela_debito_pis.applymap(formatar_moeda)

    #CRIANDO O DATAFRAME SALDO FINAL
    saldo_final = {
    'PIS': saldo_final_pis,
    'COFINS': saldo_final_cofins,
    'VALOR TOTAL BRUTO': saldo_final_bruto
    }

    tabela_saldo_final = pd.DataFrame(saldo_final, columns= ['PIS', 'COFINS', 'VALOR TOTAL BRUTO'], index= ['Saldo Final'])
    tabela_saldo_final = tabela_saldo_final.applymap(formatar_moeda)



    return valor_total_cte, cte_total_pis, cte_total_cofins, tabela_dinamica, tabela_pis_creditos, tabela_debito_pis, tabela_saldo_final, saldo_final_pis, saldo_final_cofins


#planilhacte = 'relatorio cte geral abril.xlsx'
#planilha_pis_cofins = 'Relatorio apuracao PIS e COFINS abril.xlsx'
#planilha = 'dre_geral.xlsx'
#mes = 'Abril'

#apurar_pis(planilhacte, planilha_pis_cofins, planilha, mes)






















