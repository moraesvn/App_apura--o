import streamlit as st
import pandas as pd
from funcoes_apurar_icms import apurar_icms, tratar_dre_icms
from funcoes_apurar_pis import formatar_moeda
from funcoes_apurar_pis import apurar_pis


st.set_page_config(layout="wide")

def main():
    # Define o título da barra lateral
    st.sidebar.title("Apuração de Impostos")

    # Adiciona as opções na barra lateral
    option = st.sidebar.radio("Selecione a opção desejada:", ["ICMS", "PIS/COFINS"])
    
    # Exibe o conteúdo com base na opção selecionada
    if option == "ICMS":
        st.write('# Apuração - ICMS')

        col1, col2 = st.columns(2)
        with col1:
            # SELECIONA O CNPJ
            cnpj = st.selectbox("SELECIONE O CNPJ:", ["33.737.452/0001-00", "33.737.452/0003-71"])
        with col2:
            # SELECIONA O MÊS
            mes = st.selectbox("SELECIONE O MÊS CORRESPONDENTE:", ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"])
        if cnpj == '33.737.452/0001-00':
            st.markdown(':red[**GP MATRIZ - SÃO PAULO**]')
        if cnpj == '33.737.452/0003-71':
            st.markdown(':red[**GP FILIAL - SANTA CATARINA**]')


        st.markdown('##### Importação de arquivos para apuração:')
        col1, col2, col3 = st.columns(3)
        # IMPORTA A PLANILHA DE CTE
        with col1:
            planilhacte = st.file_uploader("PLANILHA ARQUIVEI - CTEs: " + mes, key= 'importe cte')

        # IMPORTA A PLANILHA DE ICMS
        with col2:
            planilhaicms = st.file_uploader("PLANILHA ARQUIVEI ENTRADAS - ICMS: " + mes, key= 'importe icms')

        # IMPORTA A PLANILHA DRE
        if cnpj == '33.737.452/0001-00':
            with col3:
                planilhadre = st.file_uploader("IMPORTE O DRE DA :red[**GP SP - Matriz: .xlsx**]", key= 'importe dre')
                if planilhadre:
                   try:
                       tratar_dre_icms(planilhadre, mes)
                   except:
                       st.markdown('###### :red[Ops, arquivo .xls não suportado, importe o DRe em formato .xlsx.]')    
        if cnpj == '33.737.452/0003-71':
            with col3:
                planilhadre = st.file_uploader("IMPORTE O DRE DA :red[**GP SC - Filial: .xlsx**]", key= 'importe dre')
                if planilhadre:
                    try:
                        tratar_dre_icms(planilhadre, mes)
                    except:
                        st.markdown('###### :red[Ops, arquivo .xls não suportado, importe o DRe em formato .xlsx.]')



        # APURAÇÃO DE SP 
        if cnpj == "33.737.452/0001-00" and planilhacte and planilhaicms and planilhadre:
            # COLUNAS PARA RESULTADOS
            col1, col2 = st.columns(2)
            with col1:
                if 'apurar' not in st.session_state:
                    st.session_state.apurar = False
                if st.button('Apurar'):
                    st.session_state.apurar = True

                if st.session_state.apurar:
                    chamafuncao = apurar_icms(planilhacte, planilhaicms, planilhadre, cnpj, mes)
                    (st.session_state.resultado_total_bruto, st.session_state.resultado_total_baseicms,
                    st.session_state.resultado_total_icms, st.session_state.mes_receita,
                    st.session_state.total_creditado_valoricms, st.session_state.tabela_dinamica_icms,
                    st.session_state.total_bruto_pelo_cnpj, st.session_state.total_icms_pelo_cnpj,
                    st.session_state.total_baseicms_pelo_cnpj, st.session_state.tabela_creditos,
                    st.session_state.tabela_debitos, st.session_state.tabela_resultados,
                    st.session_state.tabela_detalhes_cte, st.session_state.tabela_detalhes_dre) = chamafuncao


                    # EXIBINDO OS RESULTADOS
                    st.markdown('#### Resumo - Débito e Crédito') 
                    st.table(st.session_state.tabela_creditos)
                    st.table(st.session_state.tabela_debitos)
                    st.table(st.session_state.tabela_resultados)
                    if st.session_state.resultado_total_icms > 0:
                        st.metric(label=f':red[Total ICMS Creditado em {mes}]', value=round(st.session_state.resultado_total_icms, 2))
                    if st.session_state.resultado_total_icms < 0:
                        st.metric(label=f':red[Total ICMS Debitado em {mes}]', value=round(st.session_state.resultado_total_icms, 2))
                    valor_credito = st.number_input(label=':red[Insira o valor de crédito ICMS acumulado do mês anterior.]')
                    if valor_credito:
                        credito_total = valor_credito + st.session_state.resultado_total_icms
                        st.metric(label=':red[Crédito total acumulado atualizado]', value=f"{credito_total:.2f}")

                    with col2: 
                        detalhar_checkbox = st.checkbox('Mostrar detalhes')                
                        if detalhar_checkbox:
                            st.markdown('#### Entradas por CFOP') 
                            st.table(st.session_state.tabela_dinamica_icms)
                            st.markdown('#### Dados de CTEs') 
                            st.table(st.session_state.tabela_detalhes_cte)
                            st.markdown('#### Dados do DRe - ' + cnpj) 
                            st.table(st.session_state.tabela_detalhes_dre)
                
        





        
        # APURAÇÃO DE SC
        if cnpj == "33.737.452/0003-71" and planilhacte and planilhaicms and planilhadre:

            col1, col2 = st.columns(2)
            with col1:

                # INSERIR OS VALORES FATURADOS POR ESTADO E PERNTUAL DE ICMS 
                valor_17 = st.number_input('Insira o valor faturado dentro do estado - ICMS 17%')
                valor_4 = st.number_input('Insira o valor faturado fora do estado - ICMS 4%')
            with col2:  
                valor_7 = st.number_input('Insira o valor faturado fora do estado - ICMS 7%')
                valor_12 = st.number_input('Insira o valor faturado fora do estado - ICMS 12%')

            if valor_17 and valor_4 and valor_7 and valor_12:

                # CRIANDO VARIAVEIS DE VALORES DOS ICMS ACIMA
                icms_17 = valor_17 * 0.17
                icms_4 = valor_4 * 0.04
                icms_7 = valor_7 * 0.07
                icms_12 = valor_12 * 0.12

                # CRIANDO VARIAVEIS DOS ABATIMENTOS
                abatimento_4 = icms_4 * 0.75
                abatimento_7 = icms_7 * 0.7143
                abatimento_12 = icms_12 * 0.8333
                total_abatimento = abatimento_4 + abatimento_7 + abatimento_12

                # VALORES ICMS APÓS ABATIMENTOS
                resultado_icms_4 = icms_4 - abatimento_4
                resultado_icms_7 = icms_7 - abatimento_7
                resultado_icms_12 = icms_12 - abatimento_12

                # SOMA DO VALOR FATURADO FORA DO ESTADO PARA CALCULAR O FUNDO SOCIAL
                total_fora_estado = valor_4 + valor_7 + valor_12
                
                # COLUNAS PARA RESULTADOS
                col1, col2 = st.columns(2)
                with col1:
                    if 'apurar' not in st.session_state:
                        st.session_state.apurar = False
                    if st.button('Apurar'):
                        st.session_state.apurar = True
                    if st.session_state.apurar:
                        chamafuncao = apurar_icms(planilhacte, planilhaicms, planilhadre, cnpj, mes)
                        (st.session_state.resultado_total_bruto, st.session_state.resultado_total_baseicms,
                        st.session_state.resultado_total_icms, st.session_state.mes_receita,
                        st.session_state.total_creditado_valoricms, st.session_state.tabela_dinamica_icms,
                        st.session_state.total_bruto_pelo_cnpj, st.session_state.total_icms_pelo_cnpj,
                        st.session_state.total_baseicms_pelo_cnpj, st.session_state.tabela_creditos,
                        st.session_state.tabela_debitos, st.session_state.tabela_resultados,
                        st.session_state.tabela_detalhes_cte, st.session_state.tabela_detalhes_dre) = chamafuncao
                        
                        # PERCENTUAL DO ABATIMENTO PARA CALCULO DO ICMS PROPRIO FORA DO REGIME
                        percentual_abatido = valor_17 / st.session_state.mes_receita                    
                        
                        
                        #CRIANDO O DATAFRAME RESULTADOS SC
                        resultado = {
                        'BASE ICMS': [valor_17, valor_4, valor_7, valor_12],
                        'VALOR ICMS': [icms_17, icms_4, icms_7, icms_12],
                        'ABATIMENTO': [0.0, abatimento_4, abatimento_7, abatimento_12],
                        'ICMS - ABATIDO': [0.0, resultado_icms_4, resultado_icms_7, resultado_icms_12]
                        }
                    

                        st.session_state.tabela_sc = pd.DataFrame(resultado, columns= ['BASE ICMS', 'VALOR ICMS', 'ABATIMENTO', 'ICMS - ABATIDO'], index= ['ICMS 17% - SC', 'ICMS 4%', 'ICMS 7%', 'ICMS 12%'])
                        st.session_state.tabela_sc['BASE ICMS'] = st.session_state.tabela_sc['BASE ICMS'].apply(lambda x: f"R$ {x:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
                        st.session_state.tabela_sc['VALOR ICMS'] = st.session_state.tabela_sc['VALOR ICMS'].apply(lambda x: f"R$ {x:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
                        st.session_state.tabela_sc['ABATIMENTO'] = st.session_state.tabela_sc['ABATIMENTO'].apply(lambda x: f"R$ {x:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
                        st.session_state.tabela_sc['ICMS - ABATIDO'] = st.session_state.tabela_sc['ICMS - ABATIDO'].apply(lambda x: f"R$ {x:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
            
                        
                        # EXIBINDO OS RESULTADOS
                        st.markdown('#### Resumo - Débito e Crédito')
                        st.table(st.session_state.tabela_creditos)
                        st.table(st.session_state.tabela_debitos)
                        st.table(st.session_state.tabela_resultados)
                        st.markdown('#### Apuração Regime Especial')
                        st.table(st.session_state.tabela_sc)
                    
                        with col2: 
                            col1, col2 = st.columns(2)
                            with col1:
                                st.markdown('')
                                st.markdown('')
                                st.markdown('')
                                st.markdown('')
                                st.markdown('#### Valores das Guias de pagamentos')

                                # VALORES DAS GUIAS
                                resultado_icms_regime = resultado_icms_4 + resultado_icms_7 + resultado_icms_12
                                resultado_icms_regimef = formatar_moeda(resultado_icms_regime)
                                fundes = total_abatimento * 0.02
                                fundesf = formatar_moeda(fundes)
                                fundo_social = (total_fora_estado * 0.004) - fundes
                                fundo_socialf = formatar_moeda(fundo_social)                        
                                icms_fora_regime = icms_17 - (st.session_state.total_creditado_valoricms * percentual_abatido)
                                icms_fora_regimef = formatar_moeda(icms_fora_regime)

                                totalsob_faturamento = (resultado_icms_regime + fundes + fundo_social + icms_fora_regime) / st.session_state.mes_receita

                                st.metric(label=':red[ICMS PRÓPRIO APURADO NO REGIME]', value=resultado_icms_regimef)
                                st.metric(label=':red[FUNDES]', value=fundesf)
                                st.metric(label=':red[FUNDO SOCIAL]', value=fundo_socialf)
                                st.metric(label=':red[ICMS APURADO FORA DO REGIME]', value=icms_fora_regimef)

                            with col2:
                                st.markdown('')
                                st.markdown('')
                                st.markdown('')
                                st.markdown('')
                                st.markdown('#### Outros dados')
                                st.metric(label=':red[% ABATIDO PARA ICMS FORA DO REGIME]', value=f'{percentual_abatido:.2%}')
                                st.metric(label=':red[TOTAL DAS GUIAS / FATURAMENTO]', value=f'{totalsob_faturamento:.2%}')






    if option == "PIS/COFINS":
        st.write('# Apuração - PIS/COFINS.')
            

        col1, col2 = st.columns(2)
        with col1:
            mes = st.selectbox("SELECIONE O MÊS CORRESPONDENTE:", ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"])


        st.markdown('##### Importação de arquivos para apuração:')
        col1, col2, col3 = st.columns(3)
        with col1:
            planilhactes = st.file_uploader("PLANILHA ARQUIVEI - CTEs: " + mes, key= 'importe ctes')

        # IMPORTA A PLANILHA PIS COFINS
        with col2:
            planilhapis = st.file_uploader("PLANILHA ARQUIVEI - PIS/COFINS: " + mes, key= 'importe pis')
        
        with col3:
            planilha_dre_pis = st.file_uploader("IMPORTE O :red[**DRE GERAL DO GRUPO: .xlsx**] ", key= 'importe dregeral')


        if planilhactes and planilhapis and planilha_dre_pis:
            apurar = st.button('Apurar')
            if apurar:
                funcao = apurar_pis(planilhactes, planilhapis, planilha_dre_pis, mes)
                valor_total_frete, valor_total_pis, valor_total_cofins, tabela_dinamica, tabela_pis_creditos, tabela_debito_pis, tabela_saldo_final, saldo_final_pis, saldo_final_cofins = funcao
                col1, col2 = st.columns(2)
                with col1:
                    st.markdown('#### Resumo - Débito e Crédito') 
                    st.table(tabela_dinamica)
                    st.table(tabela_pis_creditos)
                    st.table(tabela_debito_pis)
                    st.table(tabela_saldo_final)
                with col2:
                    st.markdown('#### Valores das guias de pagamentos')
                    st.metric(label=':red[PIS]', value=formatar_moeda(saldo_final_pis))
                    st.metric(label=':red[COFINS]', value=formatar_moeda(saldo_final_cofins))





if __name__ == "__main__":
    main()