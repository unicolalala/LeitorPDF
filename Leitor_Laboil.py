import tabula
import pandas as pd
import re
import os
import openpyxl
from glob import glob

diretorio_atual = os.path.dirname(os.path.abspath(__file__))
caminho_analises = os.path.join(diretorio_atual, "Analises/Relatorios")
arquivo_excel = os.path.join(diretorio_atual, "Analises/Excel/Oleott_2.xlsx")
numero_documento = os.listdir(caminho_analises)

for documento in numero_documento:
    caminho_arquivo_pdf = os.path.join(caminho_analises, documento)  # Complete file path

    tables_2 = tabula.read_pdf(caminho_arquivo_pdf, pages='all', guess=False, multiple_tables=True, lattice= True,)

        # Verificar se há pelo menos uma tabela encontrada

    if len(tables_2) >= 1:
            tabela2 = tables_2[0]

            dados_df = pd.DataFrame(tables_2[0])

    #variaveis        
    Num_serie = dados_df.iloc[5, 1]
    data_coleta = dados_df.iloc[6, 3]

    # Verifica se é cr ou fq
    if dados_df.iloc[16, 0] == "Hidrogênio":
        
    #Começa a pegar os dados

        dados_excel_CR = pd.read_excel("Analises/Excel/Oleott_2.xlsx", sheet_name="BD_CR")

        Hidrogenio  = dados_df.iloc[16, 4] #H2
        Oxigenio = dados_df.iloc[17, 4] #O2 
        Nitrogenio = dados_df.iloc[18, 4] #N2
        Dioxido_carbono = dados_df.iloc[19, 4] #CO2
        Etileno = dados_df.iloc[20, 4] #C2H4
        Etano = dados_df.iloc[21, 4] #C2H6
        if dados_df.iloc[22, 4] == "ND":
            Acetileno = "0" #C2H2
        else:
            Acetileno == dados_df.iloc[22, 4]
            
        Monoxido_carbono = dados_df.iloc[24, 4] #CO
        Metano = dados_df.iloc[23, 4] #CH4
        Total_combustivel = dados_df.iloc[25, 4] #CGC


        laboil_cr = pd.DataFrame({
                                    'Numero de Série': [Num_serie],
                                    'H2': [Hidrogenio],
                                    'O2': [Oxigenio],
                                    'N2': [Nitrogenio],
                                    'CO2': [Dioxido_carbono],
                                    'C2H4': [Etileno],
                                    'C2H6': [Etano],
                                    'C2H2': [Acetileno],
                                    'CO': [Monoxido_carbono],
                                    'CH4': [Metano],
                                    'TG': ["0"],  # Você precisa fornecer um valor aqui ou remover essa coluna
                                    'CGC': [Total_combustivel],
                                    'Diagnostico': ["0"],  # Você precisa fornecer um valor aqui ou remover essa coluna
                                })
        dados_excel_CR = dados_excel_CR.append(laboil_cr, ignore_index=True)

        with pd.ExcelWriter(arquivo_excel, mode='a', if_sheet_exists='replace') as writer:
            dados_excel_CR.to_excel(writer, sheet_name= 'BD_CR', index=False)
    else:
        # Pega os dados de FQ
        print(Num_serie)


        dados_excel_fq = pd.read_excel("Analises/Excel/Oleott_2.xlsx", sheet_name="BD_FQ")

        Aspecto  = dados_df.iloc[15, 3] 
        cor = dados_df.iloc[16, 3] #O2 
        neutraliza = dados_df.iloc[17, 3] #N2
        dieletric = dados_df.iloc[18, 3] #CO2
        interfac = dados_df.iloc[19, 3] #C2H4
        Fp = dados_df.iloc[20, 3] #C2H6
        teor_h2o = dados_df.iloc[21, 3] #CO
        Densid = dados_df.iloc[22, 3] #CH4

        laboil_fq = pd.DataFrame({  
                                    'Numero de Série' : [Num_serie],
                                    'Aspecto Visual' : [Aspecto],
                                    'Cor': [cor],
                                    'Índice de Neutralização': [neutraliza],
                                    'Rigidez Dielétrica': [dieletric],
                                    'Tensão Interfacial': [interfac],
                                    'Fator de Perdas 90 °C': [Fp],
                                    'Teor de Água': [teor_h2o],
                                    'Densidade 20/4 °C': [Densid],
                                    })

        dados_excel_FQ = dados_excel_fq.append(laboil_fq, ignore_index=True)


        with pd.ExcelWriter(arquivo_excel, mode='a', if_sheet_exists='replace') as writer:
                dados_excel_FQ.to_excel(writer, sheet_name= 'BD_FQ', index=False)
