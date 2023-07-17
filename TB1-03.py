"""
Nome do Arquivo: TB1-03.py
Autor: Leonilso Fandres Wrublak
//Data de Criação: 17/07/0023
Descrição: Cria um arquivo .tex com as respostas do TB1-03 com base em um código informado e uma equipe.

Notas Adicionais:
- Esse código  precisa de uma base de dados para funcionar, ela deve estar em formato .xlsx e estar no seguinte formato:
	    janeiro	fevereiro	março	abril	maio
Fernanda	2,4	    6,4	    2,2	    1,2	    6,6
José	    2,4	    6,7	    4,2	    4,9	    3,6
Paula	    5,1	    7,2	    4,8	    4	    7,5
Roberta	    6,6	    1,7	    6	    9,8	    1,7
Vale resaltar que para isso deve ser informado o diretório nas linhas 35, 36, 37 e 38

"""

# Foi usado a biblioteca pandas para fazer a leitura das tabelas
import pandas as pd

# Foi usado a biblioteca pylatex para geração do documento .tex
from pylatex import Document, Section, Command, Alignat, Subsection


# Pegando código do trabalho
codigodigitado = int(input("Informe seu código: "))

# Pegando a equipe
equipe = input("Informe a sua equipe (a/b): ")




# Lendo as tabelas
PV1_A_E1 = pd.read_excel('Projetos/projetofaculdade/base_dados/PV1-A-E1.xlsx')
PV1_A_E2 = pd.read_excel('Projetos/projetofaculdade/base_dados/PV1-A-E2.xlsx')
PV1_B_E1 = pd.read_excel('Projetos/projetofaculdade/base_dados/PV1-B-E1.xlsx')
PV1_B_E2 = pd.read_excel('Projetos/projetofaculdade/base_dados/PV1-B-E2.xlsx')


# Adicionando o código
PV1_A_E1[(PV1_A_E1.select_dtypes(include='number').columns)] += codigodigitado
PV1_A_E2[(PV1_A_E1.select_dtypes(include='number').columns)] += codigodigitado
PV1_B_E1[(PV1_A_E1.select_dtypes(include='number').columns)] += codigodigitado
PV1_B_E2[(PV1_A_E1.select_dtypes(include='number').columns)] += codigodigitado


# Iniciando varáveis vazias
mediaPV1_A = []
mediaPV1_B = []
desviopadraoPV1_A = []
desviopadraoPV1_B = []
varcoefPV1_A = []
varcoefPV1_B = []

PV1_A_E1strings = []
PV1_A_E1stringsdp = []
PV1_A_E1stringscf = []

# Calculando e criando as strings do modelo a equipe 1
for index, row in PV1_A_E1.iterrows():
    name = row['Unnamed: 0']
    values = row[1:]  
    average = sum(values) / len(values)
    desviopadrao = values.std()
    varcoef = desviopadrao/average
    formatted_values = [f"({value} - {average:.6f})^2" for column, value in zip(values.index, values)]
    PV1_A_E1string = f"& Média \\, de \\, {name}: \\\\ & \dfrac" + "{" + f"({' + '.join(map(str, values))}"+ "}{" + f"{len(values)}" + "}"+  f"= {average:.6f}\\\\ "
    PV1_A_E1stringdp = f"& Desvio \\, padrão\\, {name}: \\\\ " + "& \sqrt{ \dfrac" + "{" + f"({' + '.join(map(str, formatted_values))}"+ "}{" + f"({len(values)} - 1)" + "}}"+  f"= {desviopadrao:.6f}\\\\ "
    PV1_A_E1stringcf = f"& Coeficiente \\, de \\, variação \\, {name}: \\\\ & \dfrac" + "{" + f"({desviopadrao:.6f})"+ "}{" + f"({average})" + "}"+  f"= {varcoef:.6f}\\\\ "
    PV1_A_E1strings.append(PV1_A_E1string)
    PV1_A_E1stringsdp.append(PV1_A_E1stringdp)
    PV1_A_E1stringscf.append(PV1_A_E1stringcf)
PV1_A_E1stringfinal = '\\\\'.join(PV1_A_E1strings) + "\\\\ \\\\ "
PV1_A_E1stringfinaldp = '\\\\'.join(PV1_A_E1stringsdp) + "\\\\ \\\\ "
PV1_A_E1stringfinalcf = '\\\\'.join(PV1_A_E1stringscf) + "\\\\ \\\\ "
values = PV1_A_E1.iloc[:, 1:].values.flatten()
average = values.mean()
desviopadrao = values.std(ddof=1)
varcoef = desviopadrao / average
mediaPV1_A.append(average)
desviopadraoPV1_A.append(desviopadrao)
varcoefPV1_A.append(varcoef)
PV1_A_E1stringgeral = f"Média Geral: {average:.6f}"
PV1_A_E1stringdpgeral = f"Desvio Padrão Geral: {desviopadrao:.6f}"
PV1_A_E1stringcfgeral = f"Coeficiente de Variação Geral: {varcoef:.6f}"

PV1_A_E2strings = []
PV1_A_E2stringsdp = []
PV1_A_E2stringscf = []

# Calculando e criando as strings do modelo a equipe 2
for index, row in PV1_A_E2.iterrows():
    name = row['Unnamed: 0']
    values = row[1:] 
    average = sum(values) / len(values)
    desviopadrao = values.std()
    varcoef = desviopadrao/average
    formatted_values = [f"({value} - {average:.6f})^2" for column, value in zip(values.index, values)]
    PV1_A_E2string = f"& Média \\,de \\,{name}: \\\\ &\dfrac" + "{" + f"({' + '.join(map(str, values))}"+ "}{" + f"{len(values)}" + "}"+  f"= {average:.6f}\\\\ "
    PV1_A_E2stringdp = f"& Desvio \\, padrão\\, {name}: \\\\ " + "&\sqrt{ \dfrac" + "{" + f"({' + '.join(map(str, formatted_values))}"+ "}{" + f"({len(values)} - 1)" + "}}"+  f"= {desviopadrao:.6f}\\\\ "
    PV1_A_E2stringcf = f"& Coeficiente \\, de \\, variação \\,{name}: \\\\ &\dfrac" + "{" + f"({desviopadrao:.6f})"+ "}{" + f"({average})" + "}"+  f"= {varcoef:.6f}\\\\ "
    PV1_A_E2strings.append(PV1_A_E2string)
    PV1_A_E2stringsdp.append(PV1_A_E2stringdp)
    PV1_A_E2stringscf.append(PV1_A_E2stringcf)
PV1_A_E2stringfinal = '\\\\'.join(PV1_A_E2strings) + "\\\\ \\\\ "
PV1_A_E2stringfinaldp = '\\\\'.join(PV1_A_E2stringsdp) + "\\\\ \\\\ "
PV1_A_E2stringfinalcf = '\\\\'.join(PV1_A_E2stringscf) + "\\\\ \\\\ "
values = PV1_A_E2.iloc[:, 1:].values.flatten()
average = values.mean()
desviopadrao = values.std(ddof=1)
varcoef = desviopadrao / average
mediaPV1_A.append(average)
desviopadraoPV1_A.append(desviopadrao)
varcoefPV1_A.append(varcoef)
PV1_A_E2stringgeral = f"Média Geral: {average:.6f}"
PV1_A_E2stringdpgeral = f"Desvio Padrão Geral: {desviopadrao:.6f}"
PV1_A_E2stringcfgeral = f"Coeficiente de Variação Geral: {varcoef:.6f}"

PV1_B_E1strings = []
PV1_B_E1stringsdp = []
PV1_B_E1stringscf = []

# Calculando e criando as strings do modelo b equipe 1
for index, row in PV1_B_E1.iterrows():
    name = row['Unnamed: 0']
    values = row[1:] 
    average = sum(values) / len(values)
    desviopadrao = values.std()
    varcoef = desviopadrao/average
    formatted_values = [f"({value} - {average:.6f})^2" for column, value in zip(values.index, values)]
    PV1_B_E1string = f"& Média \\,de \\,{name}: \\\\ &\dfrac" + "{" + f"({' + '.join(map(str, values))}"+ "}{" + f"{len(values)}" + "}"+  f"= {average:.6f}\\\\ "
    PV1_B_E1stringdp = f"& Desvio \\, padrão\\, {name}: \\\\ " + "&\sqrt{ \dfrac" + "{" + f"({' + '.join(map(str, formatted_values))}"+ "}{" + f"({len(values)} - 1)" + "}}"+  f"= {desviopadrao:.6f}\\\\ "
    PV1_B_E1stringcf = f"& Coeficiente \\, de \\, variação \\,{name}: \\\\ &\dfrac" + "{" + f"({desviopadrao:.6f})"+ "}{" + f"({average})" + "}"+  f"= {varcoef:.6f}\\\\ "
    PV1_B_E1strings.append(PV1_B_E1string)
    PV1_B_E1stringsdp.append(PV1_B_E1stringdp)
    PV1_B_E1stringscf.append(PV1_B_E1stringcf)
PV1_B_E1stringfinal = '\\\\'.join(PV1_B_E1strings) + "\\\\ \\\\ "
PV1_B_E1stringfinaldp = '\\\\'.join(PV1_B_E1stringsdp) + "\\\\ \\\\ "
PV1_B_E1stringfinalcf = '\\\\'.join(PV1_B_E1stringscf) + "\\\\ \\\\ "
values = PV1_B_E1.iloc[:, 1:].values.flatten()
average = values.mean()
desviopadrao = values.std()
varcoef = desviopadrao / average
mediaPV1_B.append(average)
desviopadraoPV1_B.append(desviopadrao)
varcoefPV1_B.append(varcoef)
PV1_B_E1stringgeral = f"Média Geral: {average:.6f}"
PV1_B_E1stringdpgeral = f"Desvio Padrão Geral: {desviopadrao:.6f}"
PV1_B_E1stringcfgeral = f"Coeficiente de Variação Geral: {varcoef:.6f}"

PV1_B_E2strings = []
PV1_B_E2stringsdp = []
PV1_B_E2stringscf = []

# Calculando e criando as strings do modelo b equipe 2
for index, row in PV1_B_E2.iterrows():
    name = row['Unnamed: 0']
    values = row[1:]
    average = sum(values) / len(values)
    desviopadrao = values.std()
    varcoef = desviopadrao/average
    formatted_values = [f"({value} - {average:.6f})^2" for column, value in zip(values.index, values)]
    PV1_B_E2string = f"& Média \\,de \\,{name}: \\\\ &\dfrac" + "{" + f"({' + '.join(map(str, values))}"+ "}{" + f"{len(values)}" + "}"+  f"= {average:.6f}\\\\ "
    PV1_B_E2stringdp = f"& Desvio \\, padrão\\, {name}: \\\\ " + "&\sqrt{ \dfrac" + "{" + f"({' + '.join(map(str, formatted_values))}"+ "}{" + f"({len(values)} - 1)" + "}}"+  f"= {desviopadrao:.6f}\\\\ "
    PV1_B_E2stringcf = f"& Coeficiente \\, de \\, variação \\,{name}: \\\\ &\dfrac" + "{" + f"({desviopadrao:.6f})"+ "}{" + f"({average})" + "}"+  f"= {varcoef:.6f}\\\\ "
    PV1_B_E2strings.append(PV1_B_E2string)
    PV1_B_E2stringsdp.append(PV1_B_E2stringdp)
    PV1_B_E2stringscf.append(PV1_B_E2stringcf)
PV1_B_E2stringfinal = '\\\\'.join(PV1_B_E2strings) + "\\\\ \\\\ "
PV1_B_E2stringfinaldp = '\\\\'.join(PV1_B_E2stringsdp) + "\\\\ \\\\ "
PV1_B_E2stringfinalcf = '\\\\'.join(PV1_B_E2stringscf) + "\\\\ \\\\ "
values = PV1_B_E2.iloc[:, 1:].values.flatten()
average = values.mean()
desviopadrao = values.std()
varcoef = desviopadrao / average
mediaPV1_B.append(average)
desviopadraoPV1_B.append(desviopadrao)
varcoefPV1_B.append(varcoef)
PV1_B_E2stringgeral = f"Média Geral: {average:.6f}"
PV1_B_E2stringdpgeral = f"Desvio Padrão Geral: {desviopadrao:.6f}"
PV1_B_E2stringcfgeral = f"Coeficiente de Variação Geral: {varcoef:.6f}"

colunas = ["janeiro", "fevereiro", "março", "abril", "maio"]

# Criação de um dataframe para acessar mais fácil os valores

PV1_A_E1['média'] = PV1_A_E1[colunas].mean(axis=1)
PV1_A_E1['desvio padrão'] = PV1_A_E1[colunas].std(axis=1)
PV1_A_E1['coeficiente de variação'] = PV1_A_E1['desvio padrão'] / PV1_A_E1['média']

PV1_A_E2['média'] = PV1_A_E2[colunas].mean(axis=1)
PV1_A_E2['desvio padrão'] = PV1_A_E2[colunas].std(axis=1)
PV1_A_E2['coeficiente de variação'] = PV1_A_E2['desvio padrão'] / PV1_A_E2['média']

PV1_B_E1['média'] = PV1_B_E1[colunas].mean(axis=1)
PV1_B_E1['desvio padrão'] = PV1_B_E1[colunas].std(axis=1)
PV1_B_E1['coeficiente de variação'] = PV1_B_E1['desvio padrão'] / PV1_B_E1['média']

PV1_B_E2['média'] = PV1_B_E2[colunas].mean(axis=1)
PV1_B_E2['desvio padrão'] = PV1_B_E2[colunas].std(axis=1)
PV1_B_E2['coeficiente de variação'] = PV1_B_E2['desvio padrão'] / PV1_B_E2['média']


PV1AEQUIPE = pd.concat([PV1_A_E1, PV1_A_E2])
PV1BEQUIPE = pd.concat([PV1_B_E1, PV1_B_E2])

PV1AEQUIPE = PV1AEQUIPE.reset_index(drop=True)
PV1BEQUIPE = PV1BEQUIPE.reset_index(drop=True)


# Criação do documento .tex
doc = Document(geometry_options={"left": "2mm"})
doc.preamble.append(Command('tiny'))


# Adicionando os cáculos no documento conforme a equipe
with doc.create(Section(f'Cálculos código {codigodigitado}')):
    if(equipe == "a"):
        with doc.create(Subsection('Modelo PVA')) as sb:
            sb.append("Média equipe 1: \n")
            sb.append(PV1_A_E1stringgeral)
            with sb.create(Alignat(numbering=False, escape=False)) as ag:
                ag.append(PV1_A_E1stringfinal)

            sb.append("Média equipe 2: \n")
            sb.append(PV1_A_E2stringgeral)
            with sb.create(Alignat(numbering=False, escape=False)) as ag:
                ag.append(PV1_A_E2stringfinal)

            sb.append("Desvio Padrão equipe 1:\n")
            sb.append(PV1_A_E1stringdpgeral)
            with sb.create(Alignat(numbering=False, escape=False)) as ag:
                ag.append(PV1_A_E1stringfinaldp)

            sb.append("Desvio Padrão equipe 2:")
            sb.append(PV1_A_E2stringdpgeral)
            with sb.create(Alignat(numbering=False, escape=False)) as ag:
                ag.append(PV1_A_E2stringfinaldp)

            sb.append("Coeficiente de variação equipe 1: \n")
            sb.append(PV1_A_E1stringcfgeral)
            with sb.create(Alignat(numbering=False, escape=False)) as ag:
                ag.append(PV1_A_E1stringfinalcf)

            sb.append("Coeficiente de variação equipe 2: \n")
            sb.append(PV1_A_E2stringcfgeral)
            with sb.create(Alignat(numbering=False, escape=False)) as ag:
                ag.append(PV1_A_E2stringfinalcf)
    elif(equipe == "b"):
        with doc.create(Subsection('Modelo PVB')) as sb:
            sb.append("Média equipe 1: \n")
            sb.append(PV1_B_E1stringgeral)
            with sb.create(Alignat(numbering=False, escape=False)) as ag:
                ag.append(PV1_B_E1stringfinal)

            sb.append("Média equipe 2: \n")
            sb.append(PV1_B_E2stringgeral)
            with sb.create(Alignat(numbering=False, escape=False)) as ag:
                ag.append(PV1_B_E2stringfinal)

            sb.append("Desvio Padrãoe equipe 1:\n")
            sb.append(PV1_B_E1stringdpgeral)
            with sb.create(Alignat(numbering=False, escape=False)) as ag:
                ag.append(PV1_B_E1stringfinaldp)

            sb.append("Desvio Padrão equipe 2:\n")
            sb.append(PV1_B_E2stringdpgeral)
            with sb.create(Alignat(numbering=False, escape=False)) as ag:
                ag.append(PV1_B_E2stringfinaldp)

            sb.append("Coeficiente de variação equipe 1: \n")
            sb.append(PV1_B_E1stringcfgeral)
            with sb.create(Alignat(numbering=False, escape=False)) as ag:
                ag.append(PV1_B_E1stringfinalcf)

            sb.append("Coeficiente de variação equipe 2: \n")
            sb.append(PV1_B_E2stringcfgeral)
            with sb.create(Alignat(numbering=False, escape=False)) as ag:
                ag.append(PV1_B_E2stringfinalcf)


# Geração das respostas

with doc.create(Section('Respostas')):
    if(equipe == "a"):
        with doc.create(Subsection('Modelo PVA')) as pva:
            if mediaPV1_A[0] > mediaPV1_A[1]:
                pva.append(f" A equipe 1 é mais produtiva com média de: {mediaPV1_A[0]:.6f} \n")
            else:
                pva.append(f" A equipe 2 é mais produtiva com média de: {mediaPV1_A[1]:.6f} \n")
            
            if varcoefPV1_A[0] < varcoefPV1_A[1]:
                pva.append(f" A equipe 1 é mais regular com coeficiente de variação: {varcoefPV1_A[0]:.6f} \n")
            else:
                pva.append(f" A equipe 2 é mais regular com coeficiente de variação: {varcoefPV1_A[1]:.6f} \n")

            pva.append(f" O profissional mais produtivo é: {PV1AEQUIPE.loc[PV1AEQUIPE['média'].idxmax(), 'Unnamed: 0']} com {PV1AEQUIPE.loc[PV1AEQUIPE['média'].idxmax(), 'média']:.6f} de média e com {PV1AEQUIPE.loc[PV1AEQUIPE['média'].idxmax(), 'coeficiente de variação']:.6f} de coeficiente de variação\n")
            pva.append(f" O profissional mais regular é: {PV1AEQUIPE.loc[PV1AEQUIPE['coeficiente de variação'].idxmin(), 'Unnamed: 0']} com {PV1AEQUIPE.loc[PV1AEQUIPE['coeficiente de variação'].idxmin(), 'média']:.6f} de média e com {PV1AEQUIPE.loc[PV1AEQUIPE['coeficiente de variação'].idxmin(), 'coeficiente de variação']:.6f} de coeficiente de variação\n")
            pva.append(f" O profissional menos regular é: {PV1AEQUIPE.loc[PV1AEQUIPE['coeficiente de variação'].idxmax(), 'Unnamed: 0']} com {PV1AEQUIPE.loc[PV1AEQUIPE['coeficiente de variação'].idxmax(), 'média']:.6f} de média e com {PV1AEQUIPE.loc[PV1AEQUIPE['coeficiente de variação'].idxmax(), 'coeficiente de variação']:.6f} de coeficiente de variação\n")
            maiores_valoresA = PV1AEQUIPE.nlargest(5, 'média')
            maiores_valoresA = maiores_valoresA.sort_values('coeficiente de variação')
            maiores5A = maiores_valoresA['Unnamed: 0'].tolist()
            maiores5Avalor = maiores_valoresA['média'].tolist()
            pva.append(f" A melhor equipe seria com os profissionais: {maiores5A[0]}, {maiores5A[1]}, {maiores5A[2]}, {maiores5A[3]} com base na maior média e como critério de desempate o coeficiente de variação.\n")
            mediaAequipenova = (maiores5Avalor[0] + maiores5Avalor[1] + maiores5Avalor[2] + maiores5Avalor[3])/ 4
            pva.append(f" A média desta equipe: {maiores5A[0]}, {maiores5A[1]}, {maiores5A[2]}, {maiores5A[3]}. Seria {mediaAequipenova:.6f}, que é a média de cada funcionário dividido pela quantidade de funcionários.\n")
    elif(equipe == "b"): 
        with doc.create(Subsection('Modelo PVB')) as pvb:
            if mediaPV1_B[0] > mediaPV1_B[1]:
                pvb.append(f" A equipe 1 é mais produtiva com média de: {mediaPV1_B[0]:.6f} \n")
            else:
                pvb.append(f" A equipe 2 é mais produtiva com média de: {mediaPV1_B[1]:.6f} \n")
            
            if varcoefPV1_B[0] < varcoefPV1_B[1]:
                pvb.append(f" A equipe 1 é mais regular com coeficiente de variação: {varcoefPV1_B[0]:.6f} \n")
            else:
                pvb.append(f" A equipe 2 é mais regular com coeficiente de variação: {varcoefPV1_B[1]:.6f} \n")

            pvb.append(f" O profissional mais produtivo é: {PV1BEQUIPE.loc[PV1BEQUIPE['média'].idxmax(), 'Unnamed: 0']} com {PV1BEQUIPE.loc[PV1BEQUIPE['média'].idxmax(), 'média']:.6f} de média e com {PV1BEQUIPE.loc[PV1BEQUIPE['média'].idxmax(), 'coeficiente de variação']:.6f} de coeficiente de variação\n")
            pvb.append(f" O profissional mais regular é: {PV1BEQUIPE.loc[PV1BEQUIPE['coeficiente de variação'].idxmin(), 'Unnamed: 0']} com {PV1BEQUIPE.loc[PV1BEQUIPE['coeficiente de variação'].idxmin(), 'média']:.6f} de média e com {PV1BEQUIPE.loc[PV1BEQUIPE['coeficiente de variação'].idxmin(), 'coeficiente de variação']:.6f} de coeficiente de variação\n")
            pvb.append(f" O profissional menos regular é: {PV1BEQUIPE.loc[PV1BEQUIPE['coeficiente de variação'].idxmax(), 'Unnamed: 0']} com {PV1BEQUIPE.loc[PV1BEQUIPE['coeficiente de variação'].idxmax(), 'média']:.6f} de média e com {PV1BEQUIPE.loc[PV1BEQUIPE['coeficiente de variação'].idxmax(), 'coeficiente de variação']:.6f} de coeficiente de variação\n")
            maiores_valoresB = PV1BEQUIPE.nlargest(5, 'média')
            maiores_valoresB = maiores_valoresB.sort_values('coeficiente de variação')
            maiores5B = maiores_valoresB['Unnamed: 0'].tolist()
            maiores5Bvalor = maiores_valoresB['média'].tolist()
            pvb.append(f" A melhor equipe seria com os profissionais: {maiores5B[0]}, {maiores5B[1]}, {maiores5B[2]}, {maiores5B[3]} com base na maior média e como critério de desempate o coeficiente de variação.\n")
            mediaBequipenova = (maiores5Bvalor[0] + maiores5Bvalor[1] + maiores5Bvalor[2] + maiores5Bvalor[3])/ 4
            pvb.append(f" A média desta equipe: {maiores5B[0]}, {maiores5B[1]}, {maiores5B[2]}, {maiores5B[3]}. Seria {mediaBequipenova:.6f}, que é a média de cada funcionário dividido pela quantidade de funcionários.\n")

# Salvando documento .tex
doc.generate_tex('Projetos\projetofaculdade\Trabalho\TB1-03')

# Printa o termino da execução
print(f"terminou de executar o código {codigodigitado} do modelo PV{equipe} do trabalho TB1-03")