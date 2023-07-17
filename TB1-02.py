"""
Nome do Arquivo: TB1-02.py
Autor: Leonilso Fandres Wrublak
//Data de Criação: 17/07/0023
Descrição: Cria um arquivo .tex com as respostas do TB1-02 com base em um código informado.

Notas Adicionais:
- Esse código  precisa de uma base de dados para funcionar, ela deve estar em formato .xlsx e deve conter somente números.
- Vale resaltar que para funcionar isto deve ser informado no diretório das linhas 27 e 28.
"""



# Foi usado a biblioteca pandas para fazer a leitura das tabelas
import pandas as pd

# Foi usado a biblioteca pylatex para geração do documento .tex
from pylatex import Document, Section, Command, Alignat

# Pegando código do trabalho
codigodigitado = int(input("Informe seu código:"))

# Retirando 1 para garantir o indexador correto nas tabelas
codigo = codigodigitado-1

# Lendo os arquivos 
alunas = pd.read_excel('Projetos/projetofaculdade/base_dados/Alunas.xlsx', header=None)
alunos = pd.read_excel('Projetos/projetofaculdade/base_dados/Alunos.xlsx', header=None)

#  localizando código informado
alunascodigo = alunas.iloc[codigo] 
alunoscodigo = alunos.iloc[codigo] 


# Iniciando variáveis em 0
somaaluna = 0
somaaluno = 0
somadiferençasalunas = 0
somadiferençasalunos = 0
diferencaalunaquad = []
diferencaalunoquad = []

for i in range(len(alunascodigo)):
    somaaluna = somaaluna + alunascodigo[i]
    somaaluno = somaaluno + alunoscodigo[i]

# Calculando a média

mediaaluna = somaaluna/len(alunascodigo)
mediaaluno = somaaluno/len(alunoscodigo)

# Calculando a mediana
medianaaluna = alunascodigo.median()
medianaaluno = alunoscodigo.median()

# Calculando a moda
modaaluna = alunascodigo.mode()[0]
modaaluno = alunoscodigo.mode()[0]

# Calculando o valor máximo
maximoaluna = alunascodigo.max()
maximoaluno = alunoscodigo.max()

# Calculando o valor mínimo
minimoaluna = alunascodigo.min()
minimoaluno = alunoscodigo.min()

# Loop para calcular as medidas de dispersão e criação de strings
for i in range(len(alunascodigo)):
    diferenalunas = ((float(alunascodigo[i])-mediaaluna)**2)
    diferenalunos = ((float(alunoscodigo[i])-mediaaluno)**2)
    somadiferençasalunas = somadiferençasalunas + diferenalunas
    somadiferençasalunos = somadiferençasalunos + diferenalunos


    #Isso garante a quebra de linha no arquivo .tex
    if i ==5 or i == 10 or i == 15 or i == 20 or i == 25 or i == 30 or i == 35 or i == 40:
        calcdiferenalunasstr = calcdiferenalunasstr + f"+\\\\ \\left({float(alunascodigo[i]):.2f} - {mediaaluna}\\right)^2"
        calcdiferenalunosstr = calcdiferenalunosstr + f"+\\\\ \\left({float(alunoscodigo[i]):.2f} - {mediaaluno}\\right)^2"
        diferenalunasstr = diferenalunasstr + " +\\\\ " + f"{diferenalunas:.2f}"
        diferenalunosstr = diferenalunosstr + " +\\\\ " + f"{diferenalunos:.2f}"
    else:
        if i > 0:
            diferenalunasstr = diferenalunasstr + " + " + f"{diferenalunas:.2f}"
            diferenalunosstr = diferenalunosstr + " + " + f"{diferenalunos:.2f}"
            calcdiferenalunasstr = calcdiferenalunasstr + f"+ \\left({float(alunascodigo[i]):.2f} - {mediaaluna}\\right)^2"
            calcdiferenalunosstr = calcdiferenalunosstr + f"+ \\left({float(alunoscodigo[i]):.2f} - {mediaaluno}\\right)^2"
        else:
            diferenalunasstr = f"{diferenalunas:.2f}"
            diferenalunosstr = f"{diferenalunos:.2f}"
            calcdiferenalunasstr = f"\\left({float(alunascodigo[i]):.2f} - {mediaaluna}\\right)^" + "{2" + "} "
            calcdiferenalunosstr = f"\\left({float(alunoscodigo[i]):.2f} - {mediaaluno}\\right)^" + "{2" + "} "


# Criação das strings usadas no .tex

calcdiferenalunasstr = "DQT_a = \\\\" + calcdiferenalunasstr + "\\\\ \\\\"
calcdiferenalunosstr = "DQT_o = \\\\" + calcdiferenalunosstr + "\\\\ \\\\"

diferenalunasstr = "DQT_a = \\\\" + diferenalunasstr + f" = \\\\ {somadiferençasalunas:.2f}\\\\ \\\\ "
diferenalunosstr = "DQT_o = \\\\" + diferenalunosstr + f" = \\\\{somadiferençasalunos:.2f}\\\\ \\\\ "


# Cálculo das medidas de dispersão 

DQMA = somadiferençasalunas/len(alunascodigo)
DQMO = somadiferençasalunos/len(alunoscodigo)
varianciaaluna = somadiferençasalunas/(len(alunascodigo)-1)
varianciaaluno = somadiferençasalunos/(len(alunoscodigo)-1)
desviopadraoaluna = varianciaaluna**(1/2)
desviopadraoaluno = varianciaaluno**(1/2)
coeficientevara = desviopadraoaluna/mediaaluna
coeficientevaro = desviopadraoaluno/mediaaluno

# Criação das strings usadas no .tex

DQMASTR = "DQM_a = \\dfrac{" + f"{somadiferençasalunas:.2f}" + "}{"+ f"{(len(alunascodigo))}" + "} =" + f"{DQMA}"
DQMOSTR = "DQM_o = \\dfrac{" + f"{somadiferençasalunos:.2f}" + "}{"+ f"{(len(alunoscodigo))}" + "} =" + f"{DQMO}"

varianciaalunastr = "S^2_a =  \\dfrac{" + f"{somadiferençasalunas:.2f}" + "}{"+ f"{(len(alunascodigo))}" + "-1} =" + f"{varianciaaluna}"
varianciaalunostr = "S^2_a =  \\dfrac{" + f"{somadiferençasalunos:.2f}" + "}{"+ f"{(len(alunoscodigo))}" + "-1} =" + f"{varianciaaluno}"

desviopadraoalunastr = "S_a = \\sqrt{" + f"{varianciaaluna:.2f}" + "}=" + f"{desviopadraoaluna}"
desviopadraoalunostr = "S_o = \\sqrt{" + f"{varianciaaluno:.2f}" + "}=" + f"{desviopadraoaluno}" 

coeficientevarastr = "CV_a = \dfrac{" + f"{desviopadraoaluna}" + "}{" + f"{mediaaluna}" + "} = " + f"{coeficientevara}"
coeficientevarostr = "CV_o = \dfrac{" + f"{desviopadraoaluno}" + "}{" + f"{mediaaluno}" + "} = " + f"{coeficientevaro}"

mediaalunastr = "Média \\, alunas = \\dfrac{" +f"{somaaluna}" + "}{" + f"{len(alunascodigo)} " + "} =" + f"{mediaaluna}"
mediaalunostr = "Média \\,alunos = \\dfrac{" +f"{somaaluno}" + "}{" + f"{len(alunoscodigo)} " + "} =" + f"{mediaaluno}"

medianaalunastr = "Mediana \\, alunas = " + f"{medianaaluna}"
medianaalunostr = "Mediana \\,alunos = " + f"{medianaaluno}"

modaalunastr = "Moda \\, alunas = " + f"{modaaluna}"
modaalunostr = "Moda \\,alunos = " + f"{modaaluno}"

amplitudealunastr = "Amplitude \\, alunas = " + f"{(maximoaluna-minimoaluna)}"
amplitudealunostr = "Amplitude \\,alunos = " + f"{(maximoaluno-minimoaluno)}"

# Verificando quem é maior em média

if mediaaluna > mediaaluno:
    difemedia= "As \\, alunas \\, são: \dfrac{" + f"{mediaaluna} " + "}{" + f"{mediaaluno}  " + "}=" + f"{((((mediaaluna)/mediaaluno)-1)*100):.2f}" + "\\% \\, maiores \\, que \\, os \\, alunos."
else:
    difemedia= "Os \\, alunos \\, são: \dfrac{" + f"{mediaaluno} " + "}{" + f"{mediaaluna}  " + "}=" + f"{((((mediaaluno)/mediaaluna)-1)*100):.2f}" + "\\% \\, maiores \\, que \\, as \\, alunas."

# Criação do documento .tex

doc = Document(geometry_options={"left": "2mm"})
doc.preamble.append(Command('tiny'))

# Adiciona as strings das alunas

with doc.create(Section(f'Alunas {codigodigitado}')):
    with doc.create(Alignat(numbering=False, escape=False)) as ag:
        ag.append(mediaalunastr)
    with doc.create(Alignat(numbering=False, escape=False)) as ag:
        ag.append(medianaalunastr)
    with doc.create(Alignat(numbering=False, escape=False)) as ag:
        ag.append(modaalunastr)
    with doc.create(Alignat(numbering=False, escape=False)) as ag:
        ag.append(amplitudealunastr)
    with doc.create(Alignat(numbering=False, escape=False)) as ag:
        ag.append(difemedia)

    with doc.create(Alignat(numbering=False, escape=False)) as ag:
        ag.append(calcdiferenalunasstr)
    with doc.create(Alignat(numbering=False, escape=False)) as ag:
        ag.append(diferenalunasstr)
    with doc.create(Alignat(numbering=False, escape=False)) as ag:
        ag.append(DQMASTR)
    with doc.create(Alignat(numbering=False, escape=False)) as ag:
        ag.append(varianciaalunastr)
    with doc.create(Alignat(numbering=False, escape=False)) as ag:
        ag.append(desviopadraoalunastr)
    with doc.create(Alignat(numbering=False, escape=False)) as ag:
        ag.append(coeficientevarastr)

# Adiciona as strings dos alunos

with doc.create(Section('Alunos')):
    with doc.create(Alignat(numbering=False, escape=False)) as ag:
        ag.append(mediaalunostr)
    with doc.create(Alignat(numbering=False, escape=False)) as ag:
        ag.append(medianaalunostr)
    with doc.create(Alignat(numbering=False, escape=False)) as ag:
        ag.append(modaalunostr)
    with doc.create(Alignat(numbering=False, escape=False)) as ag:
        ag.append(amplitudealunostr)
    with doc.create(Alignat(numbering=False, escape=False)) as ag:
        ag.append(difemedia)

    with doc.create(Alignat(numbering=False, escape=False)) as ag:
        ag.append(calcdiferenalunosstr)
    with doc.create(Alignat(numbering=False, escape=False)) as ag:
        ag.append(diferenalunosstr)
    with doc.create(Alignat(numbering=False, escape=False)) as ag:
        ag.append(DQMOSTR)
    with doc.create(Alignat(numbering=False, escape=False)) as ag:
        ag.append(varianciaalunostr)
    with doc.create(Alignat(numbering=False, escape=False)) as ag:
        ag.append(desviopadraoalunostr)
    with doc.create(Alignat(numbering=False, escape=False)) as ag:
        ag.append(coeficientevarostr)


# Gera o documento .tex
doc.generate_tex('Projetos\projetofaculdade\Trabalho\TB1-02')

# Printa o termino da execução
print(f"terminou de executar o código {codigodigitado} do TB1-02")