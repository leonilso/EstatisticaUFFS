"""
Nome do Arquivo: TB1-01.py
Autor: Leonilso Fandres Wrublak
//Data de Criação: 17/07/0023
Descrição: Cria um arquivo .tex com as respostas do TB1-01 com base em um código informado.

Notas Adicionais:
- Esse código  precisa de uma base de dados para funcionar, ela deve estar em formato .xlsx.
- A questão 1 deve estar no seguinte formato (podendo ter várias linhas):
x1	x2	x3	x4	x5	x6	x7	x8	x9	x10	y1	y2	y3	y4	y5	y6	y7	y8	y9	y10
40	33	38	74	65	46	52	53	34	56	75	90	66	92	34	15	28	34	94	79
...............................................................................
40	33	38	74	65	46	52	53	34	56	75	90	66	92	34	15	28	34	94	79

- A questã 2 deve conter somente números.

- Vale resaltar que para funcionar isto deve ser informado no diretório das linhas 33 e 34.
"""

# Foi usado a biblioteca pandas para fazer a leitura das tabelas
import pandas as pd

# Foi usado a biblioteca pylatex para geração do documento .tex
from pylatex import Document, Section, Command, Alignat, Matrix

# Pegando código do trabalho
codigodigitado = int(input("Informe seu código do TB1-01:"))

# Retirando 1 para garantir o indexador correto nas tabelas
codigo = codigodigitado-1

# Lendo os arquivos 
dados = pd.read_excel('Projetos/projetofaculdade/Trabalho/base_dados/q1.xlsx')
dadosq2 = pd.read_excel('Projetos/projetofaculdade/Trabalho/base_dados/q2.xlsx')

# Iniciando variaveis em 0
menorque40 = 0
entre40e45 = 0
entre45e50 = 0
entre50e55 = 0
entre55e60 = 0
entre60e65 = 0
entre65e70 = 0
maiorque70 = 0


# Vericando valores da questão 2
for i in range(len(dadosq2)):
    for j in range(len(dadosq2.columns)):
        dadosq2.iloc[i, j] = dadosq2.iloc[i, j]+codigodigitado
        if dadosq2.iloc[i, j] < 40:
            menorque40 += 1
        elif dadosq2.iloc[i, j] >= 40 and dadosq2.iloc[i, j] <45:
            entre40e45 += 1
        elif dadosq2.iloc[i, j] >= 45 and dadosq2.iloc[i, j] <50:
            entre45e50 += 1
        elif dadosq2.iloc[i, j] >= 50 and dadosq2.iloc[i, j] <55:
            entre50e55 += 1
        elif dadosq2.iloc[i, j] >= 55 and dadosq2.iloc[i, j] <60:
            entre55e60 += 1
        elif dadosq2.iloc[i, j] >= 60 and dadosq2.iloc[i, j] <65:
            entre60e65 += 1
        elif dadosq2.iloc[i, j] >= 65 and dadosq2.iloc[i, j] <70:
            entre65e70 += 1
        elif dadosq2.iloc[i, j] >= 70:
            maiorque70 += 1

# Verificação do total
totalq2 = menorque40+entre40e45+entre45e50+entre50e55+entre55e60+entre60e65+entre65e70+maiorque70

#Contas de porcentagem
porcent40 = (menorque40/totalq2)*100
porcent40e45 = (entre40e45/totalq2)*100
porcent45e50 = (entre45e50/totalq2)*100
porcent50e55 = (entre50e55/totalq2)*100
porcent55e60 = (entre55e60/totalq2)*100
porcent60e65 = (entre60e65/totalq2)*100
porcent65e70 = (entre65e70/totalq2)*100
porcent70 = (maiorque70/totalq2)*100

# Criação das strings que serão exibidas no .tex
strporcent40 = "A\\, frequência \\, menor \\, que \\, 40\\, \\, absoluta \\, é \\, de " + f"{menorque40} \\,já \\, a \\, relativa \\, é \\, de " + "= \dfrac{" + f"{menorque40}" + "}{" f"{totalq2}" + "} \\times" + "100 = " + f"{porcent40:.2f}\\% \\\\ \\\\ "
strporcent40e45 = "A\\, frequência \\, entre \\, 40 \\, e\\,  45 \\, absoluta \\, é \\, de " + f"{entre40e45} \\,já \\, a \\, relativa \\, é \\, de " + "= \dfrac{" + f"{entre40e45}" + "}{" f"{totalq2}" + "} \\times" + "100 = " + f"{porcent40e45:.2f}\\% \\\\ \\\\ "
strporcent45e50 = "A\\, frequência \\, entre \\, 45 \\, e \\, 50 \\, absoluta \\, é \\, de " + f"{entre45e50} \\,já \\, a \\, relativa \\, é \\, de " + "= \dfrac{" + f"{entre45e50}" + "}{" f"{totalq2}" + "} \\times" + "100 = " + f"{porcent45e50:.2f}\\% \\\\ \\\\ "
strporcent50e55 = "A\\, frequência \\, entre \\, 50 \\, e \\, 55 \\, absoluta \\, é \\, de " + f"{entre50e55} \\,já \\, a \\, relativa \\, é \\, de " + "= \dfrac{" + f"{entre50e55}" + "}{" f"{totalq2}" + "} \\times" + "100 = " + f"{porcent50e55:.2f}\\% \\\\ \\\\ "
strporcent55e60 = "A\\, frequência \\, entre \\, 55 \\, e \\, 60 \\, absoluta \\, é \\, de " + f"{entre55e60} \\,já \\, a \\, relativa \\, é \\, de " + "= \dfrac{" + f"{entre55e60}" + "}{" f"{totalq2}" + "} \\times" + "100 = " + f"{porcent55e60:.2f}\\% \\\\ \\\\ "
strporcent60e65 = "A\\, frequência \\, entre \\, 60 \\, e \\, 65 \\, absoluta \\, é \\, de " + f"{entre60e65} \\,já \\, a \\, relativa \\, é \\, de " + "= \dfrac{" + f"{entre60e65}" + "}{" f"{totalq2}" + "} \\times" + "100 = " + f"{porcent60e65:.2f}\\% \\\\ \\\\ "
strporcent65e70 = "A\\, frequência \\, entre \\, 65 \\, e \\, 70 \\, absoluta \\, é \\, de " + f"{entre65e70} \\,já \\, a \\, relativa \\, é \\, de " + "= \dfrac{" + f"{entre65e70}" + "}{" f"{totalq2}" + "} \\times" + "100 = " + f"{porcent65e70:.2f}\\% \\\\ \\\\ "
strporcent70 = "A\\, frequência \\, maior \\, que \\, 70\\, \\, absoluta \\, é \\, de " + f"{maiorque70} \\,já \\, a \\, relativa \\, é \\, de " + "= \dfrac{" + f"{maiorque70}" + "}{" f"{totalq2}" + "} \\times" + "100 = " + f"{porcent70:.2f}\\% \\\\ \\\\ "

acumulado1 = "&O\\, absoluto \\, acumulado \\, é \\, de \\," + f"{menorque40}" + "\\\\&O \\, relativo \\, acumulado \\, é \\, de\\," + f"{porcent40:.2f} \\% \\\\ \\\\ "
acumulado2 = "&O\\, absoluto \\, acumulado \\, é \\, de \\," + f"{entre40e45 + menorque40}" + "\\\\&O \\, relativo \\, acumulado \\, é \\, de\\," + f"{porcent40e45+ porcent40:.2f} \\% \\\\ \\\\ "
acumulado3 = "&O\\, absoluto \\, acumulado \\, é \\, de \\," + f"{entre45e50 + menorque40 + entre40e45}" + "\\\\&O \\, relativo \\, acumulado \\, é \\, de\\," + f"{porcent45e50+ porcent40 + porcent40e45:.2f} \\% \\\\ \\\\ "
acumulado4 = "&O\\, absoluto \\, acumulado \\, é \\, de \\," + f"{entre50e55 + menorque40 + entre40e45 + entre45e50}" + "\\\\&O \\, relativo \\, acumulado \\, é \\, de\\," + f"{porcent50e55+ porcent40 + porcent40e45 + porcent45e50:.2f} \\% \\\\ \\\\ "
acumulado5 = "&O\\, absoluto \\, acumulado \\, é \\, de \\," + f"{entre55e60 + menorque40 + entre40e45 + entre45e50 + entre50e55}" + "\\\\&O \\, relativo \\, acumulado \\, é \\, de\\," + f"{porcent55e60+ porcent40 + porcent40e45 + porcent45e50+ porcent50e55:.2f} \\% \\\\ \\\\ "
acumulado6 = "&O\\, absoluto \\, acumulado \\, é \\, de \\," + f"{entre60e65 + menorque40 + entre40e45 + entre45e50 + entre50e55 + entre55e60}" + "\\\\&O \\, relativo \\, acumulado \\, é \\, de\\," + f"{ + porcent60e65+ porcent40 + porcent40e45 + porcent45e50+ porcent50e55 + porcent55e60:.2f} \\% \\\\ \\\\ "
acumulado7 = "&O\\, absoluto \\, acumulado \\, é \\, de \\," + f"{entre65e70 + menorque40 + entre40e45 + entre45e50 + entre50e55 + entre55e60 + entre60e65}" + "\\\\&O \\, relativo \\, acumulado \\, é \\, de\\," + f"{porcent65e70+ porcent40 + porcent40e45 + porcent45e50+ porcent50e55 + porcent55e60 + porcent60e65:.2f} \\% \\\\ \\\\ "
acumulado8 = "&O\\, absoluto \\, acumulado \\, é \\, de \\," + f"{maiorque70 + menorque40 + entre40e45 + entre45e50 + entre50e55 + entre55e60 + entre60e65 + entre65e70}" + "\\\\&O \\, relativo \\, acumulado \\, é \\, de\\," + f"{porcent70+ porcent40 + porcent40e45 + porcent45e50+ porcent50e55 + porcent55e60 + porcent60e65 + porcent65e70:.2f} \\% \\\\ \\\\ "


# Iniciando variáveis em 0
G = 0
H = 0
I = 0
J = 0
K = 0
L = 0

# Loop para percorrer os 10 valores
for i in range(10):
    x = dados[f"x{i+1}"][codigo]
    y = dados[f"y{i+1}"][codigo]

    # Parte onde calcula os valores
    somag = x/y
    somah = y/x
    somai = (x/y)**2
    somaj = somah
    somak = ((x**(1/3))/(y**(1/2)))**2
    somal = somah**(1/3)

    # Criação das strings do arquivo .tex
    G_1 = "\dfrac{" + "x_{" + f"{i+1}" "}" + "}{" + "y_{"+f"{i+1}" + "}"+ "}"
    Gtex =  "\dfrac{" + f"{x}" + "}{" + f"{y}" + "}"
    H_1 = "\dfrac{" + "y_{"+f"{i+1}" + "}"+ "}{" + "x_{" + f"{i+1}" "}" + "}"
    Htex = "\dfrac{" + f"{y}" + "}{" + f"{x}" + "}"
    I_1 = "\left(\dfrac{" + "x_{" + f"{i+1}" "}" + "}{" + "y_{"+f"{i+1}" + "}"+ "}\\right)^2"
    Itex = "\left(\dfrac{" + f"{x}" + "}{" + f"{y}" + "}\\right)^2"
    J_1 = H_1
    Jtex = Htex
    K_1 = "\left( \dfrac{" + "\left(x_{" + f"{i+1}" + "}\\right)^{\\frac{" + "1}{"+"3}" + "} }{" + "\left(y_{"+f"{i+1}" + "}\\right)^{\\frac{" + "1}{"+"2}" + "} }\\right)^2"
    Ktex = "\left(\dfrac{" + f"\left({x}\\right)" + "^{\\frac{" + "1}{"+"3}" + "} }{" + f"\left({y}\\right)" + "^{\\frac{" + "1}{"+"2}"+ "} }\\right) ^2"
    L_1 = "\left(\dfrac{" + "y_{" + f"{i+1}" "}" + "}{" + "x_{"+f"{i+1}" + "}"+ "}\\right)^{\\frac{1" + "} {3" + "}"+"}"
    Ltex = "\left(\dfrac{" + f"{y}" + "}{" + f"{x}" + "}\\right)^\\frac{1" + "} {3" + "}"
    
    # teste para verificar se não é o primeiro valor
    if i > 0:
        valorg = valorg + " + "  + Gtex
        valorh = valorh + " + "  + Htex
        valori = valori + " + "  + Itex
        valorj = valorj + " + "  + Jtex
        valork = valork + " + "  + Ktex
        valorl = valorl + " + "  + Ltex
        strg1 = strg1 + " + " + G_1
        strh1 = strh1 + " + " + H_1
        stri1 = stri1 + " + " + I_1
        strj1 = strj1 + " + " + J_1
        strk1 = strk1 + " + " + K_1
        strl1 = strl1 + " + " + L_1
        G_3 = G_3 + " + " + f"{somag:.2f}"
        H_3 = H_3 + " + " + f"{somah:.2f}"
        I_3 = I_3 + " + " + f"{somai:.2f}"
        J_3 = J_3 + " + " + f"{somaj:.2f}"
        K_3 = K_3 + " + " + f"{somak:.2f}"
        L_3 = L_3 + " + " + f"{somal:.2f}"
    # teste para verificar se é o primeiro valor (Assim ele não coloca um sinal de + no ínicio )
    else:
        valorg = Gtex
        valorh = Htex
        valori = Itex
        valorj = Jtex
        valork = Ktex
        valorl = Ltex
        strg1 = G_1
        strh1 = H_1
        stri1 = I_1
        strj1 = J_1
        strk1 = K_1
        strl1 = L_1
        G_3 = f"{somag:.2f}"
        H_3 = f"{somah:.2f}"
        I_3 = f"{somai:.2f}"
        J_3 = f"{somaj:.2f}"
        K_3 = f"{somak:.2f}"
        L_3 = f"{somal:.2f}"
    # Calculo final das questões
    G = G + somag
    H = H + somah
    I = I + somai
    J = J + somaj
    K = K + somak
    L = L + somal
# Elevando ao quadrado as questões que precisam
J = J**2
L = L**2



# Montagem de todas as strings usadas no arquivo .tex
questaog = "G = & \sum_{i=1}^{n=10} = \left[\dfrac{x_i" + "}{y_i" + "}\\right]\\\\  \\\\ "
valorg = "G = &"  + valorg + "\\\\ \\\\ "
strg1 = "G = &" + strg1 + "\\\\ \\\\ "
strg2 = "G = &" + G_3 + "\\\\ \\\\ "
G = "G = &" + f"{G} \\\\ \\\\ "

questaoh = "H = &\sum_{i=1}^{n=10} = \left[\dfrac{y_i" + "}{x_i" + "}\\right]\\\\ "
valorh = "H = &"  + valorh + "\\\\ \\\\ "
strh1 = "H = &" + strh1 + "\\\\ \\\\ "
strh2 = "H = &" + H_3 + "\\\\ \\\\ "
H = "H = &" + f"{H} \\\\ \\\\ "

questaoi = "I = &\sum_{i=1}^{n=10} = \left[\dfrac{x_i" + "}{y_i" + "}\\right]^2\\\\ "
valori = "I = &"  + valori + "\\\\ \\\\ "
stri1 = "I = &" + stri1 + "\\\\ \\\\ "
stri2 = "I = &" + I_3 + "\\\\ \\\\ "
I = "I = &" + f"{I} \\\\ \\\\ "

questaoj = "J = &(\sum_{i=1}^{n=10} = [\dfrac{y_i" + "}{x_i" + "}])^2\\\\ "
valorj = "J = &\left("  + valorj + "\\right)^2\\\\ \\\\ "
strj1 = "J = &\left(" + strj1 + "\\right)^2\\\\ \\\\ "
strj2 = "J = &\left(" + J_3 + "\\right)^2\\\\ \\\\ "
J = "J = &" + f"{J} \\\\ \\\\ "

questaok = "K = &\sum_{i=1}^{n=10} = \left[\dfrac{x_i^\\frac{1}{3}" + "}{y_i^\\frac{1}{2}" + "}\\right]^2\\\\ "
valork = "K = &"  + valork + "\\\\ \\\\ "
strk1 = "K = &" + strk1 + "\\\\ \\\\ "
strk2 = "K = &" + K_3 + "\\\\ \\\\ "
K = "K = &" + f"{K} \\\\ \\\\ "

questaol = "L = &(\sum_{i=1}^{n=10} = [\dfrac{y_i" + "}{x_i" + "}]^\\frac{1" +"}{3" + "})^2\\\\ "
valorl = "L = &\left("  + valorl + "\\right)^2\\\\ \\\\ "
strl1 = "L = &\left(" + strl1 + "\\right)^2\\\\ \\\\ "
strl2 = "L = &\left(" + L_3 + "\\right)^2\\\\ \\\\ "
L = "L = &" + f"{L} \\\\ \\\\ "


#  Criação do documento .tex

doc = Document(geometry_options={"left": "2mm"})
doc.preamble.append(Command('tiny'))

# Adiciona a questão 1

with doc.create(Section(f'Questão 1 código {codigodigitado}')):

    # Usando o Alignat para escrever equações alinhadas

    with doc.create(Alignat(numbering=False, escape=False)) as ag:
        ag.append(questaog)
        ag.append(strg1)
        ag.append(valorg)
        ag.append(strg2)
        ag.append(G)
    with doc.create(Alignat(numbering=False, escape=False)) as ag:
        ag.append(questaoh)
        ag.append(strh1)
        ag.append(valorh)
        ag.append(strh2)
        ag.append(H)
    with doc.create(Alignat(numbering=False, escape=False)) as ag:
        ag.append(questaoi)
        ag.append(stri1)
        ag.append(valori)
        ag.append(stri2)
        ag.append(I)
    with doc.create(Alignat(numbering=False, escape=False)) as ag:
        ag.append(questaoj)
        ag.append(strj1)
        ag.append(valorj)
        ag.append(strj2)
        ag.append(J)
    with doc.create(Alignat(numbering=False, escape=False)) as ag:
        ag.append(questaok)
        ag.append(strk1)
        ag.append(valork)
        ag.append(strk2)
        ag.append(K)
    with doc.create(Alignat(numbering=False, escape=False)) as ag:
        ag.append(questaol)
        ag.append(strl1)
        ag.append(valorl)
        ag.append(strl2)
        ag.append(L)

# Adiciona a questão 2

with doc.create(Section('Questão 2')):

    # Usando o Alignat para escrever equações alinhadas

    with doc.create(Alignat(numbering=False, escape=False)) as ag:
        ag.append(strporcent40)
        ag.append(strporcent40e45)
        ag.append(strporcent45e50)
        ag.append(strporcent50e55)
        ag.append(strporcent55e60)
        ag.append(strporcent60e65)
        ag.append(strporcent65e70)
        ag.append(strporcent70)
    with doc.create(Alignat(numbering=False, escape=False)) as ag:
        ag.append(acumulado1)
        ag.append(acumulado2)
        ag.append(acumulado3)
        ag.append(acumulado4)
        ag.append(acumulado5)
        ag.append(acumulado6)
        ag.append(acumulado7)
        ag.append(acumulado8)


# Salvando o documento em um arquivo .tex
doc.generate_tex('Projetos\projetofaculdade\Trabalho\TB1-01')

# Printa o termino da execução
print(f"terminou de executar o código {codigodigitado} do TB1-01")
