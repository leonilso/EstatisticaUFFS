
"""
Nome do Arquivo: TB2-03.py
Autor: Leonilso Fandres Wrublak
//Data de Criação: 17/07/0023
Descrição: Cria um arquivo .tex com as respostas do TB2-03 com base em um código informado.

Notas Adicionais:
- Esse código funciona sem quaisquer alterações.
"""


# Foi usado a biblioteca pylatex para geração do documento .tex
from pylatex import Document, Section, Command, Alignat, Subsection

# Foi usado a biblioteca scipy.stats para usar a distribuição normal
from scipy.stats import norm

# Guarda o código do trabalho
codigodigitado = int(input("Informe seu código:"))


# Soma aos valores informados
mediaalemaes = codigodigitado + 46
desviopadraoalemaes = codigodigitado + 20
mediachineses = codigodigitado + 41
desviopadraochineses = codigodigitado + 15

x11 = codigodigitado + 36
x12 = codigodigitado + 40
x13 = codigodigitado + 50
x14 = codigodigitado + 57


# Calcula o Z

Alemaesz11 =(x11-mediaalemaes)/desviopadraoalemaes
Alemaesz12 =(x12-mediaalemaes)/desviopadraoalemaes
Alemaesz13 =(x13-mediaalemaes)/desviopadraoalemaes
Alemaesz14 =(x14-mediaalemaes)/desviopadraoalemaes

chinesesz11 = (x11-mediachineses)/desviopadraochineses
chinesesz12 = (x12-mediachineses)/desviopadraochineses
chinesesz13 = (x13-mediachineses)/desviopadraochineses
chinesesz14 = (x14-mediachineses)/desviopadraochineses


# Cria as strings do .tex

Alemaesz11str = "ZA11 &= \dfrac{" + f"{x11}-{mediaalemaes}" + "}{" + f"{desviopadraoalemaes}" + "} = " + f"{(x11-mediaalemaes)/desviopadraoalemaes:.4f}\\\\"
Alemaesz12str = "ZA12 &= \dfrac{" + f"{x12}-{mediaalemaes}" + "}{" + f"{desviopadraoalemaes}" + "} = " + f"{(x12-mediaalemaes)/desviopadraoalemaes:.4f}\\\\"
Alemaesz13str = "ZA13 &= \dfrac{" + f"{x13}-{mediaalemaes}" + "}{" + f"{desviopadraoalemaes}" + "} = " + f"{(x13-mediaalemaes)/desviopadraoalemaes:.4f}\\\\"
Alemaesz14str = "ZA14 &= \dfrac{" + f"{x14}-{mediaalemaes}" + "}{" + f"{desviopadraoalemaes}" + "} = " + f"{(x14-mediaalemaes)/desviopadraoalemaes:.4f}\\\\"

chinesesz11str = "ZC11 &= \dfrac{" + f"{x11}-{mediachineses}" + "}{" + f"{desviopadraochineses}" + "} = " + f"{(x11-mediachineses)/desviopadraochineses:.4f}\\\\"
chinesesz12str = "ZC12 &= \dfrac{" + f"{x12}-{mediachineses}" + "}{" + f"{desviopadraochineses}" + "} = " + f"{(x12-mediachineses)/desviopadraochineses:.4f}\\\\"
chinesesz13str = "ZC13 &= \dfrac{" + f"{x13}-{mediachineses}" + "}{" + f"{desviopadraochineses}" + "} = " + f"{(x13-mediachineses)/desviopadraochineses:.4f}\\\\"
chinesesz14str = "ZC14 &= \dfrac{" + f"{x14}-{mediachineses}" + "}{" + f"{desviopadraochineses}" + "} = " + f"{(x14-mediachineses)/desviopadraochineses:.4f}\\\\"


# Calcula as probalidades dos alemaes e cria as strings
prob_alemaes_A1 = f"{(norm.cdf(x11, mediaalemaes, desviopadraoalemaes)):.4f}"
prob_alemaes_A2 = f"{(norm.cdf(x12, mediaalemaes, desviopadraoalemaes) - norm.cdf(x11, mediaalemaes, desviopadraoalemaes)):.4f}"
prob_alemaes_A3 = f"{(norm.cdf(x13, mediaalemaes, desviopadraoalemaes) - norm.cdf(x12, mediaalemaes, desviopadraoalemaes)):.4f}"
prob_alemaes_A4 = f"{(norm.cdf(x14, mediaalemaes, desviopadraoalemaes) - norm.cdf(x13, mediaalemaes, desviopadraoalemaes)):.4f}"
prob_alemaes_A5 = f"{(1 - norm.cdf(x14, mediaalemaes, desviopadraoalemaes)):.4f}"

A1alemaes = "A1 &= prob(x1 \leq X11) =  " + prob_alemaes_A1 + "\\\\"
A2alemaes = "A2 &= prob(X11 \leq x1 \leq X12) =  " + f"{norm.cdf(x12, mediaalemaes, desviopadraoalemaes):.4f} - {norm.cdf(x11, mediaalemaes, desviopadraoalemaes):.4f} = " + prob_alemaes_A2 + "\\\\"
A3alemaes = "A3 &= prob(X12 \leq x1 \leq X13) = " + f"{norm.cdf(x13, mediaalemaes, desviopadraoalemaes):.4f} - {norm.cdf(x12, mediaalemaes, desviopadraoalemaes):.4f} = " + prob_alemaes_A3 + "\\\\"
A4alemaes = "A4 &= prob(X13 \leq x1 \leq X14) = " + f"{norm.cdf(x14, mediaalemaes, desviopadraoalemaes):.4f} - {norm.cdf(x13, mediaalemaes, desviopadraoalemaes):.4f} = " + prob_alemaes_A4 + "\\\\"
A5alemaes = "A5 &= prob(x1 \geq X14) = " + f"1 - {norm.cdf(x14, mediaalemaes, desviopadraoalemaes):.4f} = " + prob_alemaes_A5


# Calcula as probalidades dos chineses e cria as strings
prob_chineses_A1 = f"{(norm.cdf(x11, mediachineses, desviopadraochineses)):.4f}"
prob_chineses_A2 = f"{(norm.cdf(x12, mediachineses, desviopadraochineses) - norm.cdf(x11, mediachineses, desviopadraochineses)):.4f}"
prob_chineses_A3 = f"{(norm.cdf(x13, mediachineses, desviopadraochineses) - norm.cdf(x12, mediachineses, desviopadraochineses)):.4f}"
prob_chineses_A4 = f"{(norm.cdf(x14, mediachineses, desviopadraochineses) - norm.cdf(x13, mediachineses, desviopadraochineses)):.4f}"
prob_chineses_A5 = f"{(1 - (norm.cdf(x14, mediachineses, desviopadraochineses))):.4f}"

A1chineses = "A1 &= prob(x1 \leq X11) =  " + prob_chineses_A1 + "\\\\"
A2chineses = "A2 &= prob(X11 \leq x1 \leq X12) =  " + f"{norm.cdf(x12, mediachineses, desviopadraochineses):.4f} - {norm.cdf(x11, mediachineses, desviopadraochineses):.4f} = " + prob_chineses_A2 + "\\\\"
A3chineses = "A3 &= prob(X12 \leq x1 \leq X13) = " + f"{norm.cdf(x13, mediachineses, desviopadraochineses):.4f} - {norm.cdf(x12, mediachineses, desviopadraochineses):.4f} = " + prob_chineses_A3 + "\\\\"
A4chineses = "A4 &= prob(X13 \leq x1 \leq X14) = " + f"{norm.cdf(x14, mediachineses, desviopadraochineses):.4f} - {norm.cdf(x13, mediachineses, desviopadraochineses):.4f} = " + prob_chineses_A4 + "\\\\"
A5chineses = "A5 &= prob(x1 \geq X14) = " + f"1 - {norm.cdf(x14, mediachineses, desviopadraochineses):.4f} = " + prob_chineses_A5 + "\\\\"

# Cria um documento .tex
doc = Document(geometry_options={"left": "2mm"})
doc.preamble.append(Command('tiny'))

with doc.create(Section(f'Calculos código {codigodigitado}')):
    with doc.create(Subsection('Alemães')) as sb:
        with sb.create(Alignat(numbering=False, escape=False)) as ag:
            ag.append(Alemaesz11str)
            ag.append(Alemaesz12str)
            ag.append(Alemaesz13str)
            ag.append(Alemaesz14str)
            ag.append("\\\\ Na Tabela  \\\\\ ")
            ag.append(A1alemaes)
            ag.append(A2alemaes)
            ag.append(A3alemaes)
            ag.append(A4alemaes)
            ag.append(A5alemaes)

    with doc.create(Subsection('Chineses')) as sb:
        with sb.create(Alignat(numbering=False, escape=False)) as ag:
            ag.append(chinesesz11str)
            ag.append(chinesesz12str)
            ag.append(chinesesz13str)
            ag.append(chinesesz14str)
            ag.append("\\\\ Na Tabela  \\\\\ ")
            ag.append(A1chineses)
            ag.append(A2chineses)
            ag.append(A3chineses)
            ag.append(A4chineses)
            ag.append(A5chineses)


# Gera o documento .tex
doc.generate_tex('Projetos\projetofaculdade\Trabalho\TB2-03')

# Printa o termino da execução
print(f"terminou de executar o código {codigodigitado} do TB2-03")