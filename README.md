# csv_table_pptx
---
- Gerador de slides que utiliza informações de um arquivo csv para um slide com tabela para melhor exibição das informações.  Nesse codigo é utilizado duas bibliotecas de python:
---
### biblioteca csv
Que consegue ler arquivos csv nas funções:
```
with open('table.csv', mode='r', encoding='utf-8') as archive:
    reader = csv.DictReader(archive)
```
### biblioteca pptx
Que permite o python produzir slides nas funções:
```
prs = Presentation()
title_only_slide_layout = prs.slide_layouts[5]
slide = prs.slides.add_slide(title_only_slide_layout)
shapes = slide.shapes
```
### biblioteca pptx-util
Dados os parametros permite o python personalizar o tamanho e posição dos slides nas funções:
```
left = Inches() 
top = Inches()
width = Inches()
height = Inches()
```
---
# Código
## Importando bibliotecas
```
from pptx import Presentation
from pptx.util import Inches
import csv
```

## Inicializando os vetores
- OBS: Para a aplicação no meu dia a dia eu limitei a quantidade de colunas em 5 por isso para produzir um slide com uma tabela com mais/menos de 5 colunas é necessário adicionar/remover variaveis aqui 
```
datas = []
horarios = []
dias = []
celebracoes = []
operadores = []
```
## Estabelecendo parâmetros do slide
- Escolha o layout do slide, e escolho o titulo que será escrito. (o layout é o parâmetro que decide o a posição do titulo)
```
#Configuração da formação dos slides
prs = Presentation()
title_only_slide_layout = prs.slide_layouts[5]
slide = prs.slides.add_slide(title_only_slide_layout)
shapes = slide.shapes
shapes.title.text = 'Escala Equipe de Som \n Matriz São Pedro'
```
- Escolha o tamanho da matriz que será utilizada para fazer a tabela as variáveis linhas e colunas sendo a quantidade de linhas e colunas da matriz
- As variáveis left,top,width e height são a posição da tabela no slide. 
```
linhas = len(datas) + 1  # +1 para o cabeçalho
colunas = 5
left = Inches(0.2) 
top = Inches(2.0)
width = Inches(9.0)
height = Inches(2.0)
```

## Abra e leia o arquivo CSV
- Leitura do arquivo csv
```
with open('escala.csv', mode='r', encoding='utf-8') as arquivo:
    leitor = csv.DictReader(arquivo)
```
- Separação das colunas em vetores para melhor manipulação desses dados
- Aqui eu separo as colunas do csv e atribuo cada em um vetores
```
    for linha in leitor:
        datas.append(linha['Data'])
        horarios.append(linha['Horário'])
        dias.append(linha['Dia'])
        celebracoes.append(linha['Celebração'])
        operadores.append(linha['Operador'])
        linhas+=1
```
## Criação da tabela no slide
```
table = shapes.add_table(linhas, colunas, left, top, width, height).table
```
- Escolha a largura de cada uma das colunas
```
# set column widths
table.columns[0].width = Inches(1.0)
table.columns[1].width = Inches(1.2)
table.columns[2].width = Inches(2.5)
table.columns[3].width = Inches(2.5)
table.columns[4].width = Inches(2.5)
```
- Escolha o nome de cada uma das colunas
```
# write column headings
table.cell(0, 0).text = 'Data'
table.cell(0, 1).text = 'Horário'
table.cell(0, 2).text = 'Dia'
table.cell(0, 3).text = 'Celebração'
table.cell(0, 4).text = 'Operador'
```
- Atribuição das informações dos vetores em cada célula da tabela
```
i=0
# write body cells
for i in range(len(datas)):
    table.cell(i+1, 0).text = datas[i]
    table.cell(i+1, 1).text = horarios[i]
    table.cell(i+1, 2).text = dias[i]
    table.cell(i+1, 3).text = celebracoes[i]
    table.cell(i+1, 4).text = operadores[i]
```
## Criação do Slide
```
prs.save('Escala.pptx')
```

