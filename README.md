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
Que permite o python produzir slidesnas funções:
```
prs = Presentation()
title_only_slide_layout = prs.slide_layouts[5]
slide = prs.slides.add_slide(title_only_slide_layout)
shapes = slide.shapes
```
### biblioteca pptx-util
Que dados os parametros permite o python personalizar o tamanho e posição dos slides nas funções:
```
left = Inches() 
top = Inches()
width = Inches()
height = Inches()
```
---
