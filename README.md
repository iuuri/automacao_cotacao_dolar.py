# Automação de Extração de Cotação do Dólar

Este projeto automatiza a extração da cotação atual do dólar a partir do site da Wise, tira um print da página, registra a data e hora da execução, gera um documento `.docx` com essas informações e converte o arquivo para PDF.

## Tecnologias Utilizadas
- **Python**
- **Selenium**: Para automação de navegação na web e extração da cotação do dólar.
- **python-docx**: Para geração de arquivos `.docx`.
- **docx2pdf**: Para conversão de arquivos `.docx` em PDF.
- **datetime**: Para registrar a data e hora da extração.

## Funcionalidades
1. Acessa o site da Wise e extrai a cotação atual do dólar.
2. Tira um print da tela do site mostrando a cotação.
3. Registra a data e hora do momento da extração.
4. Cria um documento `.docx` com a cotação, a data e a hora, e anexa o print da página.
5. Converte o arquivo `.docx` para PDF.