# CONAB Scraper

Este script automatiza o download das tabelas de dados mais recentes das safras de grãos, café e cana-de-açúcar do site da CONAB (Companhia Nacional de Abastecimento).

## Requisitos

- Python 3.8 ou superior
- Google Chrome instalado

## Instalação

1. Clone este repositório ou baixe os arquivos
2. Instale as dependências:

```bash
pip install -r requirements.txt
```

## Uso

Execute o script principal:

```bash
python conab_scraper.py
```

O script irá:
1. Acessar o site da CONAB
2. Navegar até as páginas de safras de grãos, café e cana-de-açúcar
3. Localizar o levantamento mais recente de cada safra
4. Baixar as tabelas de dados correspondentes
5. Salvar os arquivos em pastas específicas para cada tipo de safra:
   - `safra_graos/`: Para tabelas de dados de grãos
   - `safra_cafe/`: Para tabelas de dados de café
   - `safra_cana/`: Para tabelas de dados de cana-de-açúcar

## Estrutura do Projeto

- `conab_scraper.py`: Script principal
- `requirements.txt`: Dependências do projeto
- `safra_graos/`: Diretório onde os arquivos de safra de grãos serão salvos
- `safra_cafe/`: Diretório onde os arquivos de safra de café serão salvos
- `safra_cana/`: Diretório onde os arquivos de safra de cana-de-açúcar serão salvos
- `conab_scraper.log`: Arquivo de log gerado durante a execução

## Funcionalidades

- Download automático das tabelas de dados mais recentes
- Armazenamento em pastas distintas para cada tipo de safra
- Tratamento de erros e logging detalhado
- Execução em modo headless (sem interface gráfica)
- Verificação automática de download bem-sucedido

## Personalização

Você pode modificar as seguintes variáveis no início do script:

- `BASE_URL`: URL base do site da CONAB
- `DOWNLOAD_DIRS`: Dicionário com os diretórios para salvar os arquivos de cada tipo de safra