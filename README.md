# Processador de NFS-e IBS/CBS

Aplicacao local em Streamlit para processar arquivos ZIP com XMLs de NFS-e Nacional, extrair dados fiscais, comparar com uma base em Excel e exportar o resultado em `.xlsx`.

## Funcionalidades

- Upload de um ou mais arquivos `.zip` com XMLs.
- Leitura de XML com namespace.
- Extracao de dados de emitente, tomador, servico, ISSQN, IBS e CBS.
- Upload opcional de base de conferencia em Excel.
- Comparacao por empresa, servico, NBS, CST, cClassTrib, indicador de operacao e reducoes.
- Exportacao do resultado em Excel com abas separadas.

## Como Rodar Localmente

```bash
pip install -r requirements.txt
streamlit run app.py
```

Depois acesse o endereco exibido no terminal, normalmente:

```text
http://localhost:8501
```

## Deploy no Streamlit Cloud

1. Crie um repositorio no GitHub.
2. Envie estes arquivos para o repositorio:
   - `app.py`
   - `requirements.txt`
   - `.gitignore`
   - `.streamlit/config.toml`
   - `README.md`
3. Acesse `https://streamlit.io/cloud`.
4. Clique em `New app`.
5. Escolha o repositorio, branch `main` e arquivo principal `app.py`.
6. Clique em `Deploy`.

## Observacao de Seguranca

Nao suba XMLs, ZIPs ou bases reais com dados fiscais para o GitHub. Esses arquivos devem ser enviados apenas pela tela do app.

## Personalizacao

Para alterar nomes das colunas exportadas no Excel, edite o dicionario `NOMES_COLUNAS_EXCEL` no `app.py`.

Para adicionar novas tags XML, edite:

- `TAGS_MAP`
- `COLUNAS_ORDEM`
