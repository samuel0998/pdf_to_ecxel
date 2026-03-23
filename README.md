# PDF → Excel | Relação de Pedidos

Aplicação web Flask para converter o PDF "Relação de Pedidos - Ítens" em planilha Excel formatada.

## Estrutura

```
pdf_to_excel/
├── app.py               # Backend Flask + parser PDF + gerador Excel
├── requirements.txt
├── templates/
│   └── index.html       # Interface web
├── static/
│   ├── css/style.css
│   └── js/app.js
├── uploads/             # PDFs temporários (auto-criado)
└── downloads/           # Excel gerado (auto-criado)
```

## Como rodar

### 1. Crie e ative um ambiente virtual (recomendado)
```bash
cd pdf_to_excel
python -m venv venv

# Linux / macOS
source venv/bin/activate

# Windows
venv\Scripts\activate
```

### 2. Instale as dependências
```bash
pip install -r requirements.txt
```

### 3. Rode a aplicação
```bash
python app.py
```

### 4. Acesse no navegador
```
http://localhost:5000
```

## O que o Excel gerado contém

- **Aba "Pedidos"** com cabeçalho colorido (azul corporativo)
- Linha de título + período do relatório
- **Cabeçalho das colunas**: Nº Pedido, Emissão, Cliente, Cód. Produto, Descrição, Quantidade, Preço Unit., Total
- **Agrupamento por produto**: cada produto tem uma linha de cabeçalho destacada
- **Subtotal por produto** com fórmulas `=SUM(...)`
- **Total Geral** ao final com fórmulas
- Linhas alternadas em branco/azul claro
- Painel congelado na linha 4 (scroll com cabeçalho fixo)
- Formatação numérica brasileira
