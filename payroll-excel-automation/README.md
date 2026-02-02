# 📊 Tratamento de Planilhas de Folha de Pagamento

Script em Python para padronizar planilhas de **folha de pagamento**, criando colunas faltantes e extraindo automaticamente **Classe** e **Padrão** a partir da coluna **Cargo**, sem alterar dados já preenchidos.

Funciona com **arquivos Excel que possuem várias abas** (ex: uma aba por mês).

---

## ✨ Funcionalidades

- Lê arquivos Excel (`.xlsx`)
- Processa **todas as abas** do arquivo
- Cria automaticamente colunas que não existirem
- Mantém a **ordem padrão** das colunas
- Extrai:
  - **Classe** → número romano (`I`, `II`, `III`, …)
  - **Padrão** → uma ou duas letras (`A`, `H`, `AG`, `AB`, …)
- Preenche **somente** `Classe` e `Padrão` quando estiverem vazias
- Não altera dados já preenchidos
- Gera um novo arquivo Excel tratado

---

## 🧱 Estrutura do Projeto
```
folha_pagamento/
│
├── .venv/
├── folha.py
├── README.md
└── tratar_folha.py
```
---

## 📑 Estrutura do Excel

### Entrada (exemplo)
```bash
Classe | Padrão
II     | H
```

> O script também reconhece padrões com duas letras, como `AG` ou `AB`.

---

## 🛠️ Requisitos

- Python 3.9 ou superior
- Bibliotecas:
  - `pandas`
  - `openpyxl`

---

## ⚙️ Instalação

### 1️⃣ Criar ambiente virtual

**Windows**
```bash
python -m venv .venv
.venv\Scripts\activate
```

**Linux/WSL**
```bash
python3 -m venv .venv
source .venv/bin/activate
```
---

### 2️⃣ Instalar dependências

```bash
pip install pandas openpyxl
```
---

### ▶️ Como executar

- Coloque o arquivo Excel (ex: folha_2020.xlsx) na pasta do projeto
- Ajuste o nome do arquivo no código, se necessário
- Execute:
```bash
python tratar_folha.py
```

---

### 📤 Resultado

Será gerado um novo arquivo:

```bash
folha_2020_tratada.xlsx
```

- Todas as abas originais são mantidas
- As colunas seguem o layout padrão
- Apenas Classe e Padrão são preenchidas automaticamente

---

### ⚠️ Observações

- Feche o Excel antes de executar o script
- Os nomes das abas podem variar (Jan, Fevereiro, Mar, etc.)
- Caso alguma aba tenha estrutura diferente, o código pode ser adaptado

--- 

## © Copyright

© 2026 Rayssa Gomes. Todos os direitos reservados.

Este projeto foi desenvolvido para fins acadêmicos e administrativos, com o objetivo de automatizar e padronizar o tratamento de planilhas de folha de pagamento.

É permitida a utilização, modificação e adaptação do código para uso pessoal ou institucional, desde que mantida a referência à autora.  
A redistribuição ou uso comercial sem autorização prévia é proibida.

O software é fornecido **“como está”**, sem garantias de qualquer tipo, expressas ou implícitas, incluindo,  mas não se limitando às garantias de comercialização, adequação a um propósito específico ou ausência de erros.

A autora não se responsabiliza por eventuais danos, perdas de dados ou inconsistências decorrentes do uso deste software.

---
