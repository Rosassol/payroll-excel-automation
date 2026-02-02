import pandas as pd
import re

# 1. Ler a planilha 
df = pd.read_excel("folha2020_geral(2).xlsx")

# 2. Layout final desejado (ordem correta) 
colunas_finais = [
    "CPF", "Matrícula", "Nome", "Cargo", "Cargo Limpo", 
    "Situação", "Classe", "Padrão",
    "Cargo ocupacional", "Lotação/ Setor", "Cargo Comissionado",
    "Carga Horária", "Observação do cargo", "Vínculo",
    "Regime Jurídico", "Obs Vínculo", "Salário Base",
    "Remuneração Bruta", "Descontos", "Remuneração Líquida",
    "Data Admissão", "Atualização"
]


# 3. Criar colunas que não existem (sem mexer nas existentes) 
for col in colunas_finais:
    if col not in df.columns:
        df[col] = ""

# 4. Função para extrair Classe e Padrão do Cargo 
def extrair_classe_padrao(cargo):
    """
    Exemplos que isso cobre:
    PROFESSOR A II-H
    AUXILIAR ADM III-AG
    TECNICO IV-AB
    """
    if pd.isna(cargo):
        return None, None

    # Classe: número romano
    classe_match = re.search(r"\b(I|II|III|IV|V|VI|VII|VIII|IX|X)\b", cargo)

    # Padrão: após hífen, uma ou duas letras
    padrao_match = re.search(r"-([A-Z]{1,2})\b", cargo)

    classe = classe_match.group(1) if classe_match else None
    padrao = padrao_match.group(1) if padrao_match else None

    return classe, padrao

# 5. Preencher SOMENTE se Classe/Padrão estiverem vazios 
for idx, row in df.iterrows():
    if not row["Classe"] or not row["Padrão"]:
        classe, padrao = extrair_classe_padrao(row["Cargo"])

        if not row["Classe"] and classe:
            df.at[idx, "Classe"] = classe

        if not row["Padrão"] and padrao:
            df.at[idx, "Padrão"] = padrao

#  6. Reordenar colunas 
df = df[colunas_finais]

# 7. Salvar nova planilha 
df.to_excel("folha2020_tratada2.xlsx", index=False)
