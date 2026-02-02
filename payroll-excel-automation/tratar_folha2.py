import pandas as pd
import re

# 1. Ler e salva a planilha 
arquivo_entrada = "folha2020_geral(2).xlsx"
arquivo_saida = "folha_2020_tratada1.xlsx"


# 2. Layout final
colunas_finais = [
    "CPF", "Matrícula", "Nome", "Cargo", "Cargo Limpo", 
    "Situação", "Classe", "Padrão",
    "Cargo ocupacional", "Lotação/ Setor", "Cargo Comissionado",
    "Carga Horária", "Observação do cargo", "Vínculo",
    "Regime Jurídico", "Obs Vínculo", "Salário Base",
    "Remuneração Bruta", "Descontos", "Remuneração Líquida",
    "Data Admissão", "Atualização"
]

# 3. Função para extrair Classe e Padrão do Cargo 
def extrair_classe_padrao(cargo):
    if pd.isna(cargo):
        return None, None
    
    cargo = str(cargo).upper()

    classe = None

    # Número romano
    romano = re.search(r"\b(I|II|III|IV|V|VI|VII|VIII|IX|X)\b", cargo)
    if romano:
        classe = romano.group(1)
    else:
        # Número arábico
        arabico = re.search(r"\b(\d+)\b", cargo)
        if arabico:
            classe = arabico.group(1)
        else:
            # Texto
            if re.search(r"\b(J|JUNIOR|JÚNIOR)\b", cargo):
                classe = "JUNIOR"
            elif re.search(r"\bPLENO\b", cargo):
                classe = "PLENO"
            elif re.search(r"\b(SENIOR|SÊNIOR)\b", cargo):
                classe = "SENIOR"

    padrao = None

    padrao_match = re.search(r"[\s-]([A-Z]{1,2})$", cargo)
    if padrao_match:
        padrao = padrao_match.group(1)

    return classe, padrao

# 4. Lê TODAS as abas
abas = pd.read_excel(arquivo_entrada, sheet_name=None)

abas_tratadas = {}

for nome_aba, df in abas.items():
    
    # 5. Criar colunas faltantes
    for col in colunas_finais:
        if col not in df.columns:
            df[col] = ""

    # 6. Preencher apenas Classe e Padrão vazios
    for idx, row in df.iterrows():
        if not row["Classe"] or not row["Padrão"]:
            classe, padrao = extrair_classe_padrao(row["Cargo"])

            if not row["Classe"] and classe:
                df.at[idx, "Classe"] = classe

            if not row["Padrão"] and padrao:
                df.at[idx, "Padrão"] = padrao

    # 7. Reordenar colunas
    df = df[colunas_finais]

    abas_tratadas[nome_aba] = df

# 8. Salvar mantendo as abas
with pd.ExcelWriter(arquivo_saida, engine="openpyxl") as writer:
    for nome_aba, df in abas_tratadas.items():
        df.to_excel(writer, sheet_name=nome_aba, index=False)

print("Arquivo tratado com sucesso!")
