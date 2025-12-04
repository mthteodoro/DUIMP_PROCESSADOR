import pandas as pd
from tqdm import tqdm
import os
 
# ======================================================
# 1. DEFINIR CAMINHO DO ARQUIVO EXCEL
# ======================================================
arquivo = r"C:\Users\matheus.teodoro\Desktop\PYTHON\Catalogo WESTROCK - DUIMP 12092025.xlsx"
 
# ======================================================
# FUNÇÃO PARA FORMATAR O CÓDIGO FISCAL (0000.00.00)
# ======================================================
def formatar_classificacao(codigo):
    """
    Recebe algo como '38229000' e retorna '3822.90.00'
    """
    codigo = str(codigo).strip()
 
    if len(codigo) != 8 or not codigo.isdigit():
        return ""
 
    return f"{codigo[0:4]}.{codigo[4:6]}.{codigo[6:8]}"
 
# ======================================================
# 2. LER PLANILHA E SUAS ABAS
# ======================================================
print("Lendo todas as abas do arquivo...")
xls = pd.ExcelFile(arquivo)   # <-- CORRIGIDO
abas = xls.sheet_names
 
print(f"{len(abas)} abas encontradas.")
 
# ======================================================
# LISTA FINAL QUE SERÁ TRANSFORMADA EM DATAFRAME
# ======================================================
dados_finais = []
 
# ======================================================
# 3. PROCESSAR CADA ABA
# ======================================================
for aba in tqdm(abas, desc="Processando abas"):
 
    # Nome da aba → código fiscal formatado
    cod_fiscal = formatar_classificacao(aba)
 
    # Lê a aba
    df = pd.read_excel(arquivo, sheet_name=aba, dtype=str)
 
    # Se a aba não contém a coluna necessária, pular
    if "cod_material_cliente" not in df.columns:
        continue
 
    # Remover linhas totalmente vazias
    df = df.dropna(how="all")
 
    # Identificar a posição da coluna Descricao_CH_Completa_PT
    if "Descricao_CH_Completa_PT" not in df.columns:
        continue
 
    idx_desc = df.columns.get_loc("Descricao_CH_Completa_PT")
 
    # Todas as colunas após Descricao_CH_Completa_PT são atributos
    col_atributos = list(df.columns[idx_desc + 1:])
 
    # Barra de progresso por material
    for _, row in tqdm(df.iterrows(), total=len(df), leave=False, desc=f"Materiais da aba {aba}"):
 
        cod_material = str(row.get("cod_material_cliente", "")).strip()
 
        if cod_material == "" or cod_material.lower() == "nan":
            continue
 
        # Para cada coluna de atributo, gerar uma linha na planilha final
        for col in col_atributos:
 
            valor = row.get(col, "")
            valor = "" if str(valor).lower() == "nan" else str(valor).strip()
 
            # DESC_ATRIBUTO: se for atributo do tipo "999 - OUTROS"
            desc_atributo = ""
            if "-" in valor and len(valor.split("-")[0].strip()) <= 4:
                desc_atributo = valor
 
            dados_finais.append({
                "COD_CLIENTE": "WEST ROCK",
                "COD_MATERIAL": cod_material,
                "COD_ATRIBUTO": col,
                "VALOR_ATRIBUTO": valor,
                "DESC_ATRIBUTO": desc_atributo,
                "COD_CLASSIF_FISCAL": cod_fiscal
            })
 
# ======================================================
# 8. TRANSFORMAR EM DATAFRAME FINAL
# ======================================================
df_final = pd.DataFrame(dados_finais)
 
# Garantir ordem correta das colunas
df_final = df_final[
    ["COD_CLIENTE", "COD_MATERIAL", "COD_ATRIBUTO",
     "VALOR_ATRIBUTO", "DESC_ATRIBUTO", "COD_CLASSIF_FISCAL"]
]
 
# ======================================================
# SALVAR RESULTADO
# ======================================================
saida = os.path.join(os.path.dirname(arquivo), "Resultado_WESTROCK_PROCESSADO.xlsx")
df_final.to_excel(saida, index=False)
 
print("\n================================================")
print("PROCESSAMENTO CONCLUÍDO!")
print(f"Arquivo gerado em: {saida}")
print("================================================")
