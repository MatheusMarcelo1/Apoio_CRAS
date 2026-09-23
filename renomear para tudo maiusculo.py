
from pathlib import Path

# Pasta onde este arquivo renomear.py está
pasta_raiz = Path(__file__).resolve().parent

print(f"Pasta raiz: {pasta_raiz}")
print("\nItens que serão renomeados:\n")

# ============================================================
# LOCALIZAR ARQUIVOS E PASTAS
# ============================================================

arquivos = [p for p in pasta_raiz.rglob("*") if p.is_file()]
pastas = [p for p in pasta_raiz.rglob("*") if p.is_dir()]

# ============================================================
# MOSTRAR ARQUIVOS QUE SERÃO RENOMEADOS
# ============================================================

print("=== ARQUIVOS ===")

for arquivo in arquivos:
    novo_nome = arquivo.name.upper()

    if arquivo.name != novo_nome:
        print(f"{arquivo} -> {novo_nome}")

# ============================================================
# MOSTRAR PASTAS QUE SERÃO RENOMEADAS
# ============================================================

print("\n=== PASTAS ===")

for pasta in pastas:
    novo_nome = pasta.name.upper()

    if pasta.name != novo_nome:
        print(f"{pasta} -> {novo_nome}")

input("\nPressione ENTER para começar...")

# ============================================================
# RENOMEAR ARQUIVOS
# ============================================================

print("\nRenomeando arquivos...\n")

for arquivo in sorted(arquivos, key=lambda p: len(p.parts), reverse=True):

    novo_nome = arquivo.name.upper()

    if arquivo.name != novo_nome:

        temporario = arquivo.with_name("__TEMP_RENAME__" + arquivo.name)
        novo_caminho = arquivo.with_name(novo_nome)

        try:
            arquivo.rename(temporario)
            temporario.rename(novo_caminho)

            print(f"ARQUIVO OK: {arquivo.name} -> {novo_nome}")

        except Exception as erro:
            print(f"ERRO NO ARQUIVO: {arquivo}: {erro}")

# ============================================================
# RENOMEAR PASTAS
# ============================================================

print("\nRenomeando pastas...\n")

for pasta in sorted(pastas, key=lambda p: len(p.parts), reverse=True):

    novo_nome = pasta.name.upper()

    if pasta.name != novo_nome:

        temporario = pasta.with_name("__TEMP_RENAME__" + pasta.name)
        novo_caminho = pasta.with_name(novo_nome)

        try:
            pasta.rename(temporario)
            temporario.rename(novo_caminho)

            print(f"PASTA OK: {pasta.name} -> {novo_nome}")

        except Exception as erro:
            print(f"ERRO NA PASTA: {pasta}: {erro}")

print("\n===================================")
print("TUDO CONCLUÍDO!")
print("===================================")

input("\nPressione ENTER para sair...")

