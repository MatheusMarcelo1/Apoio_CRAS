from pathlib import Path

# Pasta onde este arquivo está
pasta_raiz = Path(__file__).resolve().parent
TEMP_PREFIX = "__CRAS_RENAME_TEMP__"

print(f"Pasta atual: {pasta_raiz}")
print("\nPastas que serão renomeadas:\n")

pastas = [p for p in pasta_raiz.rglob("*") if p.is_dir()]

for pasta in pastas:
    novo_nome = pasta.name.upper()

    if pasta.name != novo_nome:
        print(f"{pasta} -> {novo_nome}")

input("\nPressione ENTER para começar...")

for pasta in sorted(pastas, key=lambda p: len(p.parts), reverse=True):
    novo_nome = pasta.name.upper()

    if pasta.name != novo_nome:
        # O nome temporário permite renomear corretamente mesmo no Windows,
        # onde alterações apenas de maiúsculas/minúsculas podem falhar.
        temporario = pasta.with_name(TEMP_PREFIX + pasta.name)
        novo_caminho = pasta.with_name(novo_nome)

        try:
            irmaos = {item.name.casefold() for item in pasta.parent.iterdir() if item != pasta}
            if novo_nome.casefold() in irmaos:
                print(f"ERRO: já existe uma pasta com o nome {novo_nome}: {pasta.parent}")
                continue
            pasta.rename(temporario)
            temporario.rename(novo_caminho)
            print(f"OK: {pasta.name} -> {novo_nome}")

        except Exception as erro:
            print(f"ERRO: {pasta}: {erro}")

print("\nConcluído!")
input("Pressione ENTER para sair...")