from PIL import Image

entrada = "logo_cras.png"
saida = "favicon.png"

imagem = Image.open(entrada)

# Redimensiona mantendo a qualidade
imagem = imagem.resize((32, 32), Image.Resampling.LANCZOS)

# Salva como PNG
imagem.save(saida)

print(f"Favicon criado: {saida}")