import pandas as pd
import qrcode
import os

arquivo_excel = r"C:\Users\Mind\source\repos\Qrcode\Qrcode\cracha.xlsx"

pasta_saida = r"C:\Users\Mind\source\repos\Qrcode\Qrcode\teste"

if not os.path.exists(pasta_saida):
    os.makedirs(pasta_saida)


if not os.path.exists(arquivo_excel):
    print(f"Arquivo não encontrado: {arquivo_excel}")
    exit()


def ler_excel(caminho_arquivo):
    return pd.read_excel(caminho_arquivo)


df = ler_excel(arquivo_excel)


for index, linha in df.iterrows():
    nome = str(linha["Nome"]).strip()
    cpf = str(linha["CPF"]).strip()

    dados = f"{nome} - {cpf}"
    
    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_L,
        box_size=10,
        border=4,
    )
    qr.add_data(dados)
    qr.make(fit=True)

   
    img = qr.make_image(fill="black", back_color="white")

    
    caminho_saida = os.path.join(pasta_saida, f"{nome}.png")
    img.save(caminho_saida)

    print(f" QR Code gerado para {nome} ({cpf}) e salvo em: {caminho_saida}")

print(" Todos os QR Codes foram gerados com sucesso!")
