def inverter_texto(texto):
    return texto[::-1]


while True:
    palavra = input("Digite uma palavra: ")

    if palavra.lower() == "sair":
        break

    invertida = inverter_texto(palavra)

    print("Palavra invertida:", invertida)