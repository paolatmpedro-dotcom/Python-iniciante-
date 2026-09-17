import random


def sortear_dado():
    return random.randint(1, 6)


try:
    quantidade = int(input("Quantas vezes deseja lançar o dado? "))

    for i in range(quantidade):
        resultado = sortear_dado()
        print("Resultado:", resultado)

except ValueError:
    print("Digite um número inteiro válido.")