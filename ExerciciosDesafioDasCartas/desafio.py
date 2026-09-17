import random

sorteio = random.randint(1, 10)

def pedir_palpite():
    while True:
        try:
            usuario = int(input("Digite um número de 1 a 10: "))

            if usuario < 1 or usuario > 10:
                print("Número inválido. Digite apenas números entre 1 e 10!")
                continue

            return usuario

        except ValueError:
            print("Não é permitido texto, apenas números.")

while True:
    usuario = pedir_palpite()

    if usuario == sorteio:
        print("Você acertou!!!")
        break
    else:
        print("Tente novamente!!")
