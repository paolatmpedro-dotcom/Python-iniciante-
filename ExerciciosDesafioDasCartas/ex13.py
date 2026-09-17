import random


def validar_chute(chute):
    try:
        return int(chute)
    except ValueError:
        return None


numero_secreto = random.randint(1, 20)

tentativas = 5
acertou = False

print("===== JOGO DE ADIVINHAÇÃO =====")
print("Tente adivinhar um número de 1 a 20.")
print("Você tem 5 tentativas.")

for i in range(tentativas):
    chute = input("Digite seu chute: ")

    numero = validar_chute(chute)

    if numero is None:
        print("Digite um número inteiro válido.")
        continue

    if numero == numero_secreto:
        print("Parabéns! Você acertou!")
        acertou = True
        break

    elif numero < numero_secreto:
        print("O número secreto é maior.")

    else:
        print("O número secreto é menor.")

if not acertou:
    print("Você perdeu!")
    print("O número secreto era:", numero_secreto)