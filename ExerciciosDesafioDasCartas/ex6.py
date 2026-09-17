cores = ["azul", "verde", "vermelho"]

try:
    indice = int(input("Digite um índice: "))
    print("O índice digitado é referente à cor:", cores[indice])

except IndexError:
    print("O índice digitado não existe!")

except ValueError:
    print("Digite um número inteiro!")
