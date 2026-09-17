import random

nomes = []

for i in range (5):
    novo_nome = input("Digite uma nome: ")
    nomes.append(novo_nome)

random.shuffle(nomes)
monitor = random.choice(nomes)

print("Lista de nomes é:", nomes, "monitor do grupo: ",monitor)    