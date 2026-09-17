soma = 0
for i in range(5):
    lista_idade = int(input(f"Digite a idade {i+1}: "))
    soma += lista_idade

media = soma / 5
print("A média das idades é:", media)
