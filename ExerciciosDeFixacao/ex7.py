#ex7

salario = input("Digite o salario atual: ")
percentual = input("Digite o percentual de aumento: ")

salario = int(salario)
percentual = int(percentual)

aumento = (salario * percentual)/100
novoSalario = salario + aumento

print("\n O salario é: ", salario, "\nO percentual de aumento é",percentual,"%", "\nO aumento foi de", aumento, "o novo salario é de",  novoSalario)
print(f"\nO salario é: {salario} \nO percentual de aumento é {percentual} % \nO aumento foi de{aumento} o novo salraio é {novoSalario}")
