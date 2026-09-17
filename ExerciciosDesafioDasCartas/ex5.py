try:
  n1 = int(input("Digite um numero inteiro: "))
  n2 = int(input("Digite um numero inteiro: "))

  if n1 == 0 or n2 == 0:
    print("Não é possível realizar a divisão por ZERO")
  else:
    divisao = n1 / n2
    print("A divisão dos números é:", divisao)

except ValueError:
  print("Não é permitido texto, apenas números")