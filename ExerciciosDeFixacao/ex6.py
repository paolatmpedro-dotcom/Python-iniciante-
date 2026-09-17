#ex6

preco = input("Digite o preço do produto: ")
percentual = input("Digite o percentual de desconto: ")

preco = int(preco)
percentual = int(percentual)

desconto = (preco * percentual)/100
precofinal = preco - desconto


print("\n O preço é: ", preco, "\nO percentual de desconto é",percentual,"%", "\nO desconto foi de", desconto, "e o valor final é de",  precofinal)
print(f"\nO preço é: {preco} \nO percentual de desconto é {percentual} % \nO desconto foi de{desconto} e o valor final é de {precofinal}")
