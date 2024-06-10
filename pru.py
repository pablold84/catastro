resultado = None

resultado = 2

def p1():
    global resultado
    resultado = 3
    print("imprimo p1, ", resultado)
    
def p2():
    global resultado
    print("imprimo p2, ", resultado)


p1()
p2()