

import math
import pandas as pd


calculo=(input("Escriebme el calculo que quieres hacer"))
lista=list(calculo)
print(lista)
numeros=[]
operadores=[]
aceptados=["+","-","*","/","%","**","[","]","(",")"]



for i in lista: # separar numeros y operadores
    if i.isdigit(): 
        numeros.append(int(i))
    else:
        operadores.append(i)

print(numeros)
print(operadores)

for i in range(len(operadores)):
    if len(operadores(i)) in aceptados:
        continue
    else:
        print("Con eso no se peude hacer una operacion amtematica")

'''
for i in numeros:
    if i != int or float:
       numeros.replace("''","")
       operadores.append(i)
    else:
        continue   

print(numeros)
print(operadores)
var =2*7
print(var)
'''


