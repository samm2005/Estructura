#un conjunto es una colección desordenada de elementos únicos. 
# No permite duplicados y los elementos no tienen un orden específico.
s= frozenset(range(10))
s1= {1,234,5,True, "sam"}


s1.add(False)

print(s1)

st2 = {"manzana", "banana", "cereza"}
frutas = {"piña", "cherry","mango", "papaya"}

st2.update(frutas)
print(st2)

s1.remove(234)
print(s1)

copiadi= st2.copy
print(copiadi)

frutitas = {"manzana", "banana", "cherry"}



print(frutitas)

nov= frutas.difference(frutitas)
print(nov)

frutitas.clear()

