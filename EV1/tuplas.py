info_alumno =("Samantha",19, 35)
print("Alumno", info_alumno)

nombre, edad, grupo = info_alumno
print("Nombre: ", nombre)
print("edad: ", edad)
print("grupo: ", grupo)



x = ("apple", "banana", "cereza")
y = list(x)
y[1] = "kiwi"
x = tuple(y)

print(x)
print(x[-1])

auditorio = (2,6)
print("Ubicación del auditorio:  ", auditorio)