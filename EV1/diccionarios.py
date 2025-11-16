info_sam= {"Nombre": "Samantha" , "Apellidos": "Cervantes Valdez",
           "Estudios":["Primaria", "Secundaria", "Preparatoria", "Universidad"], "Color favorito": ["Morado, Negro"], 
            "Edad": "19 Años" }

print(info_sam["Color favorito"])
print(info_sam["Edad"])
print(info_sam["Nombre"], ",", info_sam["Apellidos"])
for estu in info_sam ["Estudios"]:
    print (estu) 

print("***********")


datos = dict(colores=["verde","azul","rojo"]) 
print(type(datos))
print("***********")

dicc = {"edad": 19, "materias": 7,"mascotas": 1, "hermanos": 2}
recurso=dicc.items()
print(recurso)
claves= dicc.keys()
print (claves)
valores= dicc.values()
print(valores)

dic= dict.fromkeys(["a","b","c"],(1,2,3))
print (dic)
print(info_sam.get("Edad"))
eliminar = dicc.pop("hermanos")
print(eliminar)

dicc2 ={"materias": 5}
nuevo = dicc.update(dicc2)
print(nuevo)