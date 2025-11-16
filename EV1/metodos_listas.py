series =["orange is the new black","La casa de las flores","yellow jacteks", "atpical","ab","cd"]
peliculas = ["Pitch perfect","Endgame","Iron man", "Carol","Parasitos", "Norbit","Scary movie"]
 
peliculas.pop(-1)
peliculas.remove("Iron man")
peliculas.append("Los guardianes de las galaxias")
print(peliculas )
 
 
series.insert(3,"Squid Game")
del series[4:6]
series[4]= "Atypical"
print(series)
 
 
peliculas.extend(series)
print(peliculas)
 
 
peliculas.clear()
print(peliculas)