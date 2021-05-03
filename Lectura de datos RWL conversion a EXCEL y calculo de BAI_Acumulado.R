library(dplR)
library(openxlsx)
library(xlsx)

#PARTE 1
#Lectura de todos los archivos en una carpeta
x=list.files(path="D:/pablo/Programas/OneDrive/OneDrive - Universidad Pablo de Olavide de Sevilla/Pablo Casas Gomez/Himalaya/Ring width measurements/Archivos RWL")

#Quitamos la parte del archivo que hace referencia a la extension del archivo
for(i in c(1:119)){
  x[i]=substr(x[i], 1, 7)
}

#Comprobación de que se pueden leer todos los archivos .rwl
for (i in x){
  print(paste("Leyendo archivo: ",i))
  data=read.tucson(paste("D:/pablo/Desktop/Lo mas actualizado/Doctorado/Manuscrito 7 (Himalaya)/Datos RWL/",i,".rwl",sep=""))
}

#Abrimos archivos .rwl concretos para ver el error que tenemos
data=read.tucson(paste("D:/pablo/Desktop/Lo mas actualizado/Doctorado/Manuscrito 7 (Himalaya)/Datos RWL/paki041.rwl",sep=""))

#PARTE 2
#Lectura de archivo .rwl y escritura de nuevo archivo .xlsx
for (i in x){
  data=read.tucson(paste("D:/pablo/Desktop/Lo mas actualizado/Doctorado/Manuscrito 7 (Himalaya)/Datos RWL/",i,".rwl",sep=""))
  
  write.xlsx(data,paste("D:/pablo/Desktop/Lo mas actualizado/Doctorado/Manuscrito 7 (Himalaya)/ArchivosExcel/",i,".xlsx",sep=""),row.names = TRUE,col.names = TRUE,showNA = FALSE)
  
  }

#PARTE 3
#Calculo de Incremento de BAI (Celda Anterior + Celda Actual)
x=list.files(path="D:/pablo/Desktop/Lo mas actualizado/Doctorado/Manuscrito 7 (Himalaya)/ArchivosExcel")

#Eliminamos la Carpeta donde guardaremos los nuevos archivos
x=x[!x %in% "Calculado_Incremento_BAI"]

for(i in c(1:119)){
  x[i]=substr(x[i], 1, 7)
}


#Bucle para poder leer archivo, crear nueva matriz, copiar nombre de columnas y ponerlas en la nueva matriz
#realizar el calculo de BAI y crear una nueva matriz con estos valores, eliminamos los 0 por NA y salvamos el archivo
for(z in x){

  #Lectura de matriz original
  n=read.xlsx(paste("D:/pablo/Desktop/Lo mas actualizado/Doctorado/Manuscrito 7 (Himalaya)/ArchivosExcel/",z,".xlsx",sep=""),rowNames = TRUE)

  #Dimensiones de matriz para el bucle
  dimension=dim(n)
  fil=dimension[1]
  col=dimension[2]

  #Creacion de nueva matriz
  matriz=matrix(nrow = fil, ncol = col)
  
  #Copiar los nombres de columnas
  colnames(matriz)=colnames(n)
  rownames(matriz)=rownames(n)

  #Sustituimos NA por 0 en cada dato
  n[is.na(n)] <- 0
  
  #Calculo de la nueva matriz con valores de BAI
  for(i in c(2:fil)){
    for(j in c(1:col)){
      if(i==2){
        matriz[[i,j]]=(n[[i,j]]+n[[i-1,j]])
      }else{
        matriz[[i,j]]=(n[[i,j]]+matriz[[i-1,j]])
      }
    }
  }
  
  #Volvemos a transformar los 0 en NA
  matriz[matriz == 0] <- NA
  
  #Trasnformamos la matriz en data.frame
  matriz.df <- as.data.frame(matriz)
  
  #Escritura de la nueva matriz 
  write.xlsx(matriz.df,paste("D:/pablo/Desktop/Lo mas actualizado/Doctorado/Manuscrito 7 (Himalaya)/ArchivosExcel/Calculado_Incremento_BAI/",z,".xlsx",sep=""),row.names = TRUE,col.names = TRUE,showNA = FALSE)
}

#Calculo de Incremento de Incremento Superficie Anual(=((PI()*(((M97)^2)-((M96)^2))))/100)
x=list.files(path="D:/pablo/Desktop/Lo mas actualizado/Doctorado/Manuscrito 7 (Himalaya)/ArchivosExcel/Calculado_Incremento_BAI")

#Eliminamos la Carpeta donde guardaremos los nuevos archivos
x=x[!x %in% "Incremento_Superficie_Anual"]

for(i in c(1:119)){
  x[i]=substr(x[i], 1, 7)
}

#Bucle para poder leer archivo, crear nueva matriz, copiar nombre de columnas y ponerlas en la nueva matriz
#realizar el calculo de BAI y crear una nueva matriz con estos valores, eliminamos los 0 por NA y salvamos el archivo
for(z in x){
  
  #Lectura de matriz original
  n=read.xlsx(paste("D:/pablo/Desktop/Lo mas actualizado/Doctorado/Manuscrito 7 (Himalaya)/ArchivosExcel/Calculado_Incremento_BAI/",z,".xlsx",sep=""),rowNames = TRUE)
  
  #Dimensiones de matriz para el bucle
  dimension=dim(n)
  fil=dimension[1]
  col=dimension[2]
  
  #Creacion de nueva matriz
  matriz=matrix(nrow = fil, ncol = col)
  
  #Copiar los nombres de columnas
  colnames(matriz)=colnames(n)
  rownames(matriz)=rownames(n)
  
  #Sustituimos NA por 0 en cada dato
  n[is.na(n)] <- 0
  
  #Calculo de la nueva matriz con valores de BAI
  for(i in c(2:fil)){
    for(j in c(1:col)){
      if(i==2){0
        matriz[[i,j]]=((pi*(n[[i,j]]^2))-(pi*(n[[i-1,j]])^2))/100
      }else{
        matriz[[i,j]]=((pi*(n[[i,j]]^2))-(pi*(n[[i-1,j]])^2))/100
      }
    }
  }
  

  
  #Volvemos a transformar los 0 en NA
  matriz[matriz == 0] <- NA
  
  #Trasnformamos la matriz en data.frame
  matriz.df <- as.data.frame(matriz)
  
  #Escritura de la nueva matriz 
  write.xlsx(matriz.df,paste("D:/pablo/Desktop/Lo mas actualizado/Doctorado/Manuscrito 7 (Himalaya)/ArchivosExcel/Calculado_Incremento_BAI/Incremento_Superficie_Anual/",z,".xlsx",sep=""),row.names = TRUE,col.names = TRUE,showNA = FALSE)
}
 
