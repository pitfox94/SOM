
#Se importa numpy para lograr usar semillas random 
import numpy as np
import matplotlib.pyplot as plt
import xlsxwriter
import xlrd
import datetime
from xlsxwriter import worksheet
from dateutil.utils import today
from numpy import float32
from sklearn.model_selection import train_test_split
from sklearn.ensemble import RandomForestClassifier 
from sklearn.preprocessing import StandardScaler
from sklearn.model_selection import cross_val_score
from sklearn import svm
from sklearn import utils
from sklearn import preprocessing
from sklearn.metrics import accuracy_score
import copy
from statistics import mode
from itertools import count
from warnings import catch_warnings
from matplotlib.backends.backend_template import show
import sklearn.metrics



nombreProductos = []

def get_datos_excel():
    #Nombre Archivo - Este Archivo no puede ser Subido producto de su confiencialidad -
    # Pero puede ser susituido por cualquier excel que posea el formato descrito en
    # la documentacion de la investigacion
    ubicacion = "Matx_Tesis2.xlsx"
    workbook = xlrd.open_workbook(ubicacion) 
    hoja = workbook.sheet_by_index(0) 
    clientes_guardados=[]
    for columna in range(2,hoja.ncols):
        nombreProductos.append(hoja.cell_value(0, columna))
    for fila in range(1,48000):
        clientes_guardados.append([])

    
    for fila in range(1,48000-1):
        for columna in range(2,hoja.ncols):          
            #0 ID_Transaccion         
            clientes_guardados[fila].append(hoja.cell_value(fila, columna))         
    print("Vectores de Clientes guardados:")        
    return(clientes_guardados);
    
def crear_excel_som_producto(mapa,filas_mapa,columnas_mapa):
    dia = today().day
    mes = today().month
    ano = today().year
    minuto = datetime.datetime.now().minute
    segundo = datetime.datetime.now().second   
    fecha_actual = str(dia) + ' ' + str(mes) + ' ' + str(ano) + ' ' + str(minuto) + ' ' +  str(segundo) + ' '
    nombre_Archivo = str(fecha_actual) + 'mapaSOMxProductos.xlsx'
    workbook= xlsxwriter.Workbook(nombre_Archivo)
    worksheet = workbook.add_worksheet()
    for fila in range(0,filas_mapa):
        for columna in range(0,columnas_mapa):
            productos = []
            for nProducto in range(0,len(mapa[fila][columna])):
                if mapa[fila][columna][nProducto]==1:
                    productos.append(nombreProductos[nProducto])           
            worksheet.write(fila,columna,str(productos))
    workbook.close()
    print("Archivo: ",nombre_Archivo," creado con exito")
    
def crear_excel_som_productoNum(mapa,filas_mapa,columnas_mapa):
    dia = today().day
    mes = today().month
    ano = today().year
    minuto = datetime.datetime.now().minute
    segundo = datetime.datetime.now().second   
    fecha_actual = str(dia) + ' ' + str(mes) + ' ' + str(ano) + ' ' + str(minuto) + ' ' +  str(segundo) + ' '
    nombre_Archivo = str(fecha_actual) + 'mapaSOMxProductosNum.xlsx'
    workbook= xlsxwriter.Workbook(nombre_Archivo)
    worksheet = workbook.add_worksheet()
    for fila in range(0,filas_mapa):
        for columna in range(0,columnas_mapa):
            productos = []
            for nProducto in range(0,len(mapa[fila][columna])):
                if mapa[fila][columna][nProducto]==1:
                    productos.append(nProducto)           
            worksheet.write(fila,columna,str(productos))
    workbook.close()
    print("Archivo: ",nombre_Archivo," creado con exito")          
            
def crear_excel_som(mapa,filas_mapa,columnas_mapa):
    
    fecha_actual = datetime.datetime.now()
    dia = today().day
    mes = today().month
    ano = today().year
    minuto = datetime.datetime.now().minute
    segundo = datetime.datetime.now().second
    
    fecha_actual = str(dia) + ' ' + str(mes) + ' ' + str(ano) + ' ' + str(minuto) + ' ' +  str(segundo) + ' '
    
    #print(fecha_actual)
    
    nombre_Archivo = str(fecha_actual) + 'mapaSOM.xlsx'
    workbook= xlsxwriter.Workbook(nombre_Archivo)
    worksheet = workbook.add_worksheet()
    
    
    for fila in range(0,filas_mapa):
        for columna in range(0,columnas_mapa):
            worksheet.write(fila,columna,str(mapa[fila][columna]))
    
    workbook.close()
    print("Archivo: ",nombre_Archivo," creado con exito")
    
#Uso de libreria numpy para obtener la distancia euclidiana entre  2 vectores
def distancia_euclidiana(vector1, vector2):
  return np.linalg.norm(vector1 - vector2) 
#Uso de numpy para obtener la distancia manhattan entre 2 nodos dentro del mapa
def distancia_manhattan(r1, c1, r2, c2):
    return np.abs(r1-r2) + np.abs(c1-c2)
def nodo_mas_cercano(muestra, t, mapa, filas, columnas):
  # entrega la fila y columna del nodo(vector) con menor distancia euclidiana
  #del mapa generado  y que va actualizando en cada paso
  resultado = (0,0)
  distancia_mas_corta = 1.0e20
  for i in range(filas):
    for j in range(columnas):
        #Distancia euclidiana entre el nodo en la fila i columna j, y el vector t de la muestra
        #obtenida del archivo
      ed = distancia_euclidiana(mapa[i][j], muestra[t])
      if ed < distancia_mas_corta:
        distancia_mas_corta = ed
        resultado = (i, j)
  return resultado

def split_lista_k(l, k):
    largo_lista = len(l)
    largo_lista_k = largo_lista/k
    #Donde "n" Es igual al largo de las divisiones en partes iguales de la lista recibida
    n = round(largo_lista_k-1)
    #Para cada i en el rango desde 0 hasta el largo de la lista "l"
    for i in range(0, len(l), n):
    
        yield l[i:i+n]    
        
def calcular_Precision(mapa,datos_x,Filas,Columnas):
    
    
    
    
    precision = [];
    limite = 0
    print("Comprobacion de precision...")
    for i in range(1,len(datos_x)):
        
        #randi = np.random.randint(len(datos_x))          
        p = accuracy_score(datos_x[i], mapa[Filas][Columnas])
        if p != 0:              
            precision.append(p)
    print("Comprobacion finalizada...!:)")
    print(precision)
    print("Largo lista de precisiones")
    print(len(precision))
    suma_precision=0
    for i in range(0,len(precision)):
        suma_precision+=precision[i]
    print("Promedio precision")
    print(suma_precision/len(precision));
     
def visualizar_mapa(mapa):
    largo_neurona = len(mapa[0][0])   
    lista_cant = [] 
    #Para cada dimension en las dimensiones de la neurona
    #"dim" es el indice de producto el cual se le esta revisando las relaciones
    for i in range(0,largo_neurona):
        lista_cant.append([])   
    for dim in range(0,largo_neurona):                   
        #Para cada fila en el mapa
        for fil in range(0,len(mapa)):
        #Para cada columna en el mapa
            for col in range(0,len(mapa[fil])):
                producto1 = mapa[fil][col][dim]
                if (producto1 == 1):
                #para cada dimension en la neurona con el producto indice "dim" = 1
                    for dim2 in range(0,len(mapa[fil][col])):
                    #Si el indice "dim2" es distinto del indice "dim" 
                    #quiere decir para no comprar el producto con el mismo
                        if (dim2 != dim):
                            producto2 = mapa[fil][col][dim2]
                            if (producto2==1):
                                lista_cant[dim].append(dim2)                           
    return lista_cant            
                
    
def comprobar_precision(datos_x,mapa,Filas,Columnas):
    precision = [];
    print("Comprobacion de precision...")
    for i in range(1,199):
        for filas in range(0,Filas):
            for columnas in range (0,Columnas):
                randi = np.random.randint(len(datos_x))          
                p = accuracy_score(datos_x[randi], mapa[filas][columnas])
                if p != 0:              
                    precision.append(p)
    print("Comprobacion finalizada...!:)")                
    print(precision)
    print("Largo lista de precisiones")
    print(len(precision))
    suma_precision=0
    for i in range(0,len(precision)):
        suma_precision+=precision[i]
    print("Promedio precision")
    print(suma_precision/len(precision));  

def main():
    
  # Semilla random para evitar que los datos cambien con las iteraciones
  #y cada vez que se llama a la funcion random
  np.random.seed(1)
  #Dimensiones  de los vectores a analizar
  DimensionV = 189
  #Filas y columnas del mapa, para que la informacion sea mas visible
  #se puede agrandar el mapa a gusto, pero eso conlleva mas memoria
  Filas = 10; Columnas = 12
  #Util al momento de generar los radios de los vecinos
  RangoMax = Filas + Columnas
  #Indice de aprendizaje
  AprendizajeMax = 0.1
  #Numero de iteraciones que el algoritmo aprendera a costa de obtener datos
  #aleatorios de la fuente de datos, a mas volumen de datos mayor cantidad de pasos
  PasosMax = 20000
  #get_datos_excel()
  
  #Carga de datos en el sistema
  print("\nCargando datos en el sistema\n")
  #datos_archivo = ".\\iris_data_012.txt" #Nombre del archivo fuente de datos IRIS
  #Arreglo de floats que contienen los vectores del archivo fuente de datos
  #datos_x = np.loadtxt(datos_archivo, delimiter=",", usecols=range(0,4),
  #  dtype=np.float64)
  #datos_y = np.loadtxt(datos_archivo, delimiter=",", usecols=[4],
   # dtype=np.int)
  #Caracteristicas Cliente o Flor como vector
  datos_x = get_datos_excel()
  datos_x = datos_x[1:47999]
  largo_datos_x = len(datos_x)
  
  print("Largo lista partida")

  lista = list(split_lista_k(datos_x,9))
  print("Largo divisiones:")
  print(len(lista))
  print(lista[0])
  
  
  

  #Se crea el Mapa ("mapa") con muestras aleatorias, de dimension  
  #Filas x Columnas con cada una de sus celdas
  #siendo un vector de "DimensionV"
  mapa = np.random.random_sample(size=(Filas,Columnas,DimensionV))
  print(mapa)
  #crear_excel_som(mapa,Filas,Columnas)
  
  print("Construyendo un SOM de dimensiones ",Filas,"x",Columnas," con cada nodo de dimension",DimensionV);
  #Para que el ciclo de aprendizaje no termine hasta que el numero de pasos maximos se cumpla
  for s in range(PasosMax):
    if s % (PasosMax/10) == 0: print("paso = ", str(s))
    #Porcentaje restante por recorrer
    porcentaje_faltante = 1.0 - ((s * 1.0) / PasosMax)
    #Rango actual de inflexion del radio para la asignacion de vecinos
    #Que varia segun el procentaje faltante de pasos por cumplir
    rango_actual = (int)(porcentaje_faltante * RangoMax)
    
    tasa_actual = porcentaje_faltante * AprendizajeMax
    #Se selecciona un vector aleatorio de la muestra para entrenar el mapa SOM
    t = np.random.randint(len(datos_x))
    #Se obtiene la fila y columna, del BMU (best-matching-unit) del vector t
    #dentro del mapa creado
    (bmu_fila, bmu_columna) = nodo_mas_cercano(datos_x, t, mapa, Filas, Columnas)
    #Luego para todos los vectores dentro del mapa
    for i in range(Filas):
      for j in range(Columnas):
        #si la distancia manhattan entre el BMU encontrado y el vector en la fila i y columna j
        #del mapa es menor que el rango actual (de influencia del vector, asociacion de vecinos)
        #modifica el vector encontrado segun la tasa actual.
        if distancia_manhattan(bmu_fila, bmu_columna, i, j) < rango_actual:
          mapa[i][j] = mapa[i][j] + tasa_actual * \
          (datos_x[t] - mapa[i][j])
          
          
        
        
  print("Construccion de SOM completada\n")
  X = mapa[:, 0:28]
  #print("Datos mapa 0 a 29:")
  #print(X);
  #print("Datos mapa 29 ")
  

  #feature_scaler = StandardScaler()  

  #datos_x2= datos_x[0:30]
  print("Dimensiones de la variable X ")
  print(X.shape)
  
  #y = y.reshape(y.shape[0],-1)
  print("informacion primera fila de la variable datos_x, que corresponde a los clientes vectores de compra guardados")
  print(datos_x[1])
  print("Largo de la primera fila de datos_x, ")
  print(len(datos_x[1]))
  #print("Informacion primera neurona del mapa entrenado")
  #print(mapa[0][0])
  print("Largo de la primera neurona del mapa entrenado")
  print(len(mapa[0][0]))
  
  for i in range(0,Filas):
      for j in range(0,Columnas):
          for k in range(0,len(mapa[i][j])):
              mapa[i][j][k] = abs(round(mapa[i][j][k]))

  
 
  precision = [];
  limite = 0
  print("Comprobacion de precision...")
  for i in range(1,199):
      for filas in range(0,Filas):
          for columnas in range (0,Columnas):
            randi = np.random.randint(len(datos_x))          
            p = accuracy_score(datos_x[randi], mapa[filas][columnas])
            if p != 0:              
                precision.append(p)
  
  print("Comprobacion finalizada...!:)")
  print(precision)
  print("Largo lista de precisiones")
  print(len(precision))
  suma_precision=0
  for i in range(0,len(precision)):
      suma_precision+=precision[i]
  print("Promedio precision 1")
  print(suma_precision/len(precision));
  
  
  lista1 = copy.deepcopy(visualizar_mapa(mapa))
  for i in range(0,len(lista1)):
      print("-----------------------------")
      print("Relaciones producto:",i)
      print(sorted(lista1[i]))
      try:
            
          print("el producto mas comprado con el producto ",i,"es el producto",mode(lista1[i]))
          print("Con una cantidad de:",lista1[i].count((mode(lista1[i]))))
      except:
        print("el producto",i,"no tiene moda unica")
  lista_cant_xprod = []
  listazero= [0] * len(lista1[0])  
  print("Largo lista1:")
  print(len(lista1))
  lista_cuenta_productos = []
  for i in range (0,len(lista1)):
      lista_cuenta_productos.append([])
      for j in range(0,len(lista1)):
          if (i != j):
              cuenta = lista1[i].count(j)
              lista_cuenta_productos[i].append(cuenta)
  print("Lista cuenta :")
  print(lista_cuenta_productos)
  print("Largo lista cuenta 0:")
  print(len(lista_cuenta_productos[0]))
  lista_nombre_productos =[]
  for i in range(0,len(lista1)):
      lista_nombre_productos.append([])
      for j in range(0,len(lista1)):
          if(i!=j):
              nombre_producto = str(j)
              lista_nombre_productos[i].append(nombre_producto)

  
  #fig1 = plt.figure(figsize=(12,6))

  
  
  #plt.plot(lista_nombre_productos[2],lista_cuenta_productos[2],'*')
  
  
  #plt.xlabel('Producto')
  #plt.ylabel('Cuenta Productos')
  
  
  #plt.show(fig1)
  
  
  print("Construyendo U-Matrix del SOM")
  u_matrix = np.zeros(shape=(Filas,Columnas), dtype=np.float64)
  for i in range(Filas):
    for j in range(Columnas):
      v = mapa[i][j]   
      suma_dists = 0.0; ct = 0
     
      if i-1 >= 0:    # arriba
        suma_dists += distancia_euclidiana(v, mapa[i-1][j]); ct += 1
      if i+1 <= Filas-1:   # abajo
        suma_dists += distancia_euclidiana(v, mapa[i+1][j]); ct += 1
      if j-1 >= 0:   # izquierda
        suma_dists += distancia_euclidiana(v, mapa[i][j-1]); ct += 1
      if j+1 <= Columnas-1:   # derecha
        suma_dists += distancia_euclidiana(v, mapa[i][j+1]); ct += 1
      
      u_matrix[i][j] = suma_dists / ct
  print("U-Matrix construida \n")
  fig2 = plt.imshow(u_matrix, cmap='gray')  # negros = cercanos = clusters
  plt.show(fig2)
  
  #ax3 = fig.add_subplot(133)
  #ax3.scatter(lista1[0],listazero)
  
          
          
              
              
          
    
      
  
  
  
  crear_excel_som(mapa, Filas, Columnas)
  crear_excel_som_producto(mapa,Filas, Columnas)
  crear_excel_som_productoNum(mapa, Filas, Columnas)

if __name__ == '__main__':
    main();