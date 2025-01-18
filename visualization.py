import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt
import missingno as msno

data = pd.read_csv(r'Resources/notebooks/datos_con_faltantes.csv')
#Mapa de calor
"""

sns.heatmap(data.isnull(), cbar= False, cmap='viridis')
plt.title('Mapa de calor de datos faltantes')
plt.show()
"""
#matriz de datos faltantes
"""
msno.matrix(data)
plt.title('Matriz de datos')
plt.show()
"""

#Grafico de barras
"""
data.isnull().sum().plot(kind='bar')
plt.title('Cantidad de datos faltantes por columna')
plt.xlabel('Columnas')
plt.ylabel('Numero de datos faltantes')
plt.show()
"""

"Mapa de correlacion de datos faltantes"

msno.heatmap(data)
plt.title('Mapa de correlacion de datos faltantes')
plt.show()