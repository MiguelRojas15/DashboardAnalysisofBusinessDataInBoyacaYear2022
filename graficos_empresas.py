import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from matplotlib.colors import LinearSegmentedColormap
import numpy as np

# Cargar el archivo limpio
df = pd.read_excel("PROYECTO_BOYACA_EMPRESAS_LIMPIO.xlsx")

# Configuración del estilo de los gráficos
plt.style.use('dark_background')
sns.set_palette("husl")

# Colores 
colores_gradiente = ['#FF6B6B', '#4ECDC4', '#45B7D1', '#96CEB4', '#FFEAA7', '#DDA0DD', '#98D8C8']
colores_neon = ['#00f5ff', '#ff073a', '#39ff14', '#ff9500', '#bf00ff', '#ffff00']

# 1. HISTOGRAMA 
plt.figure(figsize=(12, 7))
plt.gca().set_facecolor('#1a1a1a')


n, bins, patches = plt.hist(df['Ventas mensuales (Millones)'].dropna(), 
                           bins=30, alpha=0.8, edgecolor='white', linewidth=0.8)


cm = plt.cm.get_cmap('plasma')
for i, p in enumerate(patches):
    p.set_facecolor(cm(i / len(patches)))

#  KDE 
kde_data = df['Ventas mensuales (Millones)'].dropna()
from scipy.stats import gaussian_kde
kde = gaussian_kde(kde_data)
x_range = np.linspace(kde_data.min(), kde_data.max(), 200)
plt.plot(x_range, kde(x_range) * len(kde_data) * (bins[1] - bins[0]), 
         color='#00f5ff', linewidth=3, label='Densidad')

plt.title(" Distribución de Ventas Mensuales", fontsize=18, fontweight='bold', color='white', pad=20)
plt.xlabel("Ventas mensuales (Millones)", fontsize=14, color='white')
plt.ylabel("Número de empresas", fontsize=14, color='white')
plt.grid(True, alpha=0.3, color='gray')
plt.legend()
plt.tight_layout()
plt.savefig("grafico_histograma_mejorado.png", dpi=300, bbox_inches='tight', facecolor='#1a1a1a')
plt.show()

# 2. GRÁFICO DE CAJA 
plt.figure(figsize=(14, 8))
plt.gca().set_facecolor('#0f0f0f')

# 
box_plot = sns.boxplot(data=df, x="SectorProductivo", y="Ventas mensuales (Millones)", 
                       palette=colores_gradiente, linewidth=2)

# Personalizar cajas
for patch in box_plot.artists:
    patch.set_alpha(0.8)
    patch.set_edgecolor('white')
    patch.set_linewidth(2)

plt.title(" Ventas Mensuales por Sector Productivo", fontsize=18, fontweight='bold', color='white', pad=20)
plt.xlabel("Sector Productivo", fontsize=14, color='white')
plt.ylabel("Ventas mensuales (Millones)", fontsize=14, color='white')
plt.xticks(rotation=45, ha='right', color='white')
plt.yticks(color='white')
plt.grid(True, alpha=0.2, color='gray')
plt.tight_layout()
plt.savefig("grafico_caja_mejorado.png", dpi=300, bbox_inches='tight', facecolor='#0f0f0f')
plt.show()

# 3. GRÁFICO DE PASTEL 
plt.figure(figsize=(10, 10))
plt.gca().set_facecolor('#1a1a1a')

genero_counts = df['Genero responsable'].value_counts()

# Colores 
colors_pie = ['#ff073a', '#39ff14', '#00f5ff', '#ff9500']


explode = (0.05, 0.05, 0.05, 0.05)[:len(genero_counts)]

wedges, texts, autotexts = plt.pie(genero_counts, labels=genero_counts.index, 
                                   autopct='%1.1f%%', startangle=90, 
                                   colors=colors_pie[:len(genero_counts)],
                                   explode=explode, shadow=True,
                                   textprops={'fontsize': 12, 'fontweight': 'bold', 'color': 'white'})


for autotext in autotexts:
    autotext.set_color('black')
    autotext.set_fontweight('bold')
    autotext.set_fontsize(14)

plt.title(" Distribución por Género del Responsable", fontsize=18, fontweight='bold', color='white', pad=30)
plt.tight_layout()
plt.savefig("grafico_pastel_mejorado.png", dpi=300, bbox_inches='tight', facecolor='#1a1a1a')
plt.show()

# 4. GRÁFICO DE BARRAS HORIZONTAL 
plt.figure(figsize=(12, 8))
plt.gca().set_facecolor('#0d1117')

top_municipios = df['Municipio'].value_counts().nlargest(10)


bars = plt.barh(range(len(top_municipios)), top_municipios.values, 
                color=colores_gradiente[:len(top_municipios)], 
                edgecolor='white', linewidth=1.5, alpha=0.9)


for i, bar in enumerate(bars):
    bar.set_alpha(0.8)
    # Agregar valor al final de cada barra
    plt.text(bar.get_width() + 0.5, bar.get_y() + bar.get_height()/2, 
             f'{int(bar.get_width())}', ha='left', va='center', 
             color='white', fontweight='bold', fontsize=10)

plt.yticks(range(len(top_municipios)), top_municipios.index, color='white', fontsize=11)
plt.xticks(color='white')
plt.title(" Top 10 Municipios con Más Empresas", fontsize=18, fontweight='bold', color='white', pad=20)
plt.xlabel("Número de empresas", fontsize=14, color='white')
plt.ylabel("Municipio", fontsize=14, color='white')
plt.grid(True, alpha=0.2, axis='x', color='gray')
plt.tight_layout()
plt.savefig("grafico_barras_mejorado.png", dpi=300, bbox_inches='tight', facecolor='#0d1117')
plt.show()

# 5. GRÁFICO DE DISPERSIÓN 
plt.figure(figsize=(14, 9))
plt.gca().set_facecolor('#0a0a0a')


scatter = sns.scatterplot(data=df, x="Numero empleados", y="Ventas mensuales (Millones)", 
                         hue="SectorProductivo", size="Ventas mensuales (Millones)",
                         sizes=(50, 400), alpha=0.8, palette='Set1', edgecolor='white', linewidth=0.5)


plt.legend(bbox_to_anchor=(1.05, 1), loc='upper left', frameon=True, 
           facecolor='#1a1a1a', edgecolor='white')
for text in plt.gca().get_legend().get_texts():
    text.set_color('white')

plt.title("Empleados vs Ventas Mensuales por Sector", fontsize=18, fontweight='bold', color='white', pad=20)
plt.xlabel("Número de empleados", fontsize=14, color='white')
plt.ylabel("Ventas mensuales (Millones)", fontsize=14, color='white')
plt.xticks(color='white')
plt.yticks(color='white')
plt.grid(True, alpha=0.2, color='gray')
plt.tight_layout()
plt.savefig("grafico_dispersion_mejorado.png", dpi=300, bbox_inches='tight', facecolor='#0a0a0a')
plt.show()

# GRÁFICO DE CORRELACIÓN MODERNO
plt.figure(figsize=(12, 10))

# Seleccionar solo columnas numéricas
numeric_cols = df.select_dtypes(include=[np.number]).columns
correlation_matrix = df[numeric_cols].corr()


mask = np.triu(np.ones_like(correlation_matrix, dtype=bool))
sns.heatmap(correlation_matrix, mask=mask, annot=True, cmap='coolwarm', 
            center=0, square=True, linewidths=0.5, cbar_kws={"shrink": .8},
            fmt='.2f', annot_kws={'size': 10, 'weight': 'bold'})

plt.title("Matriz de Correlación - Variables Numéricas", fontsize=18, fontweight='bold', pad=20)
plt.xticks(rotation=45, ha='right')
plt.yticks(rotation=0)
plt.tight_layout()
plt.savefig("matriz_correlacion.png", dpi=300, bbox_inches='tight')
plt.show()

