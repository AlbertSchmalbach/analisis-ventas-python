import pandas as pd
import matplotlib.pyplot as plt

# Cargar archivo Excel
df = pd.read_excel("Dashboard_Ventas_Alberto.xlsx", sheet_name="Datos")

# Mostrar primeras filas
# print(df.head())

# Crear columna Total
df["Total"] = df["Cantidad"] * df["Precio"]

print("\nDatos con Total:")
print(df.head(12))

# KPI 1: Total ventas
total_ventas = df["Total"].sum()

# KPI 2: Producto más vendido
producto_top = df.groupby("Producto")["Total"].sum().idxmax()

# KPI 3: Mejor vendedor
mejor_vendedor = df.groupby("Vendedor")["Total"].sum().idxmax()

print("\n===== KPI =====")
print("Total ventas:", total_ventas)
print("Producto más vendido:", producto_top)
print("Mejor vendedor:", mejor_vendedor)

# Agrupar ventas por producto
ventas_producto = df.groupby("Producto")["Total"].sum()

# Crear gráfico de barra
ventas_producto.plot(kind="bar")

plt.title("Ventas por Producto")
plt.xlabel("Producto")
plt.ylabel("Total ventas")
plt.xticks(rotation=45)


plt.tight_layout()
plt.show()

# Agrupar ventas por vendedor
ventas_vendedor = df.groupby("Vendedor")["Total"].sum()

# Crear gráfico de torta
ventas_vendedor.plot(
    kind="pie",
    autopct='%1.1f%%',    # Muestra el porcentaje con un decimal
    startangle=90,        # Rota el inicio del gráfico
    ylabel="",            # Elimina la etiqueta "Total" del eje lateral
    title="Distribución de Ventas por Vendedor",
    figsize=(8, 8)        # Ajusta el tamaño para que sea circular
)

plt.tight_layout()
plt.show()

# Exportar resumen a Excel
resumen = df.groupby("Producto")["Total"].sum().reset_index()

resumen.to_excel("Reporte_Ventas.xlsx", index=False)

print("\nReporte generado: reporte_ventas.xlsx")