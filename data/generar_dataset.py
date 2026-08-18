"""
Generador de dataset sintético — Minimarket "La Esquina"
Simula exportaciones diarias de POS (una por tienda) durante 61 días
continuos (junio-julio 2026), con precios fijos por tienda-producto y
errores intencionales (duplicados, texto sucio, cantidad/precio
inválidos, fechas fuera de rango) para practicar validación y
recuperación de datos en Power Query / VBA.

Solo usa librería estándar de Python 3 — no requiere instalar nada.
"""
import os
import csv
import random
from datetime import datetime, timedelta

random.seed(42)

# ---------------------------------------------------------------------------
# CONFIGURACIÓN
# ---------------------------------------------------------------------------
TIENDAS = {
    "Centro":    {"baseline": 55, "weekend_boost": 1.25},
    "Norte":     {"baseline": 40, "weekend_boost": 1.20},
    "Sur":       {"baseline": 35, "weekend_boost": 1.15},
    "Autopista": {"baseline": 65, "weekend_boost": 1.40},
}

# codigo -> (nombre, categoria, precio_min, precio_max)
PRODUCTOS = {
    "BEB001": ("Coca Cola 1.5L", "Bebidas", 2200, 2600),
    "BEB002": ("Agua Cristal 600ml", "Bebidas", 900, 1200),
    "BEB003": ("Jugo Hit Naranja 1L", "Bebidas", 2500, 2900),
    "BEB004": ("Pony Malta 750ml", "Bebidas", 2800, 3200),
    "BEB005": ("Gatorade 500ml", "Bebidas", 3200, 3600),
    "BEB006": ("Cerveza Aguila 330ml", "Bebidas", 2100, 2400),
    "SNK001": ("Papas Margarita 30g", "Snacks", 1500, 1800),
    "SNK002": ("Doritos Nacho 55g", "Snacks", 2500, 2900),
    "SNK003": ("Galletas Festival", "Snacks", 1300, 1600),
    "SNK004": ("Chocolatina Jet", "Snacks", 900, 1100),
    "SNK005": ("Mani Moto 40g", "Snacks", 1200, 1500),
    "SNK006": ("Bon Bon Bum x5", "Snacks", 1000, 1300),
    "LAC001": ("Leche Alpina 1L", "Lacteos", 3500, 3900),
    "LAC002": ("Yogurt Alpina 200g", "Lacteos", 1800, 2100),
    "LAC003": ("Queso Campesino 250g", "Lacteos", 6500, 7200),
    "LAC004": ("Kumis Alqueria 1L", "Lacteos", 4200, 4600),
    "LAC005": ("Arequipe Alpina 250g", "Lacteos", 5200, 5700),
    "PAN001": ("Pan Tajado Bimbo", "Panaderia", 4800, 5300),
    "PAN002": ("Pan Aliado x6", "Panaderia", 3200, 3700),
    "PAN003": ("Croissant", "Panaderia", 1800, 2200),
    "PAN004": ("Ponque Ramo", "Panaderia", 2600, 3000),
    "ASE001": ("Jabon Protex", "Aseo Personal", 2800, 3200),
    "ASE002": ("Shampoo Head&Shoulders 375ml", "Aseo Personal", 12500, 14000),
    "ASE003": ("Crema Dental Colgate", "Aseo Personal", 4200, 4800),
    "ASE004": ("Papel Higienico Familia x4", "Aseo Personal", 6800, 7500),
    "ASE005": ("Desodorante Rexona", "Aseo Personal", 7500, 8300),
    "LIM001": ("Detergente Fab 500g", "Limpieza Hogar", 4500, 5000),
    "LIM002": ("Limpido 1L", "Limpieza Hogar", 3800, 4200),
    "LIM003": ("Escoba Plastica", "Limpieza Hogar", 9500, 10500),
    "LIM004": ("Esponjilla Brillo x3", "Limpieza Hogar", 2200, 2600),
    "LIM005": ("Bolsas Basura x10", "Limpieza Hogar", 3200, 3600),
    "ABA001": ("Arroz Diana 1kg", "Abarrotes", 3200, 3600),
    "ABA002": ("Aceite Girasol Premier 1L", "Abarrotes", 8500, 9400),
    "ABA003": ("Azucar Manuelita 1kg", "Abarrotes", 2900, 3300),
    "ABA004": ("Sal Refisal 500g", "Abarrotes", 1200, 1500),
    "ABA005": ("Panela Cuadrada x6", "Abarrotes", 3800, 4200),
    "ABA006": ("Cafe Sello Rojo 250g", "Abarrotes", 6200, 6900),
    "ABA007": ("Lentejas 500g", "Abarrotes", 2800, 3200),
    "ABA008": ("Pasta La Muñeca 500g", "Abarrotes", 2100, 2500),
    "ABA009": ("Huevos AA x30", "Abarrotes", 13500, 15000),
    "ABA010": ("Chocolate Corona 500g", "Abarrotes", 7800, 8600),
}

METODOS_PAGO = ["Efectivo", "Tarjeta Debito", "Tarjeta Credito", "Nequi/Daviplata"]

# ---------------------------------------------------------------------------
# PRECIOS FIJOS por tienda-producto (se calculan UNA sola vez, no por transacción)
# Cada tienda tiene un factor de ajuste propio (competencia local, costos de
# transporte, etc.) pero dentro de una misma tienda el precio de un producto
# NO cambia entre transacciones ni entre días del periodo.
# ---------------------------------------------------------------------------
FACTOR_PRECIO_TIENDA = {
    "Centro": 1.00,
    "Norte": 0.97,
    "Sur": 0.95,
    "Autopista": 1.06,   # más cara: ubicación de conveniencia, alto tráfico
}

PRECIO_BASE = {}      # codigo -> precio base (punto medio del rango original)
PRECIO_TIENDA = {}     # (tienda, codigo) -> precio fijo para esa tienda
for _codigo, (_nombre, _categoria, _pmin, _pmax) in PRODUCTOS.items():
    base = round((_pmin + _pmax) / 2 / 50) * 50  # redondeado a multiplos de 50
    PRECIO_BASE[_codigo] = base
    for _tienda, _factor in FACTOR_PRECIO_TIENDA.items():
        PRECIO_TIENDA[(_tienda, _codigo)] = round(base * _factor / 50) * 50

FECHA_INICIO = datetime(2026, 6, 1)
DIAS = 61  # junio (30) + julio (31) completos y continuos, sin huecos
HORA_APERTURA = 8
HORA_CIERRE = 20  # exclusivo (última venta puede ser 19:xx)

OUT_DIR = "/home/claude/dataset_minimarket"
os.makedirs(OUT_DIR, exist_ok=True)

# Errores a inyectar (proporciones sobre el total de filas del archivo)
PCT_DUPLICADOS = 0.02
PCT_CANTIDAD_PRECIO_INVALIDO = 0.012
PCT_TIENDA_MAL_ESCRITA = 0.01
PCT_FECHA_FUERA_RANGO = 0.006
PCT_PRODUCTO_SUCIO = 0.02

TIENDA_TYPOS = {
    "Centro": ["centro ", "CENTRO", "Centro "],
    "Norte": ["norte", "Nrote", "Norte  "],
    "Sur": ["sur ", "SUR", "Su r"],
    "Autopista": ["autopista", "Auto pista", "AUTOPISTA "],
}


def hora_ponderada():
    """Más tráfico al mediodía y en la tarde-noche (hora pico de cierre de jornada)."""
    pesos_hora = {
        8: 3, 9: 4, 10: 5, 11: 6, 12: 9, 13: 8, 14: 6,
        15: 5, 16: 6, 17: 7, 18: 9, 19: 8,
    }
    horas = list(pesos_hora.keys())
    pesos = list(pesos_hora.values())
    h = random.choices(horas, weights=pesos, k=1)[0]
    m = random.randint(0, 59)
    s = random.randint(0, 59)
    return f"{h:02d}:{m:02d}:{s:02d}"


def generar_transaccion(idx_global, fecha, tienda):
    codigo = random.choice(list(PRODUCTOS.keys()))
    nombre, categoria, pmin, pmax = PRODUCTOS[codigo]
    cantidad = random.choices([1, 2, 3, 4, 5], weights=[45, 25, 15, 10, 5])[0]
    precio = PRECIO_TIENDA[(tienda, codigo)]  # precio fijo tienda-producto, no aleatorio
    total = cantidad * precio
    hora = hora_ponderada()
    metodo = random.choices(
        METODOS_PAGO, weights=[45, 25, 20, 10], k=1
    )[0]
    return {
        "ID_Transaccion": f"TXN-{fecha.strftime('%Y%m%d')}-{tienda[:2].upper()}-{idx_global:05d}",
        "Fecha": fecha.strftime("%Y-%m-%d"),
        "Hora": hora,
        "Tienda": tienda,
        "Codigo_Producto": codigo,
        "Producto": nombre,
        "Categoria": categoria,
        "Cantidad": cantidad,
        "Precio_Unitario": precio,
        "Total": total,
        "Metodo_Pago": metodo,
    }


def inyectar_errores(filas, fecha, tienda):
    n = len(filas)

    # 1) Duplicados: repetir algunas filas completas (mismo ID_Transaccion)
    n_dup = max(1, int(n * PCT_DUPLICADOS))
    for _ in range(n_dup):
        origen = random.choice(filas).copy()
        filas.append(origen)

    # 2) Cantidad/Precio inválidos (negativos o vacíos)
    n_inv = max(1, int(n * PCT_CANTIDAD_PRECIO_INVALIDO))
    for _ in range(n_inv):
        fila = random.choice(filas)
        if random.random() < 0.5:
            fila["Cantidad"] = -abs(fila["Cantidad"])
        else:
            campo = random.choice(["Cantidad", "Precio_Unitario"])
            fila[campo] = ""

    # 3) Tienda mal escrita
    n_typo = max(1, int(n * PCT_TIENDA_MAL_ESCRITA))
    for _ in range(n_typo):
        fila = random.choice(filas)
        fila["Tienda"] = random.choice(TIENDA_TYPOS[tienda])

    # 4) Fecha fuera de rango (mes anterior o siguiente)
    n_fecha = max(1, int(n * PCT_FECHA_FUERA_RANGO))
    for _ in range(n_fecha):
        fila = random.choice(filas)
        offset_dias = random.choice([-45, -40, 35, 40])
        fecha_mala = fecha + timedelta(days=offset_dias)
        fila["Fecha"] = fecha_mala.strftime("%Y-%m-%d")

    # 5) Nombre de producto "sucio" (espacios extra / mayúsculas inconsistentes)
    n_sucio = max(1, int(n * PCT_PRODUCTO_SUCIO))
    for _ in range(n_sucio):
        fila = random.choice(filas)
        variante = random.choice([
            fila["Producto"].upper(),
            fila["Producto"].lower(),
            f"  {fila['Producto']} ",
            fila["Producto"].replace(" ", "  "),
        ])
        fila["Producto"] = variante

    random.shuffle(filas)
    return filas


COLUMNAS = [
    "ID_Transaccion", "Fecha", "Hora", "Tienda", "Codigo_Producto",
    "Producto", "Categoria", "Cantidad", "Precio_Unitario", "Total", "Metodo_Pago",
]

resumen_generacion = []
idx_global = 1

for d in range(DIAS):
    fecha = FECHA_INICIO + timedelta(days=d)
    es_finde = fecha.weekday() in (4, 5)  # viernes=4, sabado=5 (domingo tambien se mueve pero menos)
    es_domingo = fecha.weekday() == 6

    carpeta_dia = os.path.join(OUT_DIR, f"dia_{fecha.strftime('%Y-%m-%d')}")
    os.makedirs(carpeta_dia, exist_ok=True)

    for tienda, cfg in TIENDAS.items():
        base = cfg["baseline"]
        if es_finde:
            base = int(base * cfg["weekend_boost"])
        elif es_domingo:
            base = int(base * 0.85)

        n_transacciones = max(5, int(random.gauss(base, base * 0.12)))

        filas = []
        for _ in range(n_transacciones):
            filas.append(generar_transaccion(idx_global, fecha, tienda))
            idx_global += 1

        filas = inyectar_errores(filas, fecha, tienda)

        nombre_archivo = f"POS_{tienda}_{fecha.strftime('%Y-%m-%d')}.csv"
        ruta = os.path.join(carpeta_dia, nombre_archivo)
        with open(ruta, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.DictWriter(f, fieldnames=COLUMNAS)
            writer.writeheader()
            for fila in filas:
                writer.writerow(fila)

        resumen_generacion.append((fecha.strftime("%Y-%m-%d"), tienda, len(filas)))

print(f"Total de archivos generados: {DIAS * len(TIENDAS)}")
print(f"Total de filas (transacciones) generadas: {sum(r[2] for r in resumen_generacion)}")
print("Ejemplo de resumen (primeros 8):")
for r in resumen_generacion[:8]:
    print(r)
