from flask import Flask, request, render_template_string, make_response
import pandas as pd
import seaborn as sns
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import io
import base64
import os
import html

# ==================================================================
# DASHBOARD MITIENDA - VERSION LOCAL CON EXCEL
# Coloca "LaTiendaCom (1).xls" junto a este archivo.
# Ejecuta: py dashboard.py
# Abre: http://127.0.0.1:5000
# ==================================================================

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ARCHIVO_EXCEL = os.path.join(BASE_DIR, "LaTiendaCom (1).xls")

app = Flask(__name__)
sns.set_theme(style="whitegrid", palette="deep")


def cargar_excel():
    if not os.path.exists(ARCHIVO_EXCEL):
        raise FileNotFoundError(
            f"No se encontró el archivo: {ARCHIVO_EXCEL}. "
            "Coloca LaTiendaCom (1).xls en la misma carpeta que dashboard.py."
        )

    xls = pd.ExcelFile(ARCHIVO_EXCEL)
    hojas = {nombre: pd.read_excel(ARCHIVO_EXCEL, sheet_name=nombre)
             for nombre in xls.sheet_names}

    # Normalizar fechas y números conocidos
    for hoja, columna in [
        ("Facturas", "fecha_factura"),
        ("Pedidos", "fecha_pedido"),
        ("Clientes", "fecha_creacion"),
        ("Empleados", "fecha_contratacion"),
        ("Productos", "fecha_vencimiento"),
        ("Proveedores", "fecha_servicios"),
    ]:
        if hoja in hojas and columna in hojas[hoja].columns:
            hojas[hoja][columna] = pd.to_datetime(
                hojas[hoja][columna], errors="coerce"
            )

    if "Facturas" in hojas:
        facturas = hojas["Facturas"]
        for col in ["cantidad_productos", "precio_unitario"]:
            if col in facturas.columns:
                facturas[col] = pd.to_numeric(facturas[col], errors="coerce")
        if {"cantidad_productos", "precio_unitario"}.issubset(facturas.columns):
            facturas["importe_estimado"] = (
                facturas["cantidad_productos"].fillna(0)
                * facturas["precio_unitario"].fillna(0)
            )
        else:
            facturas["importe_estimado"] = 0

    if "Productos" in hojas:
        for col in ["precio", "stock"]:
            if col in hojas["Productos"].columns:
                hojas["Productos"][col] = pd.to_numeric(
                    hojas["Productos"][col], errors="coerce"
                )

    if "Empleados" in hojas and "salario" in hojas["Empleados"].columns:
        hojas["Empleados"]["salario"] = pd.to_numeric(
            hojas["Empleados"]["salario"], errors="coerce"
        )

    return hojas


def dinero(valor):
    try:
        return "$ {:,.2f}".format(float(valor)).replace(",", "X").replace(".", ",").replace("X", ".")
    except (ValueError, TypeError):
        return "$ 0,00"


def grafico_png(fig):
    buffer = io.BytesIO()
    fig.tight_layout()
    fig.savefig(buffer, format="png", dpi=130, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    buffer.seek(0)
    return base64.b64encode(buffer.read()).decode("utf-8")


def hacer_graficos(facturas, productos, pedidos, clientes, empleados, proveedores):
    graficos = []

    # 1. Importe estimado por mes
    if not facturas.empty and "fecha_factura" in facturas and "importe_estimado" in facturas:
        df = facturas.dropna(subset=["fecha_factura"]).copy()
        if not df.empty:
            df["mes"] = df["fecha_factura"].dt.to_period("M").astype(str)
            resumen = df.groupby("mes", as_index=False)["importe_estimado"].sum().tail(18)
            fig, ax = plt.subplots(figsize=(8, 3.4))
            sns.lineplot(data=resumen, x="mes", y="importe_estimado", marker="o", ax=ax)
            ax.set_title("Importe estimado por mes")
            ax.set_xlabel("Mes")
            ax.set_ylabel("Importe estimado")
            ax.tick_params(axis="x", rotation=45)
            graficos.append(("Evolución mensual", grafico_png(fig)))

    # 2. Productos con mayor stock
    if not productos.empty and {"nombre_producto", "stock"}.issubset(productos.columns):
        df = productos.dropna(subset=["nombre_producto", "stock"]).nlargest(10, "stock")
        if not df.empty:
            fig, ax = plt.subplots(figsize=(8, 3.4))
            sns.barplot(data=df, y="nombre_producto", x="stock", ax=ax)
            ax.set_title("10 productos con mayor stock")
            ax.set_xlabel("Unidades en stock")
            ax.set_ylabel("")
            graficos.append(("Inventario", grafico_png(fig)))

    # 3. Distribución de precios
    if not productos.empty and "precio" in productos.columns:
        df = productos.dropna(subset=["precio"])
        if not df.empty:
            fig, ax = plt.subplots(figsize=(8, 3.4))
            sns.histplot(data=df, x="precio", bins=24, kde=True, ax=ax)
            ax.set_title("Distribución de precios de productos")
            ax.set_xlabel("Precio")
            ax.set_ylabel("Cantidad de productos")
            graficos.append(("Precios", grafico_png(fig)))

    # 4. Métodos de pago
    if not facturas.empty and "forma_de_pago" in facturas.columns:
        df = facturas["forma_de_pago"].fillna("Sin dato").value_counts().rename_axis("forma_de_pago").reset_index(name="facturas")
        if not df.empty:
            fig, ax = plt.subplots(figsize=(8, 3.4))
            sns.barplot(data=df, x="forma_de_pago", y="facturas", ax=ax)
            ax.set_title("Facturas por forma de pago")
            ax.set_xlabel("Forma de pago")
            ax.set_ylabel("Número de facturas")
            ax.tick_params(axis="x", rotation=20)
            graficos.append(("Formas de pago", grafico_png(fig)))

    # 5. Tipo de venta
    if not facturas.empty and "tipo_de_venta" in facturas.columns:
        df = facturas["tipo_de_venta"].fillna("Sin dato").value_counts().rename_axis("tipo_de_venta").reset_index(name="facturas")
        if not df.empty:
            fig, ax = plt.subplots(figsize=(8, 3.4))
            sns.barplot(data=df, x="tipo_de_venta", y="facturas", ax=ax)
            ax.set_title("Facturas por tipo de venta")
            ax.set_xlabel("Tipo de venta")
            ax.set_ylabel("Número de facturas")
            graficos.append(("Tipo de venta", grafico_png(fig)))

    # 6. Pedidos por año
    if not pedidos.empty and "fecha_pedido" in pedidos.columns:
        df = pedidos.dropna(subset=["fecha_pedido"]).copy()
        if not df.empty:
            df["año"] = df["fecha_pedido"].dt.year.astype(str)
            resumen = df.groupby("año", as_index=False).size().rename(columns={"size": "pedidos"})
            fig, ax = plt.subplots(figsize=(8, 3.4))
            sns.barplot(data=resumen, x="año", y="pedidos", ax=ax)
            ax.set_title("Pedidos registrados por año")
            ax.set_xlabel("Año")
            ax.set_ylabel("Cantidad de pedidos")
            ax.tick_params(axis="x", rotation=45)
            graficos.append(("Pedidos por año", grafico_png(fig)))

    # 7. Clientes por fecha de creación
    if not clientes.empty and "fecha_creacion" in clientes.columns:
        df = clientes.dropna(subset=["fecha_creacion"]).copy()
        if not df.empty:
            df["año"] = df["fecha_creacion"].dt.year.astype(str)
            resumen = df.groupby("año", as_index=False).size().rename(columns={"size": "clientes"})
            resumen = resumen.tail(18)
            fig, ax = plt.subplots(figsize=(8, 3.4))
            sns.lineplot(data=resumen, x="año", y="clientes", marker="o", ax=ax)
            ax.set_title("Clientes creados por año")
            ax.set_xlabel("Año")
            ax.set_ylabel("Clientes")
            ax.tick_params(axis="x", rotation=45)
            graficos.append(("Clientes", grafico_png(fig)))

    # 8. Salarios por cargo
    if not empleados.empty and {"cargo", "salario"}.issubset(empleados.columns):
        df = empleados.dropna(subset=["cargo", "salario"])
        if not df.empty:
            resumen = df.groupby("cargo", as_index=False)["salario"].mean().nlargest(10, "salario")
            fig, ax = plt.subplots(figsize=(8, 3.4))
            sns.barplot(data=resumen, y="cargo", x="salario", ax=ax)
            ax.set_title("Salario promedio por cargo (máximo 10 cargos)")
            ax.set_xlabel("Salario promedio")
            ax.set_ylabel("")
            graficos.append(("Empleados", grafico_png(fig)))

    # 9. Proveedores por año de servicio
    if not proveedores.empty and "fecha_servicios" in proveedores.columns:
        df = proveedores.dropna(subset=["fecha_servicios"]).copy()
        if not df.empty:
            df["año"] = df["fecha_servicios"].dt.year.astype(str)
            resumen = df.groupby("año", as_index=False).size().rename(columns={"size": "proveedores"})
            resumen = resumen.tail(18)
            fig, ax = plt.subplots(figsize=(8, 3.4))
            sns.barplot(data=resumen, x="año", y="proveedores", ax=ax)
            ax.set_title("Proveedores por año de registro de servicios")
            ax.set_xlabel("Año")
            ax.set_ylabel("Proveedores")
            ax.tick_params(axis="x", rotation=45)
            graficos.append(("Proveedores", grafico_png(fig)))

    return graficos


HTML = r"""
<!doctype html>
<html lang="es">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>MiTienda | Dashboard</title>
<style>
:root { --bg:#f3f6fb; --panel:#fff; --ink:#172033; --muted:#68758b; --accent:#3659e3; --line:#e5eaf3; }
* { box-sizing:border-box; }
body { margin:0; font-family:Inter,Segoe UI,Arial,sans-serif; background:var(--bg); color:var(--ink); }
header { background:linear-gradient(120deg,#172554,#263fa9 65%,#4773e8); color:white; padding:28px clamp(18px,4vw,54px); }
.eyebrow { font-size:12px; letter-spacing:2px; text-transform:uppercase; opacity:.8; }
h1 { margin:8px 0; font-size:clamp(25px,3vw,38px); }
header p { margin:0; opacity:.85; }
main { max-width:1500px; margin:auto; padding:24px clamp(14px,3vw,38px) 42px; }
.toolbar { background:var(--panel); border:1px solid var(--line); border-radius:16px; padding:18px; display:flex; gap:14px; flex-wrap:wrap; align-items:end; box-shadow:0 5px 20px #20305008; }
.field { display:flex; flex-direction:column; gap:6px; min-width:150px; flex:1; }
label { font-size:12px; color:var(--muted); font-weight:700; }
input,select { border:1px solid #d7deeb; background:white; border-radius:9px; padding:10px; color:var(--ink); min-height:40px; }
button,.button { background:var(--accent); border:0; color:white; border-radius:9px; padding:11px 16px; font-weight:700; text-decoration:none; cursor:pointer; display:inline-block; }
.secondary { background:#eef2ff; color:#2944b8; }
.kpis { display:grid; grid-template-columns:repeat(4,minmax(0,1fr)); gap:15px; margin:20px 0; }
.kpi { background:var(--panel); border:1px solid var(--line); border-radius:15px; padding:20px; box-shadow:0 5px 20px #20305008; }
.kpi .label { color:var(--muted); font-size:13px; }
.kpi .value { font-size:clamp(21px,2vw,29px); font-weight:800; margin:9px 0 5px; overflow-wrap:anywhere; }
.kpi .note { font-size:12px; color:var(--muted); }
.section-title { margin:28px 0 14px; display:flex; justify-content:space-between; align-items:end; gap:12px; }
.section-title h2 { margin:0; font-size:21px; }
.section-title span { color:var(--muted); font-size:12px; }
.charts { display:grid; grid-template-columns:repeat(2,minmax(0,1fr)); gap:16px; }
.chart { background:var(--panel); border:1px solid var(--line); border-radius:15px; padding:16px; min-width:0; box-shadow:0 5px 20px #20305008; }
.chart h3 { font-size:15px; margin:0 0 12px; }
.chart img { display:block; width:100%; height:auto; }
.table-wrap { overflow:auto; background:white; border:1px solid var(--line); border-radius:14px; }
table { border-collapse:collapse; width:100%; font-size:13px; }
th,td { padding:11px 13px; text-align:left; border-bottom:1px solid var(--line); white-space:nowrap; }
th { background:#f8faff; color:#536078; position:sticky; top:0; }
.notice { border-radius:12px; padding:13px 15px; background:#fff7df; color:#805d0b; margin-top:16px; font-size:13px; }
footer { color:var(--muted); font-size:12px; text-align:center; padding:28px 12px; }
@media(max-width:900px){ .kpis{grid-template-columns:repeat(2,minmax(0,1fr));}.charts{grid-template-columns:1fr;} }
@media(max-width:480px){ .kpis{grid-template-columns:1fr 1fr; gap:9px;}.kpi{padding:13px;} }
</style>
</head>
<body>
<header>
  <div class="eyebrow">Analítica de datos · MiTienda</div>
  <h1>Panel de control comercial</h1>
  <p>Resumen de facturas, pedidos, inventario, clientes y proveedores.</p>
</header>
<main>
  <form class="toolbar" method="get">
    <div class="field"><label for="desde">Fecha inicial de factura</label><input type="date" id="desde" name="desde" value="{{ desde }}"></div>
    <div class="field"><label for="hasta">Fecha final de factura</label><input type="date" id="hasta" name="hasta" value="{{ hasta }}"></div>
    <div class="field"><label for="pago">Forma de pago</label><select id="pago" name="pago"><option value="">Todas</option>{% for x in pagos %}<option value="{{ x }}" {% if x == pago %}selected{% endif %}>{{ x }}</option>{% endfor %}</select></div>
    <div class="field"><label for="tipo">Tipo de venta</label><select id="tipo" name="tipo"><option value="">Todos</option>{% for x in tipos %}<option value="{{ x }}" {% if x == tipo %}selected{% endif %}>{{ x }}</option>{% endfor %}</select></div>
    <button type="submit">Aplicar filtros</button>
    <a class="button secondary" href="/">Limpiar</a>
  </form>

  <div class="kpis">
    <div class="kpi"><div class="label">Facturas filtradas</div><div class="value">{{ "{:,}".format(n_facturas) }}</div><div class="note">Registros del periodo y filtros seleccionados</div></div>
    <div class="kpi"><div class="label">Importe estimado</div><div class="value">{{ total_dinero }}</div><div class="note">Cantidad × precio unitario; no es total contable verificado</div></div>
    <div class="kpi"><div class="label">Unidades facturadas</div><div class="value">{{ "{:,}".format(unidades) }}</div><div class="note">Suma de cantidad_productos</div></div>
    <div class="kpi"><div class="label">Productos en catálogo</div><div class="value">{{ "{:,}".format(n_productos) }}</div><div class="note">Productos únicos por id_producto</div></div>
  </div>

  <div class="section-title"><h2>Gráficos de análisis</h2><span>{{ n_graficos }} visualizaciones · Seaborn / Matplotlib</span></div>
  <div class="charts">
  {% for titulo, imagen in graficos %}
    <section class="chart"><h3>{{ titulo }}</h3><img src="data:image/png;base64,{{ imagen }}" alt="{{ titulo }}"></section>
  {% endfor %}
  </div>

  <div class="section-title"><h2>Detalle de facturas</h2><span>Se muestran hasta 100 registros</span></div>
  <div class="table-wrap"><table>
    <thead><tr>{% for c in columnas_tabla %}<th>{{ c }}</th>{% endfor %}</tr></thead>
    <tbody>{% for fila in filas_tabla %}<tr>{% for celda in fila %}<td>{{ celda }}</td>{% endfor %}</tr>{% else %}<tr><td colspan="{{ columnas_tabla|length }}">No hay registros para los filtros seleccionados.</td></tr>{% endfor %}</tbody>
  </table></div>

  <div class="notice"><strong>Nota de interpretación:</strong> el archivo contiene datos de ejemplo y fechas históricas. El importe estimado se calcula con los campos disponibles; confirma con la guía de la actividad si el precio unitario representa precio por unidad o importe de línea.</div>
  <footer>MiTienda · Dashboard local · Los datos se leen desde el archivo Excel de este equipo.</footer>
</main>
</body></html>
"""


@app.route("/")
def inicio():
    try:
        hojas = cargar_excel()
    except Exception as exc:
        return (
            "<h2>No se pudo cargar el Excel</h2><p>"
            + html.escape(str(exc))
            + "</p><p>Verifica que LaTiendaCom (1).xls esté junto a dashboard.py.</p>",
            500,
        )

    facturas = hojas.get("Facturas", pd.DataFrame()).copy()
    productos = hojas.get("Productos", pd.DataFrame()).copy()
    pedidos = hojas.get("Pedidos", pd.DataFrame()).copy()
    clientes = hojas.get("Clientes", pd.DataFrame()).copy()
    empleados = hojas.get("Empleados", pd.DataFrame()).copy()
    proveedores = hojas.get("Proveedores", pd.DataFrame()).copy()

    desde = request.args.get("desde", "").strip()
    hasta = request.args.get("hasta", "").strip()
    pago = request.args.get("pago", "").strip()
    tipo = request.args.get("tipo", "").strip()

    facturas_filtradas = facturas.copy()
    if "fecha_factura" in facturas_filtradas.columns:
        facturas_filtradas["fecha_factura"] = pd.to_datetime(
            facturas_filtradas["fecha_factura"], errors="coerce"
        )
        if desde:
            fecha_desde = pd.to_datetime(desde, errors="coerce")
            if pd.notna(fecha_desde):
                facturas_filtradas = facturas_filtradas[
                    facturas_filtradas["fecha_factura"] >= fecha_desde
                ]
        if hasta:
            fecha_hasta = pd.to_datetime(hasta, errors="coerce")
            if pd.notna(fecha_hasta):
                facturas_filtradas = facturas_filtradas[
                    facturas_filtradas["fecha_factura"] < fecha_hasta + pd.Timedelta(days=1)
                ]

    if pago and "forma_de_pago" in facturas_filtradas.columns:
        facturas_filtradas = facturas_filtradas[
            facturas_filtradas["forma_de_pago"].astype(str) == pago
        ]
    if tipo and "tipo_de_venta" in facturas_filtradas.columns:
        facturas_filtradas = facturas_filtradas[
            facturas_filtradas["tipo_de_venta"].astype(str) == tipo
        ]

    pagos = sorted(facturas["forma_de_pago"].dropna().astype(str).unique().tolist()) if "forma_de_pago" in facturas else []
    tipos = sorted(facturas["tipo_de_venta"].dropna().astype(str).unique().tolist()) if "tipo_de_venta" in facturas else []

    total = facturas_filtradas["importe_estimado"].fillna(0).sum() if "importe_estimado" in facturas_filtradas else 0
    unidades = int(facturas_filtradas["cantidad_productos"].fillna(0).sum()) if "cantidad_productos" in facturas_filtradas else 0
    n_productos = productos["id_producto"].nunique() if "id_producto" in productos else len(productos)

    graficos = hacer_graficos(
        facturas_filtradas, productos, pedidos, clientes, empleados, proveedores
    )

    columnas_preferidas = [
        "id_factura", "fecha_factura", "id_cliente", "id_producto",
        "cantidad_productos", "precio_unitario", "importe_estimado",
        "forma_de_pago", "tipo_de_venta"
    ]
    columnas_tabla = [c for c in columnas_preferidas if c in facturas_filtradas.columns]
    tabla = facturas_filtradas[columnas_tabla].copy() if columnas_tabla else facturas_filtradas.head(0)
    if "fecha_factura" in tabla.columns:
        tabla["fecha_factura"] = pd.to_datetime(tabla["fecha_factura"], errors="coerce").dt.strftime("%Y-%m-%d")
    if "importe_estimado" in tabla.columns:
        tabla["importe_estimado"] = tabla["importe_estimado"].apply(dinero)
    filas_tabla = tabla.head(100).fillna("").astype(str).values.tolist()

    return render_template_string(
        HTML,
        desde=desde, hasta=hasta, pago=pago, tipo=tipo, pagos=pagos, tipos=tipos,
        n_facturas=len(facturas_filtradas), total_dinero=dinero(total),
        unidades=unidades, n_productos=n_productos, graficos=graficos,
        n_graficos=len(graficos), columnas_tabla=columnas_tabla, filas_tabla=filas_tabla
    )


@app.route("/exportar-html")
def exportar_html():
    # Guarda una copia HTML autocontenida de la vista actual con los filtros aplicados.
    with app.test_client() as cliente:
        respuesta = cliente.get("/?" + request.query_string.decode("utf-8"))
    contenido = respuesta.get_data(as_text=True)
    archivo = make_response(contenido)
    archivo.headers["Content-Type"] = "text/html; charset=utf-8"
    archivo.headers["Content-Disposition"] = 'attachment; filename="dashboard_mitienda.html"'
    return archivo


if __name__ == "__main__":
    print("\nDashboard MiTienda listo.")
    print("Abre en tu navegador: http://127.0.0.1:5000")
    print("Para exportar HTML: http://127.0.0.1:5000/exportar-html")
    app.run(host="127.0.0.1", port=5000, debug=False)
