# =============================================================================
#  EDITOR — revisar, corregir y confirmar certificaciones en Google Colab
# =============================================================================
import json
import shutil
from datetime import date, datetime
from pathlib import Path

import pandas as pd
import ipywidgets as W
from IPython.display import display

import cert_biomasa as cb

try:
    from google.colab import files as _colab_files
except ImportError:
    _colab_files = None

ANCHOS = {"del": "34px", "Batea": "90px", "Fecha": "150px", "Rto. Planta": "140px",
          "Remito Proveedor": "150px", "Producto": "120px", "Kilos": "105px", "Finos": "80px", "Hum": "80px", "DescMan": "80px", "PenMan": "80px",
          "Desc": "130px", "Punto": "150px", "obs": "40px"}


def _L(w, **kw):
    return W.Layout(width=w, **kw)


class Selector:
    """Paso de elección: año, mes, períodos (1Q / 2Q / Mensual por separado) y proveedor,
    con la lista de proveedores tomada del maestro para el mes elegido."""

    def __init__(self, maestro, conf=None, anio=None, mes=None):
        self.d, _ = cb.preparar(maestro, conf)
        hoy = date.today()
        ult = self.d[cb.C_DIA].max() - pd.Timedelta(days=10)   # último mes completo
        anio = anio or ult.year
        mes = mes or ult.month
        anios = sorted(int(a) for a in self.d[cb.C_DIA].dt.year.unique())
        st = {"description_width": "90px"}
        self.w_anio = W.Dropdown(options=anios, value=anio if anio in anios else anios[-1], description="Año:",
                                 style=st, layout=_L("200px"))
        self.w_mes = W.Dropdown(options=[(cb.MESES[m], m) for m in range(1, 13)], value=mes, description="Mes:",
                                style=st, layout=_L("240px"))
        self.w_1q = W.Checkbox(value=True, indent=False, description=f"1ª quincena (1 al {cb.DIA_CORTE})", layout=_L("220px"))
        self.w_2q = W.Checkbox(value=True, indent=False, description=f"2ª quincena ({cb.DIA_CORTE + 1} a fin de mes)", layout=_L("250px"))
        self.w_men = W.Checkbox(value=True, indent=False, description="Mensual (mes completo)", layout=_L("220px"))
        self.w_corte = W.BoundedIntText(value=cb.DIA_CORTE, min=10, max=20, description="Último día de la 1ª quincena:",
                                        style={"description_width": "200px"}, layout=_L("280px"))
        self.w_corte.observe(lambda ch: self._etiquetas(), "value")
        self.w_buscar = W.Text(placeholder="Buscar proveedor…", description="Proveedor:", style=st, layout=_L("330px"))
        self.w_prov = W.Select(description="", rows=12, layout=_L("620px"))
        self.w_info = W.HTML()
        for w in (self.w_anio, self.w_mes, self.w_buscar):
            w.observe(lambda ch: self._listar(), "value")
        for w in (self.w_prov, self.w_1q, self.w_2q, self.w_men):
            w.observe(lambda ch: self._info(), "value")
        self._listar()

    def _etiquetas(self):
        dc = self.w_corte.value
        self.w_1q.description = f"1ª quincena (1 al {dc})"
        self.w_2q.description = f"2ª quincena ({dc + 1} a fin de mes)"
        self._info()

    @property
    def corte(self):
        return int(self.w_corte.value)

    def _listar(self):
        a, m = self.w_anio.value, self.w_mes.value
        x = self.d[(self.d[cb.C_DIA].dt.year == a) & (self.d[cb.C_DIA].dt.month == m)]
        g = x.groupby(cb.C_PROV)[cb.C_NETO].agg(["size", "sum"]).sort_index(key=lambda s: s.str.casefold())
        q = self.w_buscar.value.strip().casefold()
        ops = [(f"— Todos los proveedores ({len(g)}) —", "__TODOS__")]
        tipos = (x.groupby([cb.C_PROV, cb.C_TIPO])[cb.C_NETO].sum() / 1000).sort_values(ascending=False)
        txt_t = lambda p: " / ".join(cb.nombre_producto(t) for t in tipos.loc[p].index) if p in tipos.index else ""
        ops += [(f"{p}   ·   {int(r['size'])} remitos · {cb.decimal(r['sum'] / 1000)} TN · {txt_t(p)}", p)
                for p, r in g.iterrows() if q in p.casefold()]
        actual = self.w_prov.value
        self.w_prov.options = ops
        if q and len(ops) > 1 and actual in (None, "__TODOS__"):
            self.w_prov.value = ops[1][1]
        else:
            self.w_prov.value = actual if actual in [v for _, v in ops] else "__TODOS__"
        self._info()

    @property
    def periodos(self):
        return [k for k, w in (("1Q", self.w_1q), ("2Q", self.w_2q), ("Mensual", self.w_men)) if w.value]

    @property
    def proveedores(self):
        return None if self.w_prov.value in (None, "__TODOS__") else [self.w_prov.value]

    def _info(self):
        per = self.periodos
        quien = "todos los proveedores" if self.proveedores is None else f"<b>{self.w_prov.value}</b>"
        if not per:
            self.w_info.value = "<b style=color:red>Marcá al menos un período.</b>"
        else:
            self.w_info.value = (f"Se va a certificar {quien} · <b>{', '.join(per)}</b> de "
                                 f"<b>{cb.MESES[self.w_mes.value]} {self.w_anio.value}</b> "
                                 f"(1ª quincena del 1 al {self.w_corte.value}). "
                                 "Si cambiás algo acá, volvé a ejecutar el Paso 5.")
            return
            self.w_info.value = (f"Se va a certificar {quien} · <b>{', '.join(per)}</b> de "
                                 f"<b>{cb.MESES[self.w_mes.value]} {self.w_anio.value}</b>. "
                                 "Si cambiás algo acá, volvé a ejecutar el Paso 5.")

    def mostrar(self):
        display(W.VBox([W.HBox([self.w_anio, self.w_mes]),
                        W.HTML("<b>Períodos a certificar</b>"), self.w_corte, W.HBox([self.w_1q, self.w_2q, self.w_men]),
                        W.HTML("<b>Proveedor</b> (la lista sale del maestro para el mes elegido)"),
                        self.w_buscar, self.w_prov, self.w_info]))


class App:
    def __init__(self, maestro, conf, precios, par, anio, mes, tipo="ambas", salida="certificaciones",
                 logo=None, proveedores=None, estado=None, origen="", descargar_pdf=True, ruta_config=None,
                 humedad=None, corte=None):
        anio, mes = int(anio), int(mes)          # nunca tipos de numpy: se guardan en el JSON del estado
        self.raw, self.conf, self.precios, self.par = maestro, conf, precios, par
        self.metodo_humedad = str(par.get("METODO_HUMEDAD", "PUNTOS")).upper()
        self.tipo, self.logo, self.proveedores, self.origen = tipo, logo or cb.logo_por_defecto(), proveedores, origen
        self.descargar_pdf = descargar_pdf
        self.salida = Path(salida)
        self.salida.mkdir(parents=True, exist_ok=True)
        # En la PC el config vive en la carpeta principal (ruta_config); en Colab, junto a la salida
        self.ruta_config = Path(ruta_config) if ruta_config else self.salida / "config_proveedores.xlsx"
        self._respaldo_hecho = False
        cb.fijar_finos(par.get("FINOS_TOLERANCIA", cb.FINOS_TOLERANCIA), par.get("FINOS_TOTALIDAD", cb.FINOS_TOTALIDAD))
        self.d0, self.unificados = cb.preparar(maestro, conf)
        # Si se eligió un punto de entrega que se certifica junto con otros (CERTIFICAR CON), se toman todos
        if proveedores:
            unidades = {cb.unidad(conf, p) for p in proveedores}
            self.proveedores = sorted(set(proveedores) | {p for p in self.d0[cb.C_PROV].unique()
                                                          if cb.unidad(conf, p) in unidades})
        self.humedad, self.mapa_humedad = (cb.mapear_humedad(humedad, self.d0, conf) if humedad is not None
                                           else (None, None))
        self.estado = estado or cb.estado_nuevo(anio, mes, int(par["NUMERO_INICIAL"]), par["PREFIJO"])
        self.estado["anio"], self.estado["mes"] = int(self.estado["anio"]), int(self.estado["mes"])
        self.estado["corr"].setdefault("finos", {})
        for k in ("humedad_camion", "desc_gran", "pen_hum"):
            self.estado["corr"].setdefault(k, {})
        # Día de corte de la quincena: si el mes ya tiene certificaciones confirmadas se mantiene el que se usó
        pedido = int(corte if corte is not None else par.get("DIA_CORTE_QUINCENA", cb.DIA_CORTE))
        confirmadas = any(m.get("huella") for m in self.estado["meta"].values())
        if confirmadas and self.estado.get("corte") and int(self.estado["corte"]) != pedido:
            print(f"⚠ {cb.MESES[mes]} ya tiene certificaciones confirmadas con la 1ª quincena del 1 al "
                  f"{self.estado['corte']}: se mantiene ese corte (pediste {pedido}).")
            pedido = int(self.estado["corte"])
        self.estado["corte"] = pedido
        cb.fijar_corte(pedido)
        # La numeración es una sola para todos los meses: nunca usar un número ya entregado en otro mes
        self.estado["proximo"] = max(int(self.estado["proximo"]), int(par["NUMERO_INICIAL"]))
        self.ruta_estado = self.salida / f"estado_{anio}-{mes:02d}.json"
        self.key = None
        self.reconstruir()

    # ------------------------------------------------------------------ datos
    def reconstruir(self):
        self.d = cb.aplicar_correcciones(self.d0, self.estado["corr"])
        self.certs, self.globales, self.obs_cfg = cb.construir(
            self.d, self.d0, self.conf, self.precios, self.estado, self.tipo, self.raw, self.proveedores,
            self.humedad, self.metodo_humedad)
        self._validacion = cb.construir.ultima_validacion
        self.cuad = cb.cuadre(self.d0, self.d, self.certs, self.conf, self.estado["anio"],
                              self.estado["mes"], self.proveedores)

    def recalcular(self, key):
        """Rearma solo una certificación (cuando cambian número, fecha o confirmación, no los datos)."""
        nuevos, _, _ = cb.construir(self.d, self.d0, self.conf, self.precios, self.estado, self.tipo, self.raw,
                                    self.proveedores, self.humedad, self.metodo_humedad, solo={key},
                                    validacion=self._validacion)
        if key in nuevos:
            self.certs[key] = nuevos[key]

    def actualizar_humedad(self, humedad):
        """Reemplaza la planilla de calidad (humedad y granulometría) sin perder correcciones ni confirmaciones.
        Las certificaciones ya confirmadas cuyo descuento cambie pasan a MODIFICADA."""
        antes = {k: c["estado"] for k, c in self.certs.items()}
        self.humedad, self.mapa_humedad = cb.mapear_humedad(humedad, self.d0, self.conf)
        self.reconstruir()
        cambiadas = [k for k, c in self.certs.items() if antes.get(k) == "CONFIRMADA" and c["estado"] == "MODIFICADA"]
        return cambiadas

    def guardar_estado(self):
        self.par["NUMERO_INICIAL"] = max(int(self.par["NUMERO_INICIAL"]), int(self.estado["proximo"]))
        self.ruta_estado.write_text(json.dumps(self.estado, ensure_ascii=False, indent=1, default=str),
                                    encoding="utf-8")
        self._respaldar_config()
        try:
            cb.guardar_config(self.ruta_config, self.conf, self.precios, self.par)
        except PermissionError:
            raise PermissionError(f"No pude guardar {self.ruta_config.name}: está abierto en Excel. "
                                  "Cerralo y volvé a tocar Guardar (no se perdió nada).") from None

    def _respaldar_config(self, conservar=30):
        """Antes del primer cambio de cada sesión copia el config a config_respaldos/ (guarda los últimos 30)."""
        if self._respaldo_hecho or not self.ruta_config.exists():
            return
        import shutil
        carpeta = self.ruta_config.parent / "config_respaldos"
        carpeta.mkdir(exist_ok=True)
        shutil.copy2(self.ruta_config, carpeta / f"config_proveedores_{datetime.now():%Y-%m-%d_%H%M%S}.xlsx")
        for viejo in sorted(carpeta.glob("config_proveedores_*.xlsx"))[:-conservar]:
            viejo.unlink()
        self._respaldo_hecho = True

    def _log(self, accion, cert="", rem="", campo="", antes="", despues="", motivo=""):
        self.estado["log"].append([datetime.now().strftime("%d/%m/%Y %H:%M:%S"), self.w_usuario.value.strip(),
                                   accion, cert, rem, campo, str(antes), str(despues), motivo])

    def resumen_estados(self):
        s = pd.Series([c["estado"] for c in self.certs.values()]).value_counts()
        return " · ".join(f"{cb.ICONOS[k]} {k}: {v}" for k, v in s.items())

    # ---------------------------------------------------------------- interfaz
    def mostrar(self):
        self.w_usuario = W.Text(description="Responsable:", placeholder="Tu nombre (obligatorio)",
                                style={"description_width": "initial"}, layout=_L("330px"))
        self.w_filtro = W.ToggleButtons(options=["Pendientes", "Todas", "Confirmadas"], value="Pendientes")
        self.w_sel = W.Dropdown(description="Certificación:", style={"description_width": "initial"},
                                layout=_L("640px"))
        self.w_cont = W.HTML()
        self.w_panel = W.VBox()
        self.w_out = W.Output(layout=W.Layout(border="1px solid #ddd", padding="4px"))
        self.w_filtro.observe(lambda ch: self._refrescar_lista(), "value")
        self.w_sel.observe(self._on_sel, "value")
        self._refrescar_lista()
        display(W.VBox([W.HBox([self.w_usuario, self.w_filtro]), W.HBox([self.w_sel]), self.w_cont,
                        self.w_panel, self.w_out]))

    def _etiqueta(self, c):
        return (f"{cb.ICONOS[c['estado']]} {c['periodo']:<8} {c['proveedor']}  ·  {len(c['filas'])} remitos  ·  "
                f"{cb.decimal(c['tn'])} TN" + (f"  ·  Nº {c['numero']}" if c["numero"] else ""))

    def _refrescar_lista(self, mantener=None):
        f = self.w_filtro.value
        ks = [k for k, c in self.certs.items()
              if f == "Todas" or (f == "Confirmadas") == (c["estado"] == "CONFIRMADA")]
        orden = {"1Q": 0, "2Q": 1, "Mensual": 2}
        ks.sort(key=lambda k: (orden[self.certs[k]["periodo"]], self.certs[k]["proveedor"].casefold()))
        opciones = [(self._etiqueta(self.certs[k]), k) for k in ks]
        destino = mantener if mantener in ks else (ks[0] if ks else None)
        self.w_sel.unobserve(self._on_sel, "value")
        self.w_sel.options = opciones
        self.w_sel.value = destino
        self.w_sel.observe(self._on_sel, "value")
        cuadra = all("✔" in r for r in self.cuad["Resultado"])
        self.w_cont.value = (f"<b>{len(self.certs)} certificaciones</b> — {self.resumen_estados()} &nbsp;|&nbsp; "
                             f"Cuadre con el maestro: {'<span style=color:green>✔ cuadra</span>' if cuadra else '<span style=color:red>✖ NO CUADRA</span>'}"
                             f" &nbsp;|&nbsp; Próximo Nº: {self.estado['prefijo']}-{self.estado['proximo']:08d}")
        self._render(destino)

    def _on_sel(self, ch):
        self._seguro(lambda: self._render(ch["new"]))

    # --------------------------------------------------------------- render
    def _render(self, key):
        self.key = key
        if key is None:
            self.w_panel.children = [W.HTML("<i>No hay certificaciones en esta vista.</i>")]
            return
        c = self.certs[key]
        meta = self.estado["meta"].get(key, {})
        est = c["estado"]
        color = {"CONFIRMADA": "#3A7D44", "LISTA": "#3A7D44", "MODIFICADA": "#B26B00",
                 "REVISAR": "#B26B00", "BLOQUEADA": "#C00000"}[est]
        self.multi = (len(c.get("miembros", [])) > 1 or c["filas"]["Proveedor"].nunique() > 1
                      or (c["filas"]["Proveedor"] != c["proveedor"]).any())
        self.gyg = c.get("formato") == "GYG"
        cab = W.HTML(f"<h3 style='margin:6px 0'>{c['proveedor']} — {cb.texto_periodo(c['desde'], c['hasta'])}</h3>"
                     + (f"<small>Incluye los puntos de entrega: <b>{' + '.join(c['miembros'])}</b></small><br>"
                        if self.multi else "")
                     + f"<span style='background:{color};color:white;padding:3px 8px;border-radius:4px'>"
                     f"{cb.ICONOS[est]} {est}</span> &nbsp; Grupo de precios: <b>{c['grupo']}</b><br>"
                     + "<small>Tipologías: " + " · ".join(
                         f"<b>{(cb.etiqueta_gyg if c.get('formato') == 'GYG' else cb.nombre_producto)(t)}</b> "
                         f"{cb.decimal(cb.tn_pagar(c, t))} TN" for t in cb.orden_productos(c["tn_prod"])) + "</small>")

        # Encabezado editable
        self.w_razon = W.Text(value=c["razon"], description="Razón social:", layout=_L("420px"),
                              style={"description_width": "110px"})
        self.w_cuit = W.Text(value=c["cuit"], description="CUIT:", placeholder="30-12345678-9", layout=_L("260px"),
                             style={"description_width": "50px"})
        self.w_num = W.Text(value=c["numero"], description="Nº:", placeholder="se asigna al confirmar",
                            layout=_L("260px"), style={"description_width": "30px"})
        self.w_fecha = W.DatePicker(value=c["fecha"], description="Fecha cert.:", layout=_L("250px"),
                                    style={"description_width": "80px"})
        self.w_tng = W.FloatText(value=round(c["tn_grupo"], 2), description="TN grupo mes:", layout=_L("230px"),
                                 style={"description_width": "100px"})
        encab = W.VBox([W.HBox([self.w_razon, self.w_cuit]), W.HBox([self.w_num, self.w_fecha, self.w_tng]),
                        W.HTML(f"<small>Según remitos el grupo suma {cb.decimal(c['tn_grupo_auto'])} TN en el mes "
                               "(define el tramo de precio).</small>")])

        # Tabla de remitos
        obs_id = {}
        for o in c["obs"]:
            if o["id"]:
                obs_id.setdefault(o["id"], []).append(o)
        tit = W.HBox([W.HTML(f"<b>{t}</b>", layout=_L(ANCHOS[k])) for k, t in
                      [("del", "🗑"), ("Batea", "Batea"), ("Fecha", "Fecha"), ("Rto. Planta", "Rto. Planta"),
                       ("Remito Proveedor", "Remito proveedor"), ("Producto", "Producto"), ("Kilos", "Kilos"),
                       ("Finos", "Aserrín lab %" if self.gyg else "Finos lab %"), ("Hum", "Humedad lab %"),
                       ("DescMan", "Desc. gran. %"), ("PenMan", "Desc. hum. %")]
                      + [("Desc", "resultado")] + ([("Punto", "Punto de entrega")] if self.multi else [])
                      + [("obs", "")]])
        self.filas_w = []
        self.w_tabla = W.VBox([tit])
        for _, f in c["filas"].iterrows():
            self._agregar_fila_w(dict(f), origen="cert", obs=obs_id.get(f["_id"], []))
        self.w_tn_pantalla = W.HTML()
        self._actualizar_tn()

        self.w_nuevo_rem = W.Text(placeholder="N° remito de planta, ej. 202511-12614", layout=_L("300px"))
        b_traer = W.Button(description="Traer remito del maestro", icon="search", layout=_L("230px"))
        b_vacia = W.Button(description="Agregar fila vacía", icon="plus", layout=_L("180px"))
        b_traer.on_click(lambda b: self._seguro(self._traer))
        b_vacia.on_click(lambda b: self._agregar_fila_w(
            {"Batea": "", "Fecha": c["hasta"], "Rto. Planta": "", "Remito Proveedor": "", "Producto": "CHIP",
             "Kilos": 0.0, "_id": None}, origen="manual"))
        agregar = W.HBox([b_vacia, self.w_nuevo_rem, b_traer])

        # Humedad y granulometría
        self.w_hum = {}
        self.w_cfgdesc = {}
        filas_h = []
        if True:   # opciones de descuento del proveedor: siempre visibles
            prov_cfg = self._cfg_unidad(c)
            g = lambda col: str(prov_cfg.get(col, "") if hasattr(prov_cfg, "get") else "").strip()
            ref = None
            self.w_cfgdesc = {
                "DESCUENTA HUMEDAD": W.Checkbox(value=g("DESCUENTA HUMEDAD").upper() == "SI", indent=False,
                                                description="Descuenta HUMEDAD", layout=_L("190px")),
                "TOLERANCIA HUMEDAD %": W.Text(value=g("TOLERANCIA HUMEDAD %"), placeholder="0",
                                               description="tolerancia (+ puntos %):", layout=_L("250px"),
                                               style={"description_width": "160px"}),
                "LIMITE HUMEDAD %": W.Text(value=g("LIMITE HUMEDAD %"), placeholder="el del laboratorio",
                                           description="límite fijo %:", layout=_L("250px"),
                                           style={"description_width": "100px"}),
                "DESCUENTA GRANULOMETRIA": W.Checkbox(value=g("DESCUENTA GRANULOMETRIA").upper() != "NO", indent=False,
                                                      description="Descuenta GRANULOMETRÍA", layout=_L("210px")),
                "TOLERANCIA FINOS %": W.Text(value=g("TOLERANCIA FINOS %"),
                                             placeholder=f"{cb.decimal(cb.FINOS_TOLERANCIA, 0)} (general)",
                                             description="tolerancia finos %:", layout=_L("250px"),
                                             style={"description_width": "130px"}),
            }
            self.w_sin_gran = W.Checkbox(value=bool(meta.get("sin_gran")), indent=False, layout=_L("420px"),
                                         description="Quitar el descuento de GRANULOMETRÍA en esta certificación")
            self.w_sin_hum = W.Checkbox(value=bool(meta.get("sin_hum")), indent=False, layout=_L("420px"),
                                        description="Quitar el descuento de HUMEDAD en esta certificación")
            filas_h.append(W.HTML("<b>Solo en esta certificación</b> <small>(pide motivo y queda registrado · para un "
                                  "camión puntual usá las columnas <i>Desc. gran. %</i> y <i>Desc. hum. %</i>"
                                  " de la tabla: escribí el % "
                                  "o 0 para quitarlo; borralo para volver al automático)</small>"))
            filas_h.append(W.HBox([self.w_sin_gran, self.w_sin_hum]))
            filas_h.append(W.HTML("<b>Qué se le descuenta a este proveedor</b> <small>(vale para todos sus períodos y se "
                                  "guarda en el config · cambiarlo pide motivo)</small>"))
            filas_h.append(W.HBox([self.w_cfgdesc["DESCUENTA HUMEDAD"], self.w_cfgdesc["TOLERANCIA HUMEDAD %"],
                                   self.w_cfgdesc["LIMITE HUMEDAD %"]]))
            filas_h.append(W.HBox([self.w_cfgdesc["DESCUENTA GRANULOMETRIA"], self.w_cfgdesc["TOLERANCIA FINOS %"],
                                   W.HTML(f"<small>&nbsp;por camión: finos hasta la tolerancia → 0 · hasta "
                                          f"{cb.decimal(cb.FINOS_TOTALIDAD, 0)} % → lo que excede · más de "
                                          f"{cb.decimal(cb.FINOS_TOTALIDAD, 0)} % → todos los finos · lo descontado se "
                                          "paga como ASERRÍN</small>")]))
            if self.gyg:
                filas_h.append(W.HTML(f"<small><b>Formato GYG</b>: humedad y granulometría <b>camión por camión</b>. "
                                      f"Aserrín: menos de {cb.decimal(c['tol_finos'], 0)} % (Alto) → 0 · hasta "
                                      f"{cb.decimal(c['muy_alto'], 0)} % (Muy Alto) → lo que excede · desde "
                                      f"{cb.decimal(c['muy_alto'], 0)} % → todo. Humedad: se descuenta lo que supera "
                                      "la referencia del día. Los valores de cada camión se corrigen en la tabla.</small>"))
            if c.get("desc_gran") and c.get("con_humedad"):
                fl = c["filas"]
                n_m, n_d = int(fl["Finos"].notna().sum()), int((fl["DescPct"] > 0).sum())
                filas_h.append(W.HTML(f"<small>Granulometría del período: {n_m} camiones con muestra · {n_d} con "
                                      f"descuento · <b>{cb.decimal(fl['DescKg'].sum() / 1000)} TN</b> de CHIP pasan a "
                                      "ASERRÍN. Los finos de cada camión se corrigen en la tabla de remitos.</small>"))
            if not c.get("con_humedad"):
                filas_h.append(W.HTML("<small style=color:#B26B00>⚠ No se cargó la planilla de calidad: los descuentos "
                                      "no se pueden calcular. Subila en el Paso 3 o en el Paso 5b.</small>"))
            elif any(hh["aplica_hum"] for hh in c["humedad"].values()):
                filas_h.append(W.HTML("<b>Humedad del período</b> <small>(promedio de las muestras del laboratorio · "
                                      "si la corregís a mano pide motivo)</small>"))
            for t, hh in ([(t, hh) for t, hh in c["humedad"].items() if hh["aplica_hum"]] if c.get("con_humedad") else []):
                base_l = None if hh["L"] is None else hh["L"] - hh["tol"]
                wh = W.FloatText(value=round(hh["H"] * 100, 2) if hh["H"] is not None else 0.0,
                                 description=f"{t} humedad %:", layout=_L("230px"), style={"description_width": "120px"})
                wl = W.FloatText(value=round(base_l * 100, 2) if base_l is not None else 0.0,
                                 description="límite %:", layout=_L("160px"), style={"description_width": "60px"})
                self.w_hum[t] = (wh, wl, dict(hh, L=base_l))
                partes = []
                if hh["aplica_hum"]:
                    partes.append(f"humedad {hh['muestras']} muestras, límite {cb._txt_limite(hh)} → "
                                  f"<b>−{cb.decimal(hh['desc_hum'] * 100)} %</b>")
                else:
                    partes.append("humedad: no se descuenta")
                partes.append(f"a pagar <b>{cb.decimal(hh['tn_pagar'])} TN</b> de {cb.decimal(hh['tn'])}")
                filas_h.append(W.HBox([wh, wl]))
                filas_h.append(W.HTML("<small>&nbsp;&nbsp;" + " · ".join(partes) + "</small>"))

        # Precios: todos los productos del cuadro, editables (cada proveedor tiene los suyos)
        self.w_precios, self.precios_base = {}, {}
        etiqueta = cb.etiqueta_gyg if self.gyg else cb.nombre_producto
        productos = list(dict.fromkeys(list(c["precios"]) + cb.PRODUCTOS_CUADRO))
        st = {"description_width": "110px"}
        filas_p = [W.HBox([W.HTML("<b>Producto</b>", layout=_L("110px")), W.HTML("<b>$ por TN</b>", layout=_L("150px")),
                           W.HTML("<b>TN a pagar</b>", layout=_L("110px")), W.HTML("<b>Subtotal</b>", layout=_L("150px")),
                           W.HTML("<b>De dónde sale</b>")])]
        for t in productos:
            entregado = t in c["precios"]
            p = c["precios"][t] if entregado else cb.buscar_precio(self.precios, c["grupo"], t, c["tn_grupo"])
            self.precios_base[t] = p
            w = W.FloatText(value=p if p is not None else 0.0, layout=_L("150px"))
            self.w_precios[t] = w
            tn = cb.tn_pagar(c, t) if entregado else 0.0
            origen = (c["origen_precios"][t] if entregado else ("config" if p is not None else "sin precio cargado"))
            falta = entregado and p is None
            filas_p.append(W.HBox([W.HTML(f"<b>{etiqueta(t)}</b>", layout=_L("110px")), w,
                                   W.HTML(cb.decimal(tn) if entregado else "—", layout=_L("110px")),
                                   W.HTML(cb.pesos(tn * p) if (entregado and p) else "", layout=_L("150px")),
                                   W.HTML(f"<small>{origen}</small>" + (" <b style=color:red>FALTA</b>" if falta else ""))]))
        self.w_guardar_cfg = W.RadioButtons(
            options=[("Guardar para este proveedor (se usan en las próximas certificaciones)", True),
                     ("Usar solo en esta certificación", False)], value=not meta.get("precios"), layout=_L("620px"))
        self.w_importe = W.HTML(f"<b style='font-size:15px;background:#FFFF00;padding:2px 8px'>Importe a facturar: "
                                f"{cb.pesos(c['importe'])}</b>")
        nota_g = (f"<small>Grupo de precios: <b>{c['grupo']}</b>"
                  + (f" · los comparten: {', '.join(self._grupo_comparten(c))}" if self._grupo_comparten(c) else "")
                  + " · si hay tramos por TN del mes, se cambia el del tramo actual "
                  f"({cb.decimal(c['tn_grupo'])} TN).</small>")
        descuentos_box = W.VBox(filas_h)
        precios_box = W.VBox(filas_p + [W.HTML(nota_g), self.w_guardar_cfg, self.w_importe])

        # Observaciones
        lineas = []
        for o in sorted(c["obs"], key=lambda o: (not o["bloq"], o["sev"] != cb.CRIT)):
            col = "#C00000" if o["bloq"] else ("#B26B00" if o["sev"] == cb.CRIT else "#666")
            tag = "BLOQUEANTE" if o["bloq"] else o["sev"]
            lineas.append(f"<li style='color:{col}'><b>{tag}</b> {('· ' + o['rem'] + ' · ') if o['rem'] else ''}{o['msg']}</li>")
        obs_html = W.HTML("<ul style='margin:2px'>" + "".join(dict.fromkeys(lineas)) + "</ul>" if lineas
                          else "<span style=color:green>✔ Sin observaciones</span>")
        revisables = cb.criticas_revisables(c)
        ack = meta.get("ack", {})
        self.w_ack = W.Checkbox(value=bool(revisables) and set(revisables) <= set(ack.get("msgs", [])), indent=False,
                                description="Revisé las observaciones críticas y la certificación es correcta")
        self.w_ack_com = W.Textarea(value=ack.get("comentario", ""), placeholder="Comentario obligatorio: qué revisaste "
                                    "(ej. 'el proveedor confirma que un remito cubre 7 viajes')", layout=_L("760px", height="54px"))
        ack_box = W.VBox([self.w_ack, self.w_ack_com]) if revisables else W.HTML("")

        self.w_motivo = W.Textarea(placeholder="Motivo de los cambios (obligatorio si cambiás remitos, descuentos o "
                                   "valores del laboratorio)", layout=_L("760px", height="54px"))
        b_guardar = W.Button(description="Guardar cambios", icon="save", button_style="info", layout=_L("190px"))
        b_conf = W.Button(description="Confirmar y generar PDF", icon="check", button_style="success", layout=_L("240px"))
        b_sig = W.Button(description="Siguiente pendiente", icon="arrow-right", layout=_L("190px"))
        b_desc = W.Button(description="Descartar cambios", icon="undo", layout=_L("170px"))
        b_guardar.on_click(lambda b: self._seguro(self._guardar))
        b_conf.on_click(lambda b: self._seguro(self._confirmar))
        b_sig.on_click(lambda b: self._seguro(self._siguiente))
        b_desc.on_click(lambda b: self._seguro(lambda: self._render(self.key)))

        sec = lambda t: W.HTML(f"<h4 style='margin:12px 0 4px;border-bottom:2px solid #3A7D44'>{t}</h4>")
        self.w_panel.children = [cab, sec("1 · Encabezado"), encab,
                                 sec(f"2 · Remitos ({len(c['filas'])})"), self.w_tabla, self.w_tn_pantalla, agregar,
                                 sec("3 · Descuentos de humedad y granulometría"), descuentos_box,
                                 sec("4 · Precios ($ por TN)"), precios_box,
                                 sec("5 · Observaciones"), obs_html, ack_box,
                                 sec("6 · Guardar y confirmar"), self.w_motivo,
                                 W.HBox([b_guardar, b_conf, b_sig, b_desc])]

    def _seguro(self, fn):
        """Ejecuta la acción de un botón y, si falla, muestra el error en pantalla (Colab lo oculta)."""
        import traceback
        try:
            fn()
        except Exception as e:
            with self.w_out:
                self.w_out.clear_output()
                display(W.HTML(f"<b style=color:red>Error: {type(e).__name__}: {e}</b><br><small>Copiá este detalle si "
                               "necesitás ayuda:</small><pre style='font-size:11px'>"
                               + traceback.format_exc()[-2500:].replace("<", "&lt;") + "</pre>"))

    def _cambios_descuentos(self, prov):
        """(columna, valor nuevo, valor anterior) de la configuración de descuentos que cambió en pantalla."""
        out = []
        if not self.w_cfgdesc:
            return out
        actual = self._cfg_unidad(self.certs[self.key])
        for col, w in self.w_cfgdesc.items():
            viejo = str(actual.get(col, "") if hasattr(actual, "get") else "").strip()
            if isinstance(w, W.Checkbox):
                nuevo = "SI" if w.value else "NO"
                defecto = "NO" if col == "DESCUENTA HUMEDAD" else "SI"
                if (viejo.upper() or defecto) != nuevo:
                    out.append((col, nuevo, viejo))
            else:
                nuevo = w.value.strip().replace("%", "")
                if nuevo != viejo.replace("%", ""):
                    out.append((col, nuevo, viejo))
        return out

    def _grupo_comparten(self, c):
        """Otros proveedores (fuera de esta certificación) que usan el mismo grupo de precios."""
        if "GRUPO" not in self.conf.columns:
            return []
        g = self.conf["GRUPO"].astype(str).str.casefold() == str(c["grupo"]).casefold()
        return sorted(set(self.conf.index[g]) - set(c.get("miembros", [c["proveedor"]])), key=str.casefold)

    def _cfg_unidad(self, c):
        """Fila del config que manda en la certificación (el primer punto de entrega con CUIT)."""
        ms = [p for p in c.get("miembros", [c["proveedor"]]) if p in self.conf.index]
        con_cuit = [p for p in ms if str(self.conf.at[p, "CUIT"]).strip()]
        return self.conf.loc[(con_cuit or ms)[0]] if ms else {}

    def _miembros_cfg(self, c):
        return list(c.get("miembros") or [c["proveedor"]])

    @staticmethod
    def _txt_num(v, n=2):
        if v is None or pd.isna(v):
            return ""
        t = f"{float(v):.{n}f}".rstrip("0").rstrip(".")
        return t.replace(".", ",") if t else "0"

    @staticmethod
    def _tip(w, texto):
        for attr in ("tooltip", "description_tooltip"):
            if w.has_trait(attr):
                setattr(w, attr, texto)

    @staticmethod
    def _num(txt):
        """Texto de un campo → número (acepta coma y %); '' → None; inválido → 'ERROR'."""
        t = str(txt).strip().replace("%", "").replace(" ", "")
        if t == "":
            return None
        if "," in t:
            t = t.replace(".", "").replace(",", ".")
        try:
            return float(t)
        except ValueError:
            return "ERROR"

    def _manuales_cambiados(self, campo, base_val, base_man, escala=1.0):
        """Descuentos editados por camión: [(fila, nuevo valor | None = volver a automático | 'ERROR', anterior)]."""
        out = []
        for x in self.filas_w:
            w = x["w"]
            if x["origen"] != "cert" or w["del"].value or campo not in w or w[campo].disabled:
                continue
            nuevo = self._num(w[campo].value)
            antes = x["base"].get(base_val)
            antes = 0.0 if antes is None or pd.isna(antes) else float(antes) * escala
            era_manual = bool(x["base"].get(base_man))
            if nuevo == "ERROR":
                out.append((x, "ERROR", antes))
            elif nuevo is None:
                if era_manual:
                    out.append((x, None, antes))
            elif round(nuevo, 2) != round(antes, 2):
                out.append((x, nuevo, antes))
        return out

    @staticmethod
    def _txt_finos(v):
        return "" if v is None or pd.isna(v) else f"{v * 100:.1f}".replace(".", ",")

    def _finos_cambiados(self, campo="Finos"):
        """[(fila, valor nuevo como fracción o None, valor anterior)] de los finos (o humedad) editados en la tabla."""
        out = []
        for x in self.filas_w:
            if x["origen"] != "cert" or x["w"]["del"].value or campo not in x["w"]:
                continue
            txt = x["w"][campo].value.strip().replace("%", "").replace(",", ".")
            antes = x["base"].get(campo)
            antes = None if antes is None or pd.isna(antes) else float(antes)
            if txt == "":
                nuevo = None
            else:
                try:
                    nuevo = float(txt) / 100
                except ValueError:
                    out.append((x, "ERROR", antes))
                    continue
            if (nuevo is None) != (antes is None) or (nuevo is not None and round(nuevo * 100, 1) != round(antes * 100, 1)):
                out.append((x, nuevo, antes))
        return out

    def _cambios_valores(self, t):
        wh, wl, hh = self.w_hum[t]
        man = {}
        for w, k, base in ((wh, "H", hh["H"]), (wl, "L", hh["L"])):
            b = None if base is None else round(base * 100, 2)
            if w.value and (b is None or abs(w.value - b) > 0.001):
                man[k] = w.value / 100
        return man

    def _agregar_fila_w(self, f, origen, obs=()):
        nuevo = origen != "cert"
        w = {"del": W.Checkbox(value=False, indent=False, layout=_L(ANCHOS["del"])),
             "Batea": W.Text(value=str(f["Batea"]), layout=_L(ANCHOS["Batea"])),
             "Fecha": W.DatePicker(value=f["Fecha"], layout=_L(ANCHOS["Fecha"])),
             "Rto. Planta": W.Text(value=str(f["Rto. Planta"]), disabled=origen != "manual",
                                   layout=_L(ANCHOS["Rto. Planta"])),
             "Remito Proveedor": W.Text(value=str(f["Remito Proveedor"]), layout=_L(ANCHOS["Remito Proveedor"])),
             "Producto": W.Dropdown(options=list(dict.fromkeys(cb.TIPOLOGIAS + [f["Producto"]])), value=f["Producto"],
                                    layout=_L(ANCHOS["Producto"])),
             "Kilos": W.FloatText(value=float(f["Kilos"]), layout=_L(ANCHOS["Kilos"]))}
        # Granulometría del camión (editable solo en remitos ya guardados; vacío = sin muestra)
        hora, org = f.get("FinosHora") or "", f.get("FinosOrigen") or ""
        w["Finos"] = W.Text(value=self._txt_finos(f.get("Finos")), placeholder="—", disabled=origen != "cert",
                            layout=_L(ANCHOS["Finos"]))
        for attr in ("tooltip", "description_tooltip"):
            if w["Finos"].has_trait(attr):
                setattr(w["Finos"], attr, (f"muestra {hora} · {org}" if hora else org) or "sin muestra asignada")
        dp, dk = float(f.get("DescPct") or 0), float(f.get("DescKg") or 0)
        extras = [w["Finos"]]
        gyg = getattr(self, "gyg", False)
        if True:                       # humedad del laboratorio de cada camión (todos los proveedores)
            w["Hum"] = W.Text(value=self._txt_finos(f.get("Hum")), placeholder="—", disabled=origen != "cert",
                              layout=_L(ANCHOS["Hum"]))
            extras.append(w["Hum"])
        # Descuento aplicado (editable): se puede poner a mano o en 0 para quitarlo; vacío = automático
        cact = self.certs.get(self.key, {}) if getattr(self, "key", None) else {}
        es_chip = f.get("Producto") == "CHIP" and origen == "cert" and not cact.get("sin_gran")
        w["DescMan"] = W.Text(value=self._txt_num(dp) if es_chip else "", placeholder="auto",
                              disabled=not es_chip, layout=_L(ANCHOS["DescMan"]))
        self._tip(w["DescMan"], "Puesto a mano (borralo para volver al automático)" if f.get("DescManual")
                  else "Automático según el laboratorio · escribí otro % o 0 para quitarlo")
        extras.append(w["DescMan"])
        if True:                       # descuento de humedad por camión (todos los proveedores)
            es_ch = f.get("Producto") in ("CHIP", "ASERRIN") and origen == "cert" and not cact.get("sin_hum")
            pen = f.get("Pen") or 0.0
            w["PenMan"] = W.Text(value=self._txt_num((pen or 0) * 100) if es_ch else "", placeholder="auto",
                                 disabled=not es_ch, layout=_L(ANCHOS["PenMan"]))
            self._tip(w["PenMan"], "Puesto a mano (borralo para volver al automático)" if f.get("PenManual")
                      else "Automático: humedad medida − referencia · escribí otro % o 0 para quitarlo")
            extras.append(w["PenMan"])
            pen = f.get("Pen") or 0
            partes = (["✋"] if f.get("DescManual") or f.get("PenManual") else []) + \
                     ([f"{cb.decimal(dk, 0)} kg aserrín"] if dk > 0 else []) + \
                     ([f"−{cb.decimal(pen * 100, 2)} % hum."] if pen and pen > 0 else [])
            extras.append(W.HTML(f"<span style='background:#FFF2A8;padding:1px 4px'>{' · '.join(partes)}</span>"
                                 if partes else "", layout=_L(ANCHOS["Desc"])))
        else:
            extras.append(W.HTML(f"<span style='background:#FFF2A8;padding:1px 4px'>{'✋ ' if f.get('DescManual') else ''}"
                                 f"{cb.decimal(dk, 1)} kg → aserrín</span>" if dp > 0 else
                                 ("✋ sin descuento" if f.get("DescManual") else ""), layout=_L(ANCHOS["Desc"])))
        if getattr(self, "multi", False):
            extras.append(W.HTML(f"<small>{f.get('Proveedor', '') or ''}</small>", layout=_L(ANCHOS["Punto"])))
        if obs:
            peor = "🔴" if any(o["sev"] == cb.CRIT for o in obs) else "🟠"
            txt = " | ".join(dict.fromkeys(o["msg"] for o in obs)).replace("'", "’")
            icono = W.HTML(f"<span title='{txt}' style='cursor:help'>{peor}</span>", layout=_L(ANCHOS["obs"]))
        else:
            icono = W.HTML("🆕" if nuevo else "", layout=_L(ANCHOS["obs"]))
        for k in ("Kilos", "del"):
            w[k].observe(lambda ch: self._actualizar_tn(), "value")
        caja = W.HBox([w[k] for k in ["del", "Batea", "Fecha", "Rto. Planta", "Remito Proveedor", "Producto", "Kilos"]]
                      + extras + [icono])
        self.filas_w.append({"w": w, "id": f.get("_id"), "origen": origen, "base": f})
        self.w_tabla.children = list(self.w_tabla.children) + [caja]
        if hasattr(self, "w_tn_pantalla"):
            self._actualizar_tn()

    def _actualizar_tn(self):
        tot = sum((x["w"]["Kilos"].value or 0) for x in self.filas_w if not x["w"]["del"].value)
        n = sum(1 for x in self.filas_w if not x["w"]["del"].value)
        self.w_tn_pantalla.value = f"<b>En pantalla: {n} remitos · {cb.decimal(tot / 1000)} TN</b>"

    def _msg(self, html, limpiar=True):
        with self.w_out:
            if limpiar:
                self.w_out.clear_output()
            display(W.HTML(html))

    def _traer(self):
        rem = self.w_nuevo_rem.value.strip()
        c = self.certs[self.key]
        if not rem:
            return self._msg("<b style=color:red>Escribí el N° de remito de planta.</b>")
        if any(x["w"]["Rto. Planta"].value == rem for x in self.filas_w):
            return self._msg(f"<b style=color:red>El remito {rem} ya está en la tabla.</b>")
        m = self.d[self.d[cb.C_REM_PLANTA] == rem]
        if len(m):
            for _, f in m.iterrows():
                base = {k: f[cb.CAMPOS[k]] for k in cb.FILA_COLS}
                base["Fecha"], base["_id"] = f[cb.C_DIA].date(), f["_id"]
                base["Proveedor"] = f[cb.C_PROV]
                self._agregar_fila_w(base, origen="maestro")
                fuera = not (c["desde"] <= base["Fecha"] <= c["hasta"])
                self._msg(f"Traído <b>{rem}</b>: proveedor en el maestro <b>{f[cb.C_PROV]}</b>, fecha "
                          f"<b>{base['Fecha']:%d/%m/%Y}</b>."
                          + ("<br><b style=color:#B26B00>La fecha está fuera de este período: corregila o el remito "
                             "no va a aparecer acá.</b>" if fuera else "")
                          + "<br>Completá el motivo y tocá <b>Guardar cambios</b>.")
            return
        r = self.raw[self.raw[cb.C_REM_PLANTA].apply(cb._txt) == rem]
        if len(r):
            f = r.iloc[0]
            base = {"Batea": cb._txt(f[cb.C_BATEA]), "Fecha": f[cb.C_DIA].date() if pd.notna(f[cb.C_DIA]) else c["hasta"],
                    "Rto. Planta": rem, "Remito Proveedor": cb._txt(f[cb.C_REM_PROV]),
                    "Producto": cb._canon_tipo(f[cb.C_TIPO]), "Kilos": float(pd.to_numeric(f[cb.C_NETO], errors="coerce") or 0),
                    "_id": None, "Proveedor": cb._txt(f[cb.C_PROV])}
            self._agregar_fila_w(base, origen="raw")
            return self._msg(f"Traído <b>{rem}</b> desde el maestro (estaba como movimiento "
                             f"'<b>{cb._txt(f[cb.C_MOV])}</b>' de <b>{base['Proveedor']}</b>). Revisalo, completá el motivo y guardá.")
        self._msg(f"<b style=color:red>No encontré el remito {rem} en el maestro.</b> Si es un remito real que no se "
                  "cargó, usá <i>Agregar fila vacía</i>.")

    # --------------------------------------------------------------- guardar
    def _guardar(self, silencioso=False):
        c = self.certs[self.key]
        prov = c["proveedor"]
        miembros = self._miembros_cfg(c)
        destino = miembros[0]          # punto de entrega al que se asignan los remitos agregados/traídos
        usuario = self.w_usuario.value.strip()
        motivo = self.w_motivo.value.strip()
        errores, cambios = [], []
        if not usuario:
            errores.append("Completá <b>Responsable</b> (arriba de todo).")
        ids_actuales = set(c["filas"]["_id"])
        nuevos_rem = []
        for x in self.filas_w:
            w = x["w"]
            if w["del"].value:
                continue
            rem = w["Rto. Planta"].value.strip()
            if w["Fecha"].value is None:
                errores.append(f"Remito {rem or '(nuevo)'}: falta la fecha.")
            if not w["Kilos"].value or w["Kilos"].value <= 0:
                errores.append(f"Remito {rem or '(nuevo)'}: los kilos deben ser mayores a 0.")
            elif w["Kilos"].value > cb.PESO_MAX * 1.2:
                errores.append(f"Remito {rem}: {cb.miles(w['Kilos'].value)} kg es imposible para un camión.")
            if w["Producto"].value not in cb.TIPOLOGIAS:
                errores.append(f"Remito {rem}: elegí un producto válido.")
            if x["origen"] == "manual":
                if not rem:
                    errores.append("Una fila nueva no tiene N° de remito de planta.")
                elif rem in nuevos_rem or (self.d[cb.C_REM_PLANTA] == rem).any():
                    errores.append(f"El remito {rem} ya existe: usá <i>Traer remito del maestro</i>.")
                nuevos_rem.append(rem)

        vals = lambda w: {"Batea": w["Batea"].value.strip(), "Fecha": w["Fecha"].value,
                          "Rto. Planta": w["Rto. Planta"].value.strip(),
                          "Remito Proveedor": w["Remito Proveedor"].value.strip(),
                          "Producto": w["Producto"].value, "Kilos": float(w["Kilos"].value or 0)}
        for x in self.filas_w:
            w, base = x["w"], x["base"]
            v = vals(w)
            if x["origen"] == "cert":
                if w["del"].value:
                    cambios.append(("eliminar", x["id"], v["Rto. Planta"], "", "", ""))
                    continue
                for k in ("Batea", "Fecha", "Remito Proveedor", "Producto", "Kilos"):
                    antes = base[k]
                    if (float(antes) != v[k]) if k == "Kilos" else (str(antes) != str(v[k])):
                        cambios.append(("editar", x["id"], v["Rto. Planta"], k, antes, v[k]))
            elif w["del"].value:
                continue
            elif x["origen"] == "maestro":
                if base["Proveedor"] not in miembros:
                    cambios.append(("editar", x["id"], v["Rto. Planta"], "Proveedor", base["Proveedor"], destino))
                for k in ("Batea", "Fecha", "Remito Proveedor", "Producto", "Kilos"):
                    if (float(base[k]) != v[k]) if k == "Kilos" else (str(base[k]) != str(v[k])):
                        cambios.append(("editar", x["id"], v["Rto. Planta"], k, base[k], v[k]))
                if not any(cm[1] == x["id"] for cm in cambios):
                    cambios.append(("editar", x["id"], v["Rto. Planta"], "Proveedor", base["Proveedor"],
                                    base["Proveedor"] if base["Proveedor"] in miembros else destino))
            else:   # raw o manual → remito agregado
                cambios.append(("agregar", None, v["Rto. Planta"], "", "", v))
        if cambios and not motivo:
            errores.append("Escribí el <b>motivo</b> de los cambios en remitos.")
        if self.w_ack.value and cb.criticas_revisables(c) and not self.w_ack_com.value.strip():
            errores.append("Para marcar las observaciones como revisadas escribí un <b>comentario</b>.")
        cuit = self.w_cuit.value.strip()
        if cuit and not cb.cuit_valido(cuit):
            errores.append(f"El CUIT <b>{cuit}</b> no es válido (dígito verificador). Revisalo.")
        for t, w in self.w_precios.items():
            if w.value is not None and w.value < 0:
                errores.append(f"Precio de {t} negativo.")
        cambios_cfg = self._cambios_descuentos(prov)
        if cambios_cfg and not motivo:
            errores.append("Escribí el <b>motivo</b> del cambio en los descuentos del proveedor.")
        for campo, nuevo, _ in cambios_cfg:
            if campo == "LIMITE HUMEDAD %" and nuevo and not (isinstance(cb.puntos(nuevo), float)
                                                             and 0.2 <= cb.puntos(nuevo) < 0.9):
                errores.append(f"Límite fijo de humedad '{nuevo}': tiene que ser un % entre 20 y 90.")
            elif "TOLERANCIA" in campo and nuevo and not isinstance(cb.puntos(nuevo), float):
                errores.append(f"{campo}: '{nuevo}' no es un número.")
            elif "TOLERANCIA" in campo and nuevo and not (0 <= cb.puntos(nuevo) < (0.2 if "HUMEDAD" in campo else 1)):
                errores.append(f"{campo} fuera de rango.")
        finos_mod = self._finos_cambiados()
        desc_mod = self._manuales_cambiados("DescMan", "DescPct", "DescManual")
        pen_mod = self._manuales_cambiados("PenMan", "Pen", "PenManual", escala=100)
        for x, nuevo, _ in desc_mod + pen_mod:
            rem = x["w"]["Rto. Planta"].value
            if nuevo == "ERROR":
                errores.append(f"Remito {rem}: el descuento escrito no es un número.")
            elif nuevo is not None and not (0 <= nuevo <= 100):
                errores.append(f"Remito {rem}: el descuento tiene que estar entre 0 y 100 %.")
        sw_mod = [(k, w.value) for k, w in (("sin_gran", self.w_sin_gran), ("sin_hum", self.w_sin_hum))
                  if bool(self.estado["meta"].get(self.key, {}).get(k)) != w.value]
        if (desc_mod or pen_mod or sw_mod) and not motivo:
            errores.append("Escribí el <b>motivo</b> del cambio en los descuentos.")
        for x, nuevo, _ in finos_mod:
            rem = x["w"]["Rto. Planta"].value
            if nuevo == "ERROR":
                errores.append(f"Remito {rem}: los finos '{x['w']['Finos'].value}' no son un número.")
            elif nuevo is not None and not (0 <= nuevo < 1):
                errores.append(f"Remito {rem}: finos fuera de rango (0 a 100 %).")
        if finos_mod and not motivo:
            errores.append("Escribí el <b>motivo</b> de la corrección de finos (granulometría).")
        hum_mod = self._finos_cambiados("Hum")
        for x, nuevo, _ in hum_mod:
            rem = x["w"]["Rto. Planta"].value
            if nuevo == "ERROR":
                errores.append(f"Remito {rem}: la humedad '{x['w']['Hum'].value}' no es un número.")
            elif nuevo is not None and not (0 < nuevo < 1):
                errores.append(f"Remito {rem}: humedad fuera de rango (0 a 100 %).")
        if hum_mod and not motivo:
            errores.append("Escribí el <b>motivo</b> de la corrección de humedad del camión.")
        for t, (wh, wl, hh) in self.w_hum.items():
            if self._cambios_valores(t) and not motivo:
                errores.append(f"Escribí el <b>motivo</b> de la corrección de humedad de {t}.")
            for w, n in ((wh, "humedad"), (wl, "límite")):
                if w.value and not (0 < w.value < 100):
                    errores.append(f"{n.capitalize()} de {t} fuera de rango (0 a 100 %).")
        if errores:
            self._msg("<b style=color:red>No se guardó:</b><ul>" + "".join(f"<li>{e}</li>" for e in errores) + "</ul>")
            return False

        corr, meta = self.estado["corr"], self.estado["meta"].setdefault(self.key, {})
        for accion, _id, rem, campo, antes, despues in cambios:
            if accion == "eliminar":
                corr["eliminados"][_id] = motivo
                self._log("Quitar remito", self.key, rem, "", "", "", motivo)
            elif accion == "editar":
                ed = corr["editados"].setdefault(_id, {})
                ed[campo] = despues.isoformat() if isinstance(despues, date) else despues
                self._log("Corregir remito", self.key, rem, campo,
                          antes.strftime("%d/%m/%Y") if isinstance(antes, date) else antes,
                          despues.strftime("%d/%m/%Y") if isinstance(despues, date) else despues, motivo)
            else:
                v = dict(despues, Proveedor=destino, Fecha=despues["Fecha"].isoformat())
                corr["agregados"][f"{rem}#A{len(corr['agregados']) + 1}"] = v
                self._log("Agregar remito", self.key, rem, "", "", f"{v['Kilos']:.0f} kg {v['Producto']} {v['Fecha']}", motivo)

        # Finos por camión corregidos a mano (quedan registrados)
        for x, nuevo, antes in finos_mod:
            corr.setdefault("finos", {})[x["id"]] = nuevo
            self._log("Finos (granulometría) corregidos a mano", self.key, x["w"]["Rto. Planta"].value, "Finos %",
                      "sin muestra" if antes is None else f"{antes * 100:.1f} %",
                      "sin muestra" if nuevo is None else f"{nuevo * 100:.1f} %", motivo)

        for x, nuevo, antes in hum_mod:
            corr.setdefault("humedad_camion", {})[x["id"]] = nuevo
            self._log("Humedad del camión corregida a mano", self.key, x["w"]["Rto. Planta"].value, "Humedad %",
                      "sin muestra" if antes is None else f"{antes * 100:.1f} %",
                      "sin muestra" if nuevo is None else f"{nuevo * 100:.1f} %", motivo)

        # Encabezado → config de todos los puntos de entrega de la certificación (vale para todos sus períodos)
        actual = self._cfg_unidad(c)
        g0 = lambda col: str(actual.get(col, "") if hasattr(actual, "get") else "").strip()
        for m in miembros:
            if m not in self.conf.index:
                fila = {col: "" for col in self.conf.columns}
                fila.update({"RAZÓN SOCIAL": g0("RAZÓN SOCIAL") or prov, "CUIT": g0("CUIT"), "GRUPO": g0("GRUPO") or prov,
                             "CERTIFICAR": "SI", "DESCUENTA HUMEDAD": g0("DESCUENTA HUMEDAD") or "NO",
                             "DESCUENTA GRANULOMETRIA": g0("DESCUENTA GRANULOMETRIA") or "SI",
                             "CERTIFICAR CON": prov if prov != m else ""})
                self.conf.loc[m] = fila
                self._log("Alta de proveedor en config", self.key, "", "", "", m)
        for campo, w in (("RAZÓN SOCIAL", self.w_razon), ("CUIT", self.w_cuit)):
            nuevo, viejo = w.value.strip(), g0(campo)
            if nuevo != viejo or any(str(self.conf.at[m, campo]).strip() != nuevo for m in miembros):
                for m in miembros:
                    self.conf.at[m, campo] = nuevo
                if nuevo != viejo:
                    self._log("Config proveedor", self.key, "", campo, viejo, nuevo)
        # Precios
        # Descuentos quitados o puestos a mano (por camión o en toda la certificación)
        for x, nuevo, antes in desc_mod:
            if nuevo is None:
                corr.setdefault("desc_gran", {}).pop(x["id"], None)
            else:
                corr.setdefault("desc_gran", {})[x["id"]] = nuevo
            self._log("Descuento de granulometría a mano", self.key, x["w"]["Rto. Planta"].value, "Desc. %",
                      f"{antes:g} %", "automático" if nuevo is None else f"{nuevo:g} %", motivo)
        for x, nuevo, antes in pen_mod:
            if nuevo is None:
                corr.setdefault("pen_hum", {}).pop(x["id"], None)
            else:
                corr.setdefault("pen_hum", {})[x["id"]] = nuevo / 100
            self._log("Descuento de humedad a mano", self.key, x["w"]["Rto. Planta"].value, "Penalización %",
                      f"{antes:g} %", "automático" if nuevo is None else f"{nuevo:g} %", motivo)
        for k, v in sw_mod:
            if v:
                meta[k] = True
            else:
                meta.pop(k, None)
            self._log("Descuento " + ("de granulometría" if k == "sin_gran" else "de humedad")
                      + (" QUITADO en la certificación" if v else " vuelve a aplicarse"), self.key, "", "", "", "", motivo)
        for t, w in self.w_precios.items():
            viejo = self.precios_base.get(t)
            if w.value and (viejo is None or abs(w.value - viejo) > 0.001):
                if self.w_guardar_cfg.value:
                    self.precios = cb.fijar_precio(self.precios, c["grupo"], t, c["tn_grupo"], w.value)
                    meta.get("precios", {}).pop(t, None)
                    self._log("Precio en config", self.key, "", f"{c['grupo']} / {t}", viejo or "", w.value)
                else:
                    meta.setdefault("precios", {})[t] = w.value
                    self._log("Precio solo para esta certificación", self.key, "", t, viejo or "", w.value)
        # Qué se descuenta y tolerancias del proveedor → config
        for campo, nuevo, viejo in cambios_cfg:
            if campo not in self.conf.columns:
                self.conf[campo] = ""
            for m in miembros:
                self.conf.at[m, campo] = nuevo
            self._log("Descuentos del proveedor (config)", self.key, "", campo, viejo, nuevo, motivo)
        # Humedad / finos corregidos a mano (quedan registrados)
        for t in self.w_hum:
            man = self._cambios_valores(t)
            previo = meta.get("humedad", {}).get(t) or {}
            if man and dict(previo, **man) != previo:
                meta.setdefault("humedad", {})[t] = dict(previo, **man)
                for k, v in man.items():
                    self._log("Valor de laboratorio corregido a mano", self.key, "", f"{t} · "
                              + {"H": "humedad", "L": "límite humedad"}[k], "", f"{v * 100:g} %", motivo)
        # TN grupo, número, fecha
        if abs(self.w_tng.value - c["tn_grupo_auto"]) > 0.005:
            if abs(self.w_tng.value - c["tn_grupo"]) > 0.005:
                self._log("TN grupo manual", self.key, "", "TN grupo", c["tn_grupo"], self.w_tng.value)
            meta["tn_grupo"] = self.w_tng.value
        else:
            meta.pop("tn_grupo", None)
        if self.w_num.value.strip() != c["numero"]:
            self._log("Número de certificación", self.key, "", "Nº", c["numero"], self.w_num.value.strip())
            meta["numero"] = self.w_num.value.strip()
        if self.w_fecha.value and self.w_fecha.value != c["fecha"]:
            meta["fecha"] = self.w_fecha.value.isoformat()
        # Revisión de observaciones
        if self.w_ack.value and cb.criticas_revisables(c):
            meta["ack"] = {"msgs": cb.criticas_revisables(c), "comentario": self.w_ack_com.value.strip(),
                           "usuario": usuario, "fecha": datetime.now().isoformat(timespec="seconds")}
            self._log("Observaciones revisadas", self.key, "", "", "", self.w_ack_com.value.strip())

        self.reconstruir()
        self.guardar_estado()
        if not silencioso:
            fuera = [x["w"]["Rto. Planta"].value for x in self.filas_w if not x["w"]["del"].value
                     and x["w"]["Fecha"].value and not (c["desde"] <= x["w"]["Fecha"].value <= c["hasta"])]
            self._msg(f"<b style=color:green>✔ Guardado.</b> {len(cambios)} cambio(s) en remitos."
                      + (f"<br>Los remitos {', '.join(fuera)} quedaron con fecha fuera de este período y "
                         "pasaron a la certificación que corresponde." if fuera else ""))
        self._refrescar_lista(mantener=self.key)
        return True

    # ------------------------------------------------------------- confirmar
    def _confirmar(self):
        key = self.key
        if not self._guardar(silencioso=True):
            return
        ok, msg, rutas = self.confirmar_cert(key, self.w_usuario.value.strip())
        self._msg(msg)
        if ok and self.descargar_pdf:
            self._descargar(rutas)
        self._refrescar_lista(mantener=key)

    def _descargar(self, rutas):
        """Descarga archivos desde el área de mensajes (en Colab las descargas de un botón tienen que salir de ahí)."""
        if not _colab_files:
            return
        with self.w_out:
            for r in rutas:
                try:
                    _colab_files.download(str(r))
                except Exception as e:
                    display(W.HTML(f"<small style=color:#B26B00>No se pudo descargar {Path(r).name} ({e}). "
                                   "Queda en el zip del Paso 7.</small>"))

    def confirmar_cert(self, key, usuario=""):
        """Confirma una certificación: número, PDF y Excel. Devuelve (ok, mensaje_html, [rutas])."""
        if key not in self.certs:
            return False, "<b style=color:red>Esta certificación no existe o quedó sin remitos.</b>", []
        c = self.certs[key]
        meta = self.estado["meta"].setdefault(key, {})
        if c["estado"] == "CONFIRMADA":
            return False, (f"Ya estaba confirmada (Nº {c['numero']}) y no cambió.<br>PDF: {Path(meta.get('pdf', '')).name}"
                           f"<br>Excel: {Path(meta.get('excel', '')).name}"), []
        bloq = [o["msg"] for o in c["obs"] if o["bloq"]]
        if bloq:
            return False, ("<b style=color:red>⛔ No se puede confirmar. Corregí primero:</b><ul>"
                           + "".join(f"<li>{m}</li>" for m in dict.fromkeys(bloq)) + "</ul>"), []
        if c["estado"] == "REVISAR":
            return False, ("<b style=color:#B26B00>⚠️ Hay observaciones críticas sin revisar.</b> Revisalas, marcá "
                           "<i>Revisé las observaciones críticas</i>, escribí el comentario y volvé a confirmar."), []
        if not c["numero"]:
            meta["numero"] = f"{self.estado['prefijo']}-{self.estado['proximo']:08d}"
            self.estado["proximo"] += 1
        meta.setdefault("fecha", c["fecha"].isoformat())
        self.recalcular(key)
        c = self.certs[key]
        nombre = "".join(ch if ch.isalnum() or ch in "-_" else "_" for ch in c["proveedor"]).strip("_")
        base = f"{c['numero']}_{nombre}_{c['periodo']}_{cb.MESES[c['desde'].month]}_{c['desde'].year}"
        carpeta = self.salida / "PDF" / c["periodo"]
        carpeta.mkdir(parents=True, exist_ok=True)
        ruta = carpeta / f"{base}.pdf"
        cb.pdf_certificacion(ruta, c, self.logo)
        regenerada = bool(meta.get("huella"))
        meta.update({"huella": cb.huella(c), "pdf": str(ruta), "confirmado_por": usuario,
                     "confirmado_el": datetime.now().isoformat(timespec="seconds")})
        self.estado["log"].append([datetime.now().strftime("%d/%m/%Y %H:%M:%S"), usuario,
                                   "PDF regenerado" if regenerada else "Confirmada", key, "", "Nº", "", c["numero"],
                                   f"{cb.decimal(c['tn'])} TN · {cb.pesos(c['importe'])}"])
        self.recalcular(key)
        c = self.certs[key]
        carpeta_x = self.salida / "EXCEL" / c["periodo"]
        carpeta_x.mkdir(parents=True, exist_ok=True)
        ruta_x = carpeta_x / f"{base}.xlsx"
        cb.excel_certificacion(ruta_x, c, self.logo)
        meta["excel"] = str(ruta_x)
        self.guardar_estado()
        tn_p = sum(cb.tn_pagar(c, t) for t in c["precios"])
        desc = f" · {cb.decimal(tn_p)} TN a pagar (descuentos −{cb.decimal(c['tn'] - tn_p)})" if abs(c["tn"] - tn_p) > 0.004 else ""
        msg = (f"<b style=color:green>✅ Confirmada Nº {c['numero']}</b> · {cb.decimal(c['tn'])} TN{desc} · "
               f"<b>{cb.pesos(c['importe'])}</b><br>PDF: {ruta.name}<br>Excel: {ruta_x.name}")
        return True, msg, [ruta, ruta_x]

    def marcar_revisadas(self, usuario, comentario, keys):
        """Sin editor: marca como revisadas las observaciones críticas (no bloqueantes) de esas certificaciones."""
        for k in keys:
            c = self.certs[k]
            crit = cb.criticas_revisables(c)
            if crit:
                self.estado["meta"].setdefault(k, {})["ack"] = {"msgs": crit, "comentario": comentario, "usuario": usuario,
                                                               "fecha": datetime.now().isoformat(timespec="seconds")}
                self.estado["log"].append([datetime.now().strftime("%d/%m/%Y %H:%M:%S"), usuario,
                                           "Observaciones revisadas (sin editor)", k, "", "", "", comentario, ""])
            self.recalcular(k)

    def confirmar_lote(self, usuario, proveedores=None, periodos=None):
        """Sin editor: confirma todas las certificaciones LISTAS (o las elegidas) y genera PDF + Excel.
        Las que no se pueden confirmar se listan con el motivo."""
        hechas, pendientes, rutas = [], [], []
        for k, c in sorted(self.certs.items()):
            if proveedores and c["proveedor"] not in proveedores and not set(c.get("miembros", [])) & set(proveedores):
                continue
            if periodos and c["periodo"] not in periodos:
                continue
            if c["estado"] == "CONFIRMADA":
                continue
            ok, msg, rr = self.confirmar_cert(k, usuario)
            if ok:
                hechas.append(k)
                rutas += rr
            else:
                motivos = [o["msg"] for o in c["obs"] if o["bloq"]] or (
                    ["observaciones críticas sin revisar: revisalas en el editor"] if c["estado"] == "REVISAR" else [])
                pendientes.append((k, list(dict.fromkeys(motivos))))
        return hechas, pendientes, rutas

    def _siguiente(self):
        ks = [k for k, _ in self.w_sel.options]
        pend = [k for k in ks if self.certs[k]["estado"] != "CONFIRMADA"]
        if not pend:
            return self._msg("No quedan certificaciones pendientes en esta vista. 🎉")
        i = ks.index(self.key) if self.key in ks else -1
        sig = next((k for k in ks[i + 1:] if k in pend), pend[0])
        self.w_sel.value = sig

    # --------------------------------------------------------------- exportar
    def archivos_confirmados(self):
        """PDF y Excel de cada certificación confirmada (los que existen en disco)."""
        pdfs, xls = [], []
        for k, c in sorted(self.certs.items(), key=lambda x: (x[1]["numero"], x[0])):
            if c["estado"] != "CONFIRMADA":
                continue
            m = self.estado["meta"].get(k, {})
            if m.get("pdf") and Path(m["pdf"]).exists():
                pdfs.append(Path(m["pdf"]))
            if m.get("excel") and Path(m["excel"]).exists():
                xls.append(Path(m["excel"]))
        return pdfs, xls

    def descargar_archivos(self, rutas):
        """Descarga cada archivo por separado (sin zip)."""
        if not _colab_files:
            for r in rutas:
                print("   ", r)
            return
        import time
        for r in rutas:
            _colab_files.download(str(r))
            time.sleep(0.4)          # el navegador puede pedir 'Permitir descargas múltiples'

    def guardar_cierre(self):
        """Escribe el libro Excel por período (todos los proveedores) y el RESUMEN de control. Sin descargas."""
        anio, mes = self.estado["anio"], self.estado["mes"]
        libros = []
        for per in ("1Q", "2Q", "Mensual"):
            cs = [c for c in self.certs.values() if c["periodo"] == per]
            if cs:
                cs.sort(key=lambda c: c["proveedor"].casefold())
                ruta = self.salida / f"Certificaciones_{per}_{cb.MESES[mes]}_{anio}.xlsx"
                try:
                    cb.escribir_libro(ruta, cs, self.logo)
                except PermissionError:
                    print(f"⚠ {ruta.name} está abierto en Excel: cerralo y volvé a ejecutar para actualizarlo.")
                    continue
                libros.append(ruta)
        ruta_res = self.salida / f"RESUMEN_{cb.MESES[mes]}_{anio}.xlsx"
        try:
            cb.escribir_resumen(ruta_res, self.certs, self.globales, self.obs_cfg, self.cuad, self.estado["log"],
                                self.mapa_humedad)
        except PermissionError:
            print(f"⚠ {ruta_res.name} está abierto en Excel: cerralo y volvé a ejecutar para actualizarlo.")
        self.guardar_estado()
        return libros, ruta_res

    def exportar(self, pdf=True, excel=True, config_y_estado=True, otros=False, descargar=True):
        """Guarda todo en la carpeta de trabajo y descarga, archivo por archivo, solo lo elegido:
        PDF y/o Excel de cada certificación confirmada; config + estado (para seguir la numeración);
        otros = libro Excel por período (todos los proveedores) y RESUMEN de control."""
        libros, ruta_res = self.guardar_cierre()
        pdfs, xls = self.archivos_confirmados()
        a_bajar = (pdfs if pdf else []) + (xls if excel else [])
        if config_y_estado:
            a_bajar += [self.ruta_config, self.ruta_estado]
        if otros:
            a_bajar += libros + [ruta_res]
        print(f"✔ {len(pdfs)} certificaciones confirmadas: {len(pdfs)} PDF y {len(xls)} Excel")
        print(self.cuad[["Período", "TN maestro", "Ajustes manuales", "TN certificadas", "Diferencia", "Resultado"]]
              .to_string(index=False))
        if not a_bajar:
            print("No hay nada para descargar con las opciones elegidas.")
            return []
        print(f"\n⬇ Descargando {len(a_bajar)} archivos (si el navegador pregunta, elegí 'Permitir'):")
        for r in a_bajar:
            print("   ", Path(r).name)
        if descargar:
            self.descargar_archivos(a_bajar)
        return a_bajar
