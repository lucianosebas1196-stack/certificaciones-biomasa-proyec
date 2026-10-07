# Certificaciones de biomasa

Cuaderno de Google Colab que arma las certificaciones de entrega de biomasa por proveedor (1ª quincena, 2ª quincena y mensual) a partir del maestro **R-O-08-01 – Control ingreso biomasa** (hoja *Datos Entrada*).

## Qué hace

- Lee el maestro (`.xlsb` o `.xlsx`) y agrupa los ingresos por proveedor y período.
- Controla errores de tipeo: fechas fuera de secuencia, pesos imposibles, remitos repetidos, CUIT, precios y nombres de proveedor.
- Aplica los descuentos de calidad de la planilla del laboratorio (hoja *GENERAL*):
  - **Granulometría por camión** (finos < #04): hasta 10 % sin descuento; de 10 a 15 % se descuenta lo que excede el 10 %; más de 15 % se descuentan todos los finos. Lo descontado se paga como ASERRÍN.
  - **Humedad por camión**, solo para proveedores con `DESCUENTA HUMEDAD = SI`.
  - Formato propio para GyG (humedad y granulometría camión por camión).
- Abre un editor (ipywidgets) para revisar, corregir y agregar remitos, con registro de cambios (quién, cuándo, motivo).
- Al confirmar asigna número correlativo y genera el **PDF** y el **Excel** de cada certificación.
- Exporta un Excel por período con todos los proveedores y un RESUMEN de control (cuadre contra el maestro).

## Archivos

| Archivo | Contenido |
|---|---|
| `Certificaciones_biomasa.ipynb` | Cuaderno para Colab. Es autosuficiente: sus celdas del Paso 2 escriben los dos módulos de abajo. |
| `cert_biomasa.py` | Motor: lectura, validaciones, descuentos, armado de certificaciones, PDF y Excel. |
| `editor_colab.py` | Editor interactivo (selector de período/proveedor, revisión y confirmación). |

Los `.py` son copias de las celdas `%%writefile` del cuaderno, para poder leerlos y comparar versiones en GitHub. Si cambiás el código, mantené las dos copias iguales.

## Uso

1. Abrí `Certificaciones_biomasa.ipynb` en Google Colab.
2. Ejecutá las celdas en orden.
3. En el Paso 3 subí el maestro `.xlsb` (obligatorio) y, si los tenés, la planilla de calidad, `config_proveedores.xlsx` y `estado_AAAA-MM.json` para retomar.

Uso local (fuera de Colab):

```bash
pip install -r requirements.txt
jupyter notebook Certificaciones_biomasa.ipynb
```

## Configuración

Los datos propios de la empresa y de los proveedores **no están en el código**. Se cargan en `config_proveedores.xlsx`:

- **Hoja Proveedores:** razón social, CUIT, grupo de precios, qué descuentos se aplican, unificaciones y separación por patente.
- **Hoja Precios:** precio por TN de cada grupo y tipología (con tramos opcionales).
- **Hoja Parametros:**
  - número inicial de certificación y punto de venta;
  - día de corte de la quincena y tolerancias de finos;
  - `EMPRESA_LINEA_1..3`: encabezado con los datos de la empresa;
  - `PIE_1..3`: pie de la certificación (reclamos, a quién enviar la factura).

El logo se toma de la certificación modelo que se sube en el Paso 3.

Si no se sube un config, el cuaderno crea uno nuevo con valores de ejemplo para completar.

## Datos

Los archivos de trabajo (maestro, planilla de calidad, config, estado, PDF y Excel generados) **no se suben al repositorio**: están excluidos en `.gitignore`.
