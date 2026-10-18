# PEF Concesión Vial — App Streamlit

Versión en Python del Plan Económico Financiero (modelo **PEF_X.xlsx**), con una app web
para cambiar los supuestos y ver al instante TIR, VAN, payback, flujo de fondos, impuestos,
amortizaciones y préstamo.

El motor (`pef_model.py`) replica las fórmulas del Excel. Validado contra los valores
calculados por Excel en PEF_I a PEF_X: diferencia de $0 en todas las filas del flujo e
impuestos, y TIR, TIR modificada, VAN y payback idénticos.

## Contenido

| Archivo | Qué es |
|---|---|
| `app.py` | La app Streamlit (interfaz) |
| `pef_model.py` | Motor de cálculo del PEF |
| `data/PEF_X.xlsx` | Modelo que se carga por defecto |
| `requirements.txt` | Dependencias |
| `tools/validar_contra_excel.py` | Compara Python vs. Excel para cualquier PEF |
| `.streamlit/config.toml` | Tema y tamaño máximo de archivo |

## Correr en tu computadora

```bash
pip install -r requirements.txt
streamlit run app.py
```

Se abre en el navegador (http://localhost:8501).

## Publicar en Streamlit Community Cloud (gratis)

1. Creá un repositorio en GitHub y subí **todos** los archivos de esta carpeta, respetando
   las subcarpetas (`data/`, `tools/`, `.streamlit/`).
2. Entrá a https://share.streamlit.io con tu cuenta de GitHub.
3. **Create app** → elegí el repositorio, la rama (`main`) y como archivo principal `app.py`.
4. **Deploy**. En uno o dos minutos la app queda online con un link para compartir.

Si el modelo es confidencial, dejá el repositorio **privado**. Streamlit Cloud permite apps
desde repos privados, y podés limitar quién la ve desde *Settings → Sharing*.

## Cómo se usa

- **Barra lateral:** tránsito, tarifa, crecimiento, sensibilidades, préstamo, impuestos,
  amortización y garantías. Con **Restablecer** se vuelve a los valores del archivo.
- **Cargar otro PEF:** cualquier Excel con la misma estructura (hojas CONTROL, FLUJO,
  INPUTS, HIPOTESIS, IMPUESTOS, PRESTAMO). Se leen sus datos y se recalcula en Python.
- **Pestaña CAPEX y OPEX:** editás los montos anuales y los factores como en una planilla.
- **Pestaña Sensibilidad:** TIR del proyecto para combinaciones de tarifa y tránsito, y el
  efecto de ±10% en cada variable.
- **Pestaña Descargar:** exporta todos los resultados a Excel.

## Criterios del modelo

- Montos en pesos con IVA. El préstamo se calcula como un % del CAPEX total y entra en el año 1.
- La app muestra **dos TIR**: la del Excel (que incluye el préstamo) y la **del proyecto sin
  préstamo**, que es la que conviene informar cuando el crédito distorsiona el flujo del año 1.
- Las amortizaciones usan la versión automática: cualquier obra, en cualquier año, se amortiza
  sin pasar el fin de la concesión.
- Se mantienen a propósito algunas particularidades del Excel para que los números coincidan:
  el factor de nuevos peajes del año 5 se arrastra a los años siguientes; el préstamo forma
  parte del flujo; Sellos se calcula sobre el CAPEX total en 5 años; y en el escenario de
  15 años las garantías de los años 16 a 20 siguen entrando en el VAN.
