# Medios Magnéticos Nómina JMC

App en Streamlit para validar:

- **CC-nóminas** = pagado
- **PCP0** = contabilizado

## Flujo

1. Cargar archivos de **CC-nóminas**
2. Elegir **Año** según `Fecha de pago`
3. Elegir uno o varios **Períodos para nómina**
4. Cargar archivo de **conceptos**
5. Cargar archivos de **PCP0**
6. Oprimir **Procesar validación**

## Qué genera

- Resumen pagado en CC
- Resumen contabilizado en PCP0
- Comparativo por empleado
- Comparativo por período
- Detalle de diferencias por concepto
- Excel final
- ZIP con CSV y logs

## Archivo de conceptos

Debe tener estas columnas:

- `CC-nómina`
- `Texto expl.CC-nómina`
- `Tipo`

En `Tipo` usa solamente:

- `Salarial`
- `Beneficio`
- `No aplica`

## Despliegue en Streamlit Cloud

Sube al repositorio estos archivos:

- `app.py`
- `requirements.txt`

Luego en Streamlit Cloud selecciona:

- repositorio
- branch `main`
- main file path: `app.py`
