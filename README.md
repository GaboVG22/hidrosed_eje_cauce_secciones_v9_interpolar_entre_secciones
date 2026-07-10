# HidroSed · Módulo Eje Cauce y Secciones v9

Versión corregida para revisión hidráulica del eje del cauce y secciones.

## Cambios v9

1. **Interpolación entre secciones con separación fija**
   - Nueva herramienta en `Ventana sección por km`.
   - Permite definir `km inicial`, `km final` y `separación entre secciones [m]`.
   - Genera una nueva serie de secciones intermedias interpolando progresivamente la forma `station-cota` entre las secciones vecinas aguas arriba y aguas abajo.
   - Puede reemplazar las secciones existentes del tramo o agregar solo secciones intermedias.
   - Permite usar semi-ancho interpolado entre secciones vecinas o semi-ancho fijo definido por el usuario.

2. **Se mantiene la interpolación entre curvas de nivel**
   - Útil cuando se cargan curvas de apoyo y se quiere reconstruir la sección por cruces sección-curva.

3. **Canal trapecial o rectangular por tramo**
   - Se puede insertar un canal prismático entre `km X` y `km Y`.
   - Permite tipo rectangular o trapecial.
   - Parámetros: ancho de fondo, altura, taludes izquierdo/derecho y Manning n.
   - La cota de fondo sigue el perfil longitudinal disponible, por lo que el canal se ajusta progresivamente al eje.

4. **Excel prismático visible**
   - Se mantiene la opción explícita para cargar planilla Excel prismática.

5. **Vista satelital de referencia**
   - Eje útil y secciones se proyectan sobre imagen satelital web.

## Entradas principales

- Eje del cauce KMZ/KML.
- PC hidrológico KMZ/KML.
- PC cuenca soporte KMZ/KML.
- DEM descargado desde OpenTopography con API Key o GeoTIFF manual.
- Perfil longitudinal opcional.
- Curvas de nivel de apoyo opcionales.
- Excel prismático opcional.

## Streamlit Cloud

Main file path:

```text
app.py
```

Python version:

```text
3.11
```
