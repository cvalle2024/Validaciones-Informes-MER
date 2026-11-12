# Validador MER: versión con historial de calidad del dato

## Inicio rápido
1. Instalar dependencias: `pip install -r requirements.txt`.
2. Ejecutar: `streamlit run app.py` **o** `streamlit run validación_informes_streamlit.py` (nombre de entrada anterior, ahora también actualizado).
3. Indicar **Mes de reporte por defecto** con formato `OCT 2026`, `NOV 2026`, etc. Es fundamental cuando una hoja de auditoría no tiene columna de mes.
4. Cargar archivos MER `.xlsx` o `.zip` y seleccionar **Procesar validaciones MER**.
5. Consultar el apartado **Historial de errores y seguimiento trimestral**, filtrar país/trimestre y descargar su Excel.

## Cambios de nombres
- La comparación del país ahora equipara «El Salvador», «El_Salvador» y variantes con espacios o tildes.
- Ignora encabezados repetidos incrustados en los datos de las hojas.
- Mantiene una auditoría estricta de nombres de establecimientos con sugerencias no automáticas; no renombra los archivos originales.
- Si no reconoce el país, señala **PAÍS NO RECONOCIDO** en vez de comparar con un país ajeno.

## Historial y reincidencia
- Cada vez que se pulsa **Procesar validaciones MER** se crea una ejecución en SQLite; actualizar filtros o descargar archivos no crea ejecuciones.
- La firma distingue archivo, país, establecimiento, hoja, tipo de error, fila y mes del reporte. La misma firma detectada nuevamente pasa a «2», «3», etc. Si un error cambia de lugar o de detalle, puede tratarse como un error distinto.
- **Detecciones** cuenta todas las apariciones en ejecuciones, mientras **Errores diferentes** cuenta firmas distintas; **Reincidencias** cuenta apariciones después de la primera.
- Los trimestres usan año fiscal MER: Q1=octubre–diciembre, Q2=enero–marzo, Q3=abril–junio, Q4=julio–septiembre. Ejemplo: octubre 2026 pertenece a FY27-Q1.
- El historial no guarda IDs/expedientes u otros datos clínicos de pacientes; solo metadatos operativos del error y huella SHA-256.
- **Auditoria_Sitio TX_CURR** se omite del historial porque es una tabla auxiliar del mismo proceso de conciliación y duplicaría el mismo tipo de error.
- El detalle completo del reporte individual sigue disponible en las exportaciones originales.

## Persistencia importante
El almacenamiento predeterminado es `data/historial_mer.sqlite3`. No borre esa carpeta ni su base de datos cuando actualice el código. Si ejecuta en un servidor con almacenamiento temporal, monte un volumen persistente y defina `MER_HISTORY_DB=/ruta/persistente/historial_mer.sqlite3`. Las ejecuciones en diferentes servidores NO compartirán historial si utilizan discos separados. Cree copias de seguridad periódicas del archivo SQLite (idealmente con la función de respaldo de SQLite).

La base comienza vacía: los reportes Excel antiguos no se importan automáticamente, evitando registrar como históricos falsos positivos de versiones anteriores. Puede volver a procesar los archivos corregidos para iniciar una línea base limpia.

## Si no ves el historial
- Comprueba que en la parte superior de la web figure **V8** y el aviso **Historial de errores y seguimiento trimestral**. Si aparece **V7**, sigue corriendo código antiguo.
- Si trabajas en Streamlit Community Cloud/GitHub, debes subir o desplegar los archivos actualizados del ZIP en el repositorio asociado y reiniciar la app. Descargar el ZIP por sí solo NO modifica una web publicada.
- El resumen de historial está **al final de la página principal**, debajo de los resultados del validador. Se puede desplazar hasta el final; aun con la base vacía, aparecerá el título y el mensaje de que no hay registros.
- Si ejecutas localmente, extrae el ZIP completo en otra carpeta y ejecuta desde esa carpeta, para evitar abrir archivos antiguos por error.
