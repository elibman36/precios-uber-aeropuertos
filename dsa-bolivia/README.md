# DSA Bolivia con bloque externo integrado

Análisis de sostenibilidad de deuda (2025-2035) que integra la deuda pública del SPNF con la balanza de pagos, las reservas y un tipo de cambio endógeno.

- `DSA_Bolivia.xlsx`: el modelo con fórmulas vivas. El escenario se elige en `Portada!C6` (del 1 al 8) y los insumos están en azul en la hoja `Supuestos`.
- `modelo.py`: el motor con las mismas ecuaciones. Corre los escenarios, calibra el balance primario y el escenario "Réplica FMI", y hace el Monte Carlo (5.000 sendas).
- `construir_excel.py`: regenera el Excel (`python construir_excel.py DSA_Bolivia.xlsx 1`).

Validación: para los 8 escenarios, las hojas del Excel recalculadas en LibreOffice coinciden con `modelo.py`, con diferencias máximas de alrededor de 1e-14.

Fecha de los datos: 14-oct-2026. Las fuentes están en la hoja `Fuentes`. Varias cifras vienen de prensa que cita al FMI, al BCB o al INE, así que hay que verificarlas contra los documentos originales.
