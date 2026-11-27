# Sistema de Bodega de Emergencias (programa de escritorio)

Programa para Windows escrito en Python (ventanas Tkinter), empaquetado como `.exe` con PyInstaller.
No necesita internet. Los datos se guardan como archivos JSON en `Documentos\BodegaEmergencias`:

- `datos\AAAA-MM_SERVICIO.json`: uno por cada mes y servicio
- `modelos\SERVICIO.json`: lista modelo de insumos de cada servicio
- `matrices\`: matrices por período (JSON y PDF)
- `documentos\`: documentos PDF generados

Este repositorio **no contiene datos**. Al abrir el programa por primera vez se importa el respaldo del programa anterior.
