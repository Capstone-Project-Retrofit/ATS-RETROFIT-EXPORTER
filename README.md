# ATS Exporter

Herramienta para **exportar los datos del módulo ATS** (ESP32 + firmware V4) por USB‑Serial y generar un reporte `REPORTE_ATS.xlsx`.

> **Solo Windows. No se instala nada** (sin Python, sin Excel). Todo corre con PowerShell nativo.

---

## Uso

1. Conecta el módulo a la laptop con el cable USB.
2. **Cierra el Monitor Serial de Arduino** (para liberar el puerto COM).
3. **Doble clic a `exportar_ats.bat`.**

La herramienta captura los datos, los guarda en `Documentos\ATS_EXPORT\<fecha_hora>\` y genera + abre el **`REPORTE_ATS.xlsx`** (hojas: Resumen, Eventos, Sucesos).

> *Watchdog:* nunca se cuelga. Si algo falla, los datos crudos quedan guardados igual — no se pierde nada. Si el reporte no abriera, vuelve a dar doble clic a `exportar_ats.bat`.

---

## Si no funciona

- **"No se detectó el módulo"** → revisa que el cable esté conectado y que el firmware esté corriendo.
- **"No se pudo abrir COMx"** → cierra el Monitor Serial de Arduino (o lo que use el puerto) y reintenta.

---

## Archivos

| Archivo | Para qué |
|---|---|
| `exportar_ats.bat` | **La herramienta — dale doble clic.** |
| `exportar_ats.ps1` | Captura los datos del módulo. |
| `Generar-ReporteATS.ps1` | Arma el `REPORTE_ATS.xlsx`. |
| `plantilla_reporte_ats.xlsx` | Plantilla del reporte. |

---

## Descarga

https://github.com/Capstone-Project-Retrofit/ATS-RETROFIT-EXPORTER — botón verde **Code → Download ZIP**, descomprime y listo.

---

*INTEC · Módulo Retrofit No Invasivo para ATS Legacy.*
