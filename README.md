# Control de Rechazos FME · V7

Prototipo Streamlit para automatizar la gestión de rechazos GNOC, contacto con FME, seguimiento de feedback en OWS y casos Closed.

## Archivos que debes cargar

1. `RECHAZOS Maria.xlsx`
2. `WO Result Query.xlsx`
3. `DATA 2026 - COORDINADORES.xlsx`
4. `CUADRILLAS ACT. NUEVO PROY. 2026.xlsx`

## Qué hace V7

- Detecta rechazos por `Audit status = Reject` y por el mes/año de `Audit time`.
- Obtiene coordinador con la misma prioridad del BUSCARX usado en Excel:
  1. `Hoja2!D:D` → `Hoja2!B:B`
  2. `Hoja2!E:E` → `Hoja2!B:B`
  3. `Hoja1!A:A` → `Hoja1!B:B`
- Busca el celular del FME en `CUADRILLAS ACT. NUEVO PROY. 2026.xlsx`:
  - primero por la cédula incluida en `FME ID`;
  - como respaldo por nombre normalizado.
- Genera un botón **Abrir WhatsApp** con el número y el mensaje precargado.
- Permite registrar manualmente el feedback que se realiza en OWS:
  - Fecha de feedback
  - Estado: Pendiente / Aprobado / Rechazado / Otro
  - Feedback o comentario dejado en OWS
- Guarda esos datos en `feedback_ows.json` y la vista cambia automáticamente entre **Sin feedback** y **Con feedback**.
- Incluye vista filtrable por coordinador, FME, sitio, estado de actividad y estado OWS, con resumen dinámico por coordinador.
- Gestiona actividades `Closed`, solicitando únicamente la fecha/hora de cierre consultada en OWS.
- Genera correo `.eml` listo para Outlook con destinatarios Para y CC configurados.
- Exporta `RECHAZOS Maria` actualizado, incluyendo celular y seguimiento OWS.

## Correos Closed

**Para**
- griselda.dolores.castillo@eysglobal.co
- camilo.jimenez@huawei.com

**CC**
- sebastian.rolong_costa@manpowercolombia.com
- ali.villa_costa@manpowercolombia.com

## Ejecutar

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Nota sobre persistencia

El prototipo guarda el seguimiento OWS en `feedback_ows.json` dentro de la carpeta del proyecto. Si se despliega en un servidor, conviene reemplazarlo por SQLite o una base de datos persistente.
