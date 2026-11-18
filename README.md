# Artikos · Banmédica
Aplicación local en PowerShell para consultar reportes SQL, validar/copiar ZIP, preparar volúmenes y revisar la integridad. No realiza envíos.

## Requisitos
- Windows con Windows PowerShell 5.1.
- Acceso autorizado al servidor SQL configurado para la aplicación.
- Microsoft Excel instalado para generar el XLSX (el CSV sigue disponible si Excel falla).
- WinRAR instalado para probar, extraer y recomprimir archivos ZIP.

## Inicio
Ejecuta `Start.bat`. El registro de inicio se guarda en `Logs/inicio.log` y la auditoría de procesos, en formato JSON Lines, en `Logs/auditoria_procesos.jsonl`. Las carpetas locales de ejecución también reciben `auditoria_procesos.jsonl`.

La auditoría registra eventos, fechas, usuario/equipo, identificadores de sesión y operación, resultados y duración. Los logs y preferencias locales se excluyen de Git para evitar publicar datos de uso o configuraciones sensibles.

## Publicar cambios
El historial del proyecto se administra con Git en la rama `main`. Revisa cada cambio con `git status` y `git diff` antes de confirmarlo. Para enviar los commits a GitHub se requiere configurar un remoto `origin` y autenticar la cuenta autorizada.
