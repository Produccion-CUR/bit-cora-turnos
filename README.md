# Bitácora de Turnos 📋

Sistema digital de bitácora para planta frutícola — Curicó.

## Funciones

- **Lavado de bandejas** — inicio/cierre de turno, pallets por tipo
- **Selección** — dotación por línea (L1, L3, L4, L5), materiales, indicadores
- **Envasado** — SKU, MTC bolsa/caja, codificación por cliente, cajas producidas
- **Programa de producción** — grilla semanal por línea/turno/especie
- **Insumos y Consumo** — necesito vs. tengo en piso de planta
- **Verificador de hora** — código de bloque horario para codificación
- **Modo Jefe** (clave: `produccion2026`) — indicadores, dashboard, programa

## Despliegue rápido

### 1. Clonar y publicar en GitHub

```bash
git init
git add .
git commit -m "inicial"
git remote add origin https://github.com/TU_USUARIO/bitacora-turnos.git
git push -u origin main
```

### 2. Activar GitHub Pages

1. Ve a **Settings → Pages** en tu repositorio
2. En **Source**, selecciona **GitHub Actions**
3. El workflow se ejecuta automáticamente al hacer push a `main`
4. En 1–2 minutos la app estará disponible en `https://TU_USUARIO.github.io/bitacora-turnos/`

### 3. Conectar Google Sheets (opcional, para sincronización)

1. Abre el archivo `Bitacora_GoogleSheets.xlsx` en Google Drive → Guardar como Google Sheet
2. Ve a **Extensiones → Apps Script** en ese Sheet
3. Pega el contenido de `Code_Bitacora.gs` y guarda
4. Haz clic en **Implementar → Nueva implementación**
   - Tipo: **Aplicación web**
   - Ejecutar como: **Yo**
   - Acceso: **Cualquier persona**
5. Copia la URL del script
6. En `src/main.jsx`, reemplaza la línea:
   ```js
   window.__SHEETS_URL__ = "";
   ```
   con:
   ```js
   window.__SHEETS_URL__ = "https://script.google.com/macros/s/TU_ID/exec";
   ```
7. Haz commit y push → el workflow redesplega automáticamente

## Desarrollo local

```bash
npm install
npm run dev
```

La app queda disponible en `http://localhost:5173`

## Datos

Los datos se guardan en **localStorage del navegador** (por dispositivo). Con Google Sheets configurado, se sincronizan entre todos los dispositivos que usen la misma URL de Apps Script.

## Claves de acceso

| Pantalla | Clave |
|----------|-------|
| Modo Jefe de producción | `produccion2026` |
