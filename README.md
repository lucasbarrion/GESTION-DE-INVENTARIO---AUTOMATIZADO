📦 Sistema de Gestión de Inventario Automatizado (Google Sheets + Apps Script)

Sistema web de gestión de inventario y movimientos desarrollado con Google Apps Script, Google Sheets y HTML/CSS/JavaScript.

Permite:

Alta de productos

Control de stock

Registro de ingresos, egresos y gastos

Dashboard financiero

Inventario valorizado

Alertas de stock bajo o sin stock

🧠 Cómo funciona (resumen corto)

Google Sheets → base de datos

Google Apps Script (.gs) → backend (lógica y cálculos)

HTML (index.html) → frontend (interfaz visual)

Se publica como Aplicación Web

🚀 Instalación completa (paso a paso, en orden)
PASO 1 — Crear el Google Sheet

Crear un Google Sheet vacío

Ir a Extensiones → Apps Script

Se abre el editor de Apps Script

PASO 2 — Backend (PEGAR TODO JUNTO)

📍 Dónde pegar
En Apps Script → archivo Código.gs (o el que exista por defecto)

📋 Qué pegar
Copiar y pegar TODO el código backend (el archivo .gs completo)

👉 Este código:

Crea las hojas automáticamente

Maneja productos, stock, movimientos y gastos

Calcula dashboard y alertas de stock

📌 Luego de pegar:

Guardar el proyecto

Ejecutar manualmente esta función una sola vez:

crearHojasSistema()


Esto crea las hojas:

Productos

MOVIMIENTOS

GASTOS PROOVEDOR

PASO 3 — Frontend (PEGAR TODO JUNTO)

📍 Dónde pegar
En Apps Script:

Archivo → Nuevo → HTML

Nombre del archivo:

index


📋 Qué pegar
Copiar y pegar TODO el contenido del archivo index.html (HTML + CSS + JS juntos)

📌 Importante:

No separar CSS ni JS

No borrar google.script.run

Todo va en un solo archivo HTML

PASO 4 — Publicar la aplicación

En Apps Script:

Implementar → Nueva implementación

Tipo:

Aplicación web

Configuración:

Ejecutar como: Yo

Acceso: Cualquiera

Implementar

Copiar la URL final

👉 Esa URL es el sistema funcionando.

⚙️ Configuración importante (backend)

Dentro del backend existe este bloque:

const CONFIG = {
  SHEET_PRODUCTOS: 'Productos',
  SHEET_MOVIMIENTOS: 'MOVIMIENTOS',
  SHEET_GASTOS: 'GASTOS PROOVEDOR',
  CATEGORIAS: ['Mate', 'Bombilla', 'Yerba', 'Accesorio'],
  STOCK_BAJO_UMBRAL: 3
};


🔧 Qué podés modificar:

CATEGORIAS: rubros del negocio

STOCK_BAJO_UMBRAL: cantidad mínima para alerta

⚠️ Los nombres de hojas deben coincidir con el Sheet.

⚠️ Alertas de stock

Un producto entra en alerta si:

Stock = 0

Stock ≤ STOCK_BAJO_UMBRAL

El backend ya incluye:

function getAlertasStock() { ... }


Solo requiere ser llamado desde el frontend.

🔒 Seguridad

Este repositorio NO incluye:

IDs reales de Sheets

Credenciales

Datos privados

Cada usuario debe:

Crear su propio Google Sheet

Publicar su propia Web App

📌 Notas técnicas

Sin frameworks

100% Apps Script

Comunicación frontend/backend con:

google.script.run

👤 Autor

Desarrollado por Lucas Barrionuevo
Proyecto de automatización – Gestión de inventario

🧩 Licencia

Uso libre con fines educativos y demostrativos.
