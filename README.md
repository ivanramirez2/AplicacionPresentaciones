# 🚇 Servicios Metro Ligero

> Aplicación web para optimizar la gestión de servicios e incidencias en el metro ligero.

---

## 📝 Descripción

**Servicios Metro Ligero** es una aplicación web desarrollada en **ASP.NET MVC** con **SQL Server**, diseñada para ser utilizada en **tablets por conductores** del metro ligero (líneas **ML2** y **ML3**). Permite iniciar sesión de forma segura, consultar servicios e incidencias del día, e imprimir documentos rápidamente. Esta solución reemplaza procesos manuales, agilizando operaciones y reduciendo errores en horas punta.

---

## ✨ Características

- 🔐 **Inicio de Sesión Seguro**  
  Autenticación con número de empleado y PIN.

- 📅 **Consulta de Servicios**  
  Visualización de servicios asignados con detalles (hora, línea, ruta, estado).

- 🚨 **Gestión de Incidencias**  
  Filtros por fecha y línea para revisar incidencias activas.

- 🖨️ **Impresión Rápida**  
  Generación de documentos imprimibles (PDF/Excel) desde la tablet.

- 📱 **Interfaz Responsive**  
  Diseño adaptado a tablets con navegación táctil e intuitiva.

- ⚙️ **Alta Disponibilidad**  
  Soporte para múltiples sesiones concurrentes durante horas punta.

---

## 🧰 Tecnologías

- **Backend:** ASP.NET MVC, C#  
- **Frontend:** Razor, Bootstrap, JavaScript, jQuery  
- **Base de Datos:** SQL Server  
- **Librerías:** iTextSharp (PDF), OpenXML (Excel)  
- **Seguridad:** HTTPS, PIN, gestión de sesiones  
- **Despliegue:** Windows Server + IIS  

---

## 📋 Requisitos

- 🖥️ Servidor con **Windows Server** e **IIS** configurado  
- 🗃️ **SQL Server** 2016 o superior  
- ⚙️ **.NET Framework 4.8** o superior  
- 🌐 Navegador moderno (Chrome o Edge) en tablets  
- 🖨️ Impresora conectada para funcionalidad de impresión  

---

## 🚀 Instalación

### 1. Clona el repositorio

```bash
git clone https://github.com/ivanramirez2/AplicacionPresentaciones
