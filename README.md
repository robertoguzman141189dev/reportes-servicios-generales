# 🚀 Automatización de Reportes - Servicios Generales

Sistema automatizado para generar reportes ejecutivos desde Google Drive y publicarlos en GitHub Pages en tiempo real.

## 📋 Requisitos Previos

- Cuenta de GitHub
- Cuenta de Google (para Drive)
- Python 3.9+ (para desarrollo local)

## 🏗️ Arquitectura

```
Google Drive (Excel) 
    → Google Apps Script (detecta cambios) 
    → GitHub Actions (webhook) 
    → Genera HTML 
    → GitHub Pages (publica)
```

## 📁 Estructura del Proyecto

```
reportes-servicios-generales/
├── .github/
│   └── workflows/
│       └── generate-report.yml    # GitHub Action principal
├── scripts/
│   ├── process_excel.py           # Procesa Excel y genera HTML
│   ├── requirements.txt           # Dependencias Python
│   └── template.html              # Template del reporte
├── google-apps-script/
│   └── Code.gs                    # Script para Google Drive
├── docs/                          # Carpeta de GitHub Pages (output)
│   └── index.html                 # Reporte generado
├── .gitignore
└── README.md
```

## 🔧 Instalación Paso a Paso

### Paso 1: Crear Repositorio en GitHub

1. Ve a https://github.com/new
2. Nombre: `reportes-servicios-generales`
3. Selecciona "Public" (necesario para GitHub Pages gratis)
4. ✅ Add README
5. Click "Create repository"

### Paso 2: Configurar GitHub Pages

1. Ve a Settings → Pages
2. Source: "Deploy from a branch"
3. Branch: `main` → `/docs`
4. Save

### Paso 3: Crear Personal Access Token (PAT)

1. Ve a https://github.com/settings/tokens?type=beta
2. "Generate new token"
3. Nombre: `reportes-automation`
4. Expiration: 90 days (o más)
5. Repository access: "Only select repositories" → tu repo
6. Permissions:
   - Contents: Read and Write
   - Actions: Read and Write
7. Generate token
8. **¡COPIA EL TOKEN!** (solo se muestra una vez)

### Paso 4: Configurar Google Apps Script

1. Ve a https://script.google.com
2. Nuevo proyecto → pega el código de `google-apps-script/Code.gs`
3. Reemplaza `TU_GITHUB_TOKEN` y `TU_USUARIO`
4. Deploy → New deployment → Web app
5. Execute as: "Me"
6. Who has access: "Anyone"
7. Deploy y copia la URL del webhook

### Paso 5: Configurar Trigger en Google Drive

1. En Google Apps Script, ve a "Triggers" (⏰)
2. Add Trigger:
   - Function: `checkForNewFiles`
   - Event source: Time-driven
   - Type: Minutes timer
   - Interval: Every minute
3. Save

### Paso 6: Subir archivos al repositorio

Clona el repo y sube todos los archivos de este proyecto.

```bash
git clone https://github.com/TU_USUARIO/reportes-servicios-generales.git
cd reportes-servicios-generales
# Copia todos los archivos aquí
git add .
git commit -m "Initial setup"
git push
```

## 🎯 Uso

1. **Sube tu Excel** a la carpeta designada en Google Drive
2. **Espera ~1 minuto** (el trigger detecta el archivo)
3. **Revisa GitHub Actions** para ver el progreso
4. **Accede al reporte** en `https://TU_USUARIO.github.io/reportes-servicios-generales/`

## 🔗 URLs Importantes

- **Reporte publicado:** `https://TU_USUARIO.github.io/reportes-servicios-generales/`
- **GitHub Actions:** `https://github.com/TU_USUARIO/reportes-servicios-generales/actions`
- **Google Apps Script:** `https://script.google.com`

## 🐛 Troubleshooting

### El reporte no se genera
- Verifica que el archivo Excel tenga el formato correcto
- Revisa los logs en GitHub Actions
- Confirma que el token de GitHub no haya expirado

### Error de permisos en Google Drive
- Asegúrate de que el script tenga acceso a Drive
- Revisa los triggers en Google Apps Script

## 📞 Soporte

Si tienes problemas, revisa los logs en:
1. GitHub Actions → Tu workflow → Logs
2. Google Apps Script → Executions

---
Desarrollado para la Dirección de Servicios Generales UPGCH 2026
