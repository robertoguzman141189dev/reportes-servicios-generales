# 🚀 Guía Rápida de Configuración

## Tiempo estimado: 15-20 minutos

---

## Paso 1: Crear Repositorio en GitHub (2 min)

1. Ve a https://github.com/new
2. Nombre: `reportes-servicios-generales`
3. ✅ Public
4. ✅ Add README
5. Click **Create repository**

---

## Paso 2: Crear Service Account en Google Cloud (5 min)

1. Ve a https://console.cloud.google.com/
2. Crea un proyecto nuevo (o usa uno existente)
3. Habilita la **Google Drive API**:
   - Menú → APIs & Services → Library
   - Busca "Google Drive API" → Enable

4. Crea Service Account:
   - Menú → IAM & Admin → Service Accounts
   - **+ Create Service Account**
   - Nombre: `reportes-automation`
   - Click **Create and Continue**
   - Skip roles → **Done**

5. Genera la clave JSON:
   - Click en el service account creado
   - **Keys** tab → **Add Key** → **Create new key**
   - Tipo: **JSON**
   - Se descarga automáticamente

6. Guarda el contenido del JSON (lo necesitarás después)

---

## Paso 3: Configurar Google Drive (2 min)

1. Crea una carpeta en Drive: `Reportes-Servicios-Generales`
2. Click derecho → **Share**
3. Comparte con el email del Service Account:
   - Busca el email en el JSON descargado (campo `client_email`)
   - Ejemplo: `reportes-automation@tu-proyecto.iam.gserviceaccount.com`
   - Permiso: **Editor**

4. Copia el **ID de la carpeta** de la URL:
   ```
   https://drive.google.com/drive/folders/XXXXXXXXXXXXXXXXXXXXXXXXX
                                          ↑ Este es el FOLDER_ID
   ```

---

## Paso 4: Configurar Secrets en GitHub (3 min)

1. Ve a tu repo → **Settings** → **Secrets and variables** → **Actions**
2. Click **New repository secret** para cada uno:

| Nombre | Valor |
|--------|-------|
| `GOOGLE_CREDENTIALS` | Todo el contenido del JSON descargado |
| `DRIVE_FOLDER_ID` | El ID de la carpeta de Drive |

---

## Paso 5: Subir los archivos al repo (3 min)

```bash
# Clona el repo
git clone https://github.com/TU_USUARIO/reportes-servicios-generales.git
cd reportes-servicios-generales

# Copia todos los archivos de la carpeta reportes-automation/
# (los que te generé)

# Sube los cambios
git add .
git commit -m "Setup inicial del sistema de reportes"
git push
```

---

## Paso 6: Configurar GitHub Pages (1 min)

1. Ve a **Settings** → **Pages**
2. Source: **Deploy from a branch**
3. Branch: `main` | Folder: `/docs`
4. **Save**

---

## Paso 7: Configurar Google Apps Script (5 min)

1. Ve a https://script.google.com
2. **New Project**
3. Borra el código existente
4. Pega el contenido de `google-apps-script/Code.gs`
5. **Edita estas líneas** con tus datos:
   ```javascript
   GITHUB_TOKEN: 'ghp_XXXXXXXXXX',  // Tu token (ver abajo)
   GITHUB_OWNER: 'tu-usuario',
   GITHUB_REPO: 'reportes-servicios-generales',
   DRIVE_FOLDER_ID: 'XXXXXXXXXX',
   ```

6. Para crear el GitHub Token:
   - Ve a https://github.com/settings/tokens?type=beta
   - **Generate new token**
   - Nombre: `reportes-script`
   - Repository: tu repo
   - Permissions: Contents (Read/Write), Actions (Read/Write)
   - **Generate** y copia el token

7. En Apps Script, click en **Run** → `testGitHubConnection`
   - Autoriza los permisos
   - Verifica que diga "Conexión exitosa!"

8. Configura el Trigger:
   - Click en ⏰ **Triggers**
   - **+ Add Trigger**
   - Function: `checkForNewFiles`
   - Event: Time-driven → Minutes timer → Every minute
   - **Save**

---

## ✅ ¡Listo!

### Prueba el sistema:

1. Sube tu archivo Excel a la carpeta de Google Drive
2. Espera ~1 minuto
3. Ve a **Actions** en tu repo para ver el progreso
4. El reporte estará en: `https://TU_USUARIO.github.io/reportes-servicios-generales/`

---

## 🔧 Troubleshooting

### El Action no se dispara
- Verifica que el token de GitHub tenga los permisos correctos
- Revisa los logs en Apps Script → Executions

### Error de credenciales de Google
- Asegúrate de pegar TODO el JSON en el secret
- Verifica que el service account tenga acceso a la carpeta

### El reporte no se actualiza
- Revisa GitHub Actions → tu workflow → logs
- Verifica que el archivo Excel tenga el formato correcto

---

## 📞 URLs Importantes

- **Tu reporte:** `https://TU_USUARIO.github.io/reportes-servicios-generales/`
- **GitHub Actions:** `https://github.com/TU_USUARIO/reportes-servicios-generales/actions`
- **Google Apps Script:** `https://script.google.com`
- **Google Cloud Console:** `https://console.cloud.google.com`
