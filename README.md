# 🎓 Asignador de Calificaciones SIIUPP

Aplicación de escritorio en Python para transferir automáticamente calificaciones desde hojas de cálculo de Excel (`.xlsx`) hacia el portal de calificaciones de la Universidad Politécnica de Pachuca (SIIUPP).

---

## 🔒 Principio de Seguridad
- **Solo pegar:** La aplicación únicamente escribe las calificaciones en los campos correspondientes del parcial seleccionado (`Parcial 1`, `Parcial 2` o `Parcial 3`).
- **Control humano total:** La aplicación **NO hace clic en Guardar ni Enviar**, ni modifica asistencias ni otras columnas. Tú revisas que todo coincida y guardas manualmente en Chrome.

---

## 🚀 Inicio Rápido

1. Haz doble clic en el archivo **`ejecutar.bat`** (o ejecuta en consola `python app.py`).
2. Se abrirá la ventana del programa.

---

## 📋 Pasos de Uso en la Aplicación

1. **Paso 1: Chrome (SIIUPP)**
   - Haz clic en **"🚀 Abrir / Conectar Chrome"**.
   - Se abrirá la ventana de Chrome en el portal de UPP (o se conectará si ya lo tienes abierto).
   - Inicia sesión con tus credenciales y selecciona el grupo al que vas a calificar hasta llegar a la pantalla de captura de calificaciones.
   - En la aplicación, haz clic en **"🔄 Detectar Alumnos"**. Verás la confirmación de cuántos alumnos se detectaron (ej. *24 alumnos detectados*).

2. **Paso 2: Excel y Parcial**
   - Haz clic en **"📂 Seleccionar Excel..."** y elige tu archivo `.xlsx`.
   - Selecciona el **Parcial** a calificar (*Parcial 1*, *Parcial 2* o *Parcial 3*).
   - La aplicación seleccionará automáticamente la pestaña del Excel correspondiente y detectará la columna **"Total"**.
   - En la tabla central verás una vista previa comparativa de cada alumno con la nota que le corresponde.

3. **Paso 3: Pegar Calificaciones**
   - Haz clic en **"✍️ PEGAR CALIFICACIONES EN CHROME"**.
   - Las calificaciones se escribirán automáticamente en las casillas correspondientes y el sistema de UPP calculará los promedios en vivo.
   - Revisa en tu pantalla de Chrome y haz clic en Guardar cuando estés conforme.

