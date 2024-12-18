# Generador de Tablas de Asignación de Personal

Este programa genera una tabla de asignación de personal para fechas específicas y exporta los datos a un archivo Excel. Está diseñado para asignar personas de manera equitativa en diferentes roles, utilizando reglas de exclusión y equilibrio.

---

## **Requisitos previos**

1. **Python**: Asegúrate de tener Python 3.8 o superior instalado.
2. **Dependencias**:
   - `pandas`
   - `openpyxl`
3. **Instalación de librerías**: Puedes instalar las dependencias ejecutando:
   ```bash
   pip install pandas openpyxl
   ```

---

## **Uso**

1. **Clonar el repositorio**:
   ```bash
   git clone <URL_DEL_REPOSITORIO>
   cd <CARPETA_DEL_REPOSITORIO>
   ```

2. **Editar listas de nombres**:
   - Abre el archivo del programa y edita las siguientes variables para incluir los nombres correspondientes:
     - `nombres_mesa`: Lista de nombres para la columna "Mesa".
     - `nombres_plataforma`: Lista de nombres para las columnas "Plataforma", "Micros 1", y "Micros 2".
   - Si hay fechas específicas con restricciones, edítalas en `nombres_fecha_especial` con el formato `{fecha: [nombres]}`.

3. **Ejecutar el programa**:
   Ejecuta el script con:
   ```bash
   python <nombre_del_archivo>.py
   ```

4. **Salida**:
   - Se generarán dos archivos en la ubicación especificada:
     - `Listado_de_sonido.xlsx`: Tabla con las asignaciones de personal.
     - `Conteos_de_sonido.xlsx`: Estadísticas de asignación por persona y por rol.

---

## **Reglas de asignación**

1. **Equilibrio**:
   - Se priorizan las personas con menor número de asignaciones.
   - Cada persona es asignada equitativamente entre los diferentes roles.

2. **Restricciones**:
   - Ninguna persona puede ser asignada al mismo rol en días consecutivos.
   - Se consideran restricciones por fecha para asignaciones específicas.

3. **Formato de fechas**:
   - Las fechas deben estar en formato `dd/mm/yyyy`.

---

## **Personalización**

### Cambiar rango de fechas:
   - Edita las variables `fecha_inicio` y `fecha_fin` en el programa.

### Cambiar días seleccionados:
   - El script considera únicamente martes (1) y domingos (6). Puedes modificar esta lógica en el siguiente bloque:
     ```python
     if fecha_inicio.weekday() in [1, 6]:  # Modifica los números según los días deseados
     ```

---

## **Contribuciones**

Si deseas contribuir con mejoras o reportar problemas, envía un pull request o crea un issue en el repositorio.

---

## **Licencia**

Este proyecto está bajo la licencia MIT. Consulta el archivo `LICENSE` para más detalles.
