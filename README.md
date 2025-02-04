# 🔥 Cálculo de Transformadores Sobrecargados

**Versión:** 3  
**Desarrollador:** Juan Ortiz (2025)  
**Metodología de Cálculo:** Bryan Estrella  
**Data de Cálculo:** Jenny Guashpa  
**Empresa:** Consorcio SIG-ELECTRIC  

## 📖 Descripción  
Este script de MATLAB calcula los **transformadores sobrecargados** con base en la metodología sugerida por Bryan Estrella. Se asigna un **estrato de consumo** a cada usuario, desde A hasta E, y se determina el más frecuente. Luego, según el número de usuarios, se consulta su consumo en la **tabla DMD de la Empresa Eléctrica Quito**.  

El cálculo toma en cuenta:  
- **Pérdidas del 3.6%** en el sistema  
- **Capacidad máxima del transformador:** hasta **1.25 veces** su capacidad nominal  

## 📂 Estructura del Código  
El script sigue estos pasos principales:  
1. **Carga de datos** desde archivos Excel  
2. **Filtrado y preprocesamiento** de datos de transformadores, postes y medidores  
3. **Categorización de consumos** en los estratos A - E  
4. **Cálculo de demanda total** con base en el consumo de los usuarios  
5. **Cálculo de potencia de luminarias** asociadas a cada transformador  
6. **Comparación con la capacidad máxima del transformador** para detectar sobrecargas  
7. **Generación de resultados** en un archivo Excel  

## 📑 Requisitos  
- MATLAB (versión recomendada: R2021a o superior)  
- Archivos de entrada en formato `.xlsx` con los siguientes datos:  
  - `TRAFO`: Información de los transformadores  
  - `POSTE`: Potencias de luminarias asociadas  
  - `MEDIDORES`: Consumo de usuarios  

## ⚙️ Uso  
1. **Modificar la variable `Alimentador`** con el identificador correcto  
2. **Asegurar que los archivos de entrada están en la carpeta correcta**  
3. **Ejecutar el script en MATLAB**  
4. **Revisar el archivo de salida** generado en la carpeta de resultados  

## 📌 Notas  
- El código puede requerir ajustes si la estructura de los archivos de entrada cambia  
- La función `F_leerDatosExcel()` debe estar disponible en la carpeta `F_Funciones`  
- Si un transformador no se encuentra en la base de datos, se generará un mensaje de advertencia  

---

### 📧 Contacto  
Si tienes dudas o sugerencias, puedes contactarme en juan.ortiz.e@hotmail.com.  

🚀 _Este proyecto está en desarrollo, ¡cualquier contribución es bienvenida!_  
