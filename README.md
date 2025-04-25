# Limpiador de Datos Automático

Esta aplicación de escritorio está diseñada para automatizar el proceso de limpieza y formateo de datos en archivos Excel, con un enfoque especial en:

- Limpieza y normalización de números telefónicos dominicanos
- Formateo de fechas
- Generación de columnas de comentarios

## Características Principales

### 1. Limpieza de Teléfonos
- Normalización de números telefónicos
- Eliminación de duplicados
- Validación de prefijos dominicanos (809, 829, 849)
- Relleno automático de celdas vacías

### 2. Formateo de Fechas
- Conversión automática a formato dd/mm/yyyy
- Soporte para múltiples formatos de entrada
- Previsualización antes de aplicar cambios
- Respaldo automático de datos originales

### 3. Generador de Comentarios
- Concatenación personalizable de columnas
- Soporte para montos en RD$ y US$
- Reemplazo de palabras configurable
- Previsualización de resultados

## Requisitos
- Python 3.x
- pandas
- tkinter
- numpy

## Instalación
1. Clonar el repositorio
```bash
git clone https://github.com/varocode/limpiar-data-automatic.git
```

2. Instalar dependencias
```bash
pip install pandas numpy
```

## Uso
1. Ejecutar el archivo principal:
```bash
python limpieza_telefonos_gui.py
```

2. Seleccionar el archivo Excel a procesar
3. Elegir las columnas a limpiar/formatear
4. Seguir las instrucciones en pantalla

## Características Adicionales
- Interfaz gráfica intuitiva
- Logs detallados del proceso
- Respaldo automático de datos originales
- Fusión opcional con base de datos original

## Contribuir
Las contribuciones son bienvenidas. Por favor, abre un issue para discutir los cambios propuestos. 