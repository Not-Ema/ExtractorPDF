# Extractor de Datos PDF a Excel/CSV v0.1.0

Una aplicación robusta y escalable para extraer datos de cupones en PDFs, soportando tanto texto digital como imágenes escaneadas.

## Arquitectura Modular

La aplicación está estructurada en carpetas organizadas para mayor mantenibilidad:

- `config/`: Configuración centralizada.
  - `config.json`: Patrones de extracción, formatos de salida y ajustes.
- `src/`: Código fuente modular.
  - `logger.py`: Sistema de logging centralizado con salida a archivo y consola.
  - `pdf_processor.py`: Manejo de extracción de texto de PDFs con procesamiento concurrente.
  - `data_extractor.py`: Extracción de campos específicos usando patrones regex y parsing de cupones.
  - `data_writer.py`: Escritura de datos a Excel o CSV, con soporte para append.
  - `gui.py`: Interfaz gráfica de usuario con barra de progreso y logging en tiempo real.
- `tests/`: Pruebas del código.
- `run.py`: Punto de entrada principal que inicializa la aplicación.

## Características

- **Extracción Robusta**: Soporte para PDFs con texto digital.
- **Manejo de Cupones**: Parsing especial para textos con categorías y contenido (fondos azul/blanco).
- **Codificación UTF-8**: Soporte completo para caracteres españoles (tildes, ñ).
- **Escalabilidad**: Procesamiento concurrente de múltiples PDFs usando ThreadPoolExecutor.
- **Configurabilidad**: Patrones de extracción y ajustes en archivo JSON.
- **Flexibilidad de Salida**: Soporte para Excel (.xlsx) y CSV.
- **Interfaz Intuitiva**: GUI moderna con feedback en tiempo real.
- **Logging Detallado**: Registros de actividades para debugging y monitoreo.
- **Manejo de Errores**: Timeouts y excepciones para prevenir crashes.

## Requisitos

- Python 3.x
- Librerías: ver `requirements.txt`

Instalar dependencias:
```bash
pip install -r requirements.txt
```

## Uso

1. Ejecutar la aplicación:
    ```bash
    python run.py
    ```

2. Seleccionar carpeta con PDFs.
3. Elegir archivo de salida (Excel o CSV).
4. Iniciar extracción.

## Configuración

Editar `config.json` para:
- Modificar patrones regex para extracción de campos.
- Cambiar formato de salida (xlsx/csv).
- Ajustar número de workers para procesamiento concurrente.
- Configurar nivel de logging.

## Campos Extraídos

- Cliente
- Identificación
- Contrato
- Dirección
- Valor a Pagar
- No. Solicitud
- No. Rel. Pago
- Tipo de Cupón
- Válido hasta
- Código de Barras
- Archivo Origen

## Escalabilidad

- Procesamiento concurrente reduce tiempo para lotes grandes.
- Logging a archivo permite monitoreo de procesos largos.
- Módulos separados facilitan extensiones futuras (e.g., más formatos, APIs).

## Manejo de Errores

- Validación de entradas y existencia de archivos.
- Reintentos y mensajes de error informativos.
- Logging de excepciones para debugging.

## Desarrollo

Para extender:
- Agregar nuevos patrones en `config.json`.
- Implementar nuevos extractores en `data_extractor.py`.
- Añadir soporte para más formatos en `data_writer.py`.# ExtractorPDF
