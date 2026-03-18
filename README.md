# Informe de Mortalidad y Morbilidad ICD-11 para SISPRO

Esta herramienta basada en Excel/VBA está diseñada para analizar informes clínicos y hospitalarios generados por SISPRO (Sistema de Información de Salud), el sistema público de información en salud utilizado en Venezuela. Proporciona información automatizada sobre las condiciones de salud regionales identificando y resumiendo las principales causas de mortalidad y morbilidad.

## Características

- **Análisis de las 25 principales causas de muerte**  
  Extrae y clasifica automáticamente las 25 principales causas subyacentes de muerte de los informes SISPRO utilizando criterios personalizables.
- **Panel interactivo**  
  Genera un panel dinámico para visualizar y resumir los principales indicadores de salud.
- **Integración con la API ICD-11**  
  Se conecta a la API oficial de ICD-11 para recuperar descripciones de cada código diagnóstico, mejorando la claridad y precisión del informe.
- **Optimizado para grandes volúmenes de datos**  
  Algoritmos eficientes manejan conjuntos de datos extensos (más de 20,000 filas) con alto rendimiento.
- **Acceso seguro a la API**  
  Implementa autenticación en dos pasos para el acceso seguro a la API ICD-11.

## Requisitos

- **Sistema Operativo**: Windows 10 o superior
- **Versión de Excel**: Microsoft Excel 2010 o superior
- **Soporte de Macros**: Las macros deben estar habilitadas
- **Credenciales de API**: Se requiere un `CLIENT_ID` y `CLIENT_SECRET` válidos del [portal de la API ICD-11](https://icd.who.int/icdapi)

## Instalación

1. Descargue el repositorio como un archivo ZIP y extráigalo en la ubicación deseada. Alternativamente, clone el repositorio usando Git:
   ```bash
   git clone https://github.com/Angi-1401/icd11-mortality-morbidity-sispro.git
   ```
2. Abra el archivo `ICD11_Mortality_Morbidity_SISPRO.xlsb` en Microsoft Excel.
3. Acceda al módulo `ICD11.bas` en el editor de VBA (alt + F11) para ingresar sus credenciales de la API ICD-11.
   ```vba
   Const CLIENT_ID As String = "TU_CLIENT_ID_AQUI"
   Const CLIENT_SECRET As String = "TU_CLIENT_SECRET_AQUI"
   ```
4. Guarde el libro para conservar sus credenciales.

## Primeros Pasos

1. Abra el libro de Excel y habilite las macros cuando se le solicite.
2. Cargue su informe SISPRO en la hoja.
3. Ejecute la macro para analizar los datos y generar el panel.

## Notas

- Asegúrese de tener conectividad a internet para el acceso a la API.
- La herramienta está adaptada para formatos de informes SISPRO; no se admiten otras fuentes de datos.

## Contacto

Para preguntas o soporte, por favor contacte al responsable del proyecto.

---

© 2026 – Informe de Mortalidad y Morbilidad ICD-11 para SISPRO