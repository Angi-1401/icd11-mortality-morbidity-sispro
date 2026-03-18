# Estructura del Proyecto

El proyecto está organizado en varios directorios y archivos clave, cada uno con un propósito específico en la funcionalidad general de la aplicación.

## Estructura de Directorios

```
/ICD11-Mortality-Morbidity-SISPRO
├── /docs                       ' Documentos y archivos relacionados con el proyecto
├── /release                    ' Versiones compiladas de la aplicación
├── /src                        ' Archivos fuente de los módulos VBA
│   ├── /forms                  ' Archivos de formularios de usuario
│   │   ├── frmProgress.frm
│   │   └── frmProgress.frx
│   └── /modules                ' Módulos principales
│       ├── ICD11.bas
│       ├── ReportOperations.bas
│       ├── TableOperations.bas
│       └── Utils.bas
├── README.md
└── LICENSE
```

## Consideraciones para Editar Módulos VBA

Los módulos VBA pueden editarse en cualquier editor de texto o código, pero los cambios realizados fuera del entorno VBA no se aplican automáticamente al libro/proyecto de Excel.

Para que esos módulos se reflejen en el proyecto debe:

- Importar/agregar manualmente el módulo a través del Editor de VBA de Excel (por ejemplo, Archivo → Importar archivo...),
- O usar una herramienta de terceros que permita importar/exportar módulos VBA al libro/proyecto.

Si un archivo de módulo no se importa al proyecto VBA, los cambios en archivos externos no afectarán al libro.
