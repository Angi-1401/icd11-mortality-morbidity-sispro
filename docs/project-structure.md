# Project Structure

The project is organized into several key directories and files, each serving a specific purpose in the overall functionality of the application.

## Directory Structure
```
/ICD11-Mortality-Morbidity-SISPRO
├── /docs                       ' Docs and project-related documents
├── /release                    ' Compiled releases of the application
├── /src                        ' Source code files for VBA modules
│   ├── /forms                  ' UserForm files
│   │   ├── frmProgress.frm
│   │   └── frmProgress.frx
│   └── /modules                ' Main modules
│       ├── ICD11.bas
│       ├── ReportOperations.bas
│       ├── TableOperations.bas
│       └── Utils.bas
├── README.md
└── LICENSE
```
## Considerations for Editing VBA Modules

VBA modules can be edited in any text or code editor, but edits made outside the VBA environment are not automatically applied to the Excel workbook/project.

To have those modules reflected in the project you must either:

- Manually import/add the module through the Excel VBA Editor (e.g., File → Import File...),
- Or use a third-party tool that supports importing/exporting VBA modules into the workbook/project.

If a module file is not imported into the VBA project, changes in external files will not affect the workbook.