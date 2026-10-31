<div align="center" width="100%">

![Sistema de Reportes](images/spiral_256.png)

</div>

[![codecov](https://codecov.io/github/YamanquiChacala/GeneradorDeReportes/graph/badge.svg?token=3H7BTF98TM)](https://codecov.io/github/YamanquiChacala/GeneradorDeReportes)

# Sistema de Reportes

Complemento interno de Google Workspace para Drive y Google Sheets que ayuda a administrar la asistencia, las evaluaciones y los reportes escolares de cada grupo.

El trabajo se organiza en dos archivos:

1. 📋 **Registro Inicial**: configuración del grupo, estudiantes, asignaturas, áreas formativas y calendario.
2. 📊 **Reporte**: hoja de cálculo diaria para docentes y administradores, con asistencia, seguimiento del grupo y una hoja por estudiante.

Desde el Reporte, el complemento también permite generar archivos PDF individuales o para todo el grupo.

> **Estado del proyecto:** la administración de estudiantes y varias acciones de fechas que aparecen en el complemento todavía no están implementadas. Estarán disponibles en futuras versiones.

## Contenido

- [Roles](#roles)
- [Cómo obtener el complemento](#cómo-obtener-el-complemento)
- [Flujo de trabajo](#flujo-de-trabajo)
- [Guía para administradores](#guía-para-administradores)
- [Guía para docentes](#guía-para-docentes)
- [Hojas del archivo Reporte](#hojas-del-archivo-reporte)
- [Generación de reportes PDF](#generación-de-reportes-pdf)
- [Funciones pendientes](#funciones-pendientes)
- [Créditos y licencia](#créditos-y-licencia)

## Roles

### Administradores

Preparan el Registro Inicial desde Drive, crean el Reporte desde ese archivo y administran las funciones disponibles del complemento en la hoja de cálculo. También generan los reportes en PDF y consultan el estado de captura del grupo.

### Docentes

Trabajan directamente en el archivo Reporte. No necesitan instalar ni abrir el complemento: utilizan las hojas de asistencia y las hojas individuales para registrar asistencia, evaluaciones, habilidades y comentarios.

![Archivo Reporte](images/screenshots/Report_full.png)

En la imagen puede verse el archivo Reporte, donde los maestros pueden tomar asistencia, y el complemento para los administradores a la derecha.

## Cómo obtener el complemento

El complemento es **interno y exclusivo para la escuela**. Los administradores pueden instalarlo desde Google Workspace Marketplace, en la sección **Aplicaciones Internas**. No es un complemento público.

![Google Marketplace, aplicaciones internas](images/screenshots/Marketplace_internal.png) | ![Google Marketplace, Generador de Reportes](images/screenshots/Markteplace_install.png)
:---:|:---:
![Permisos de la app 1](images/screenshots/oauth_1.png) | ![Permisos de la app 2](images/screenshots/oauth_2.png)

Una vez instalado, se abre desde Drive para crear un Registro Inicial y desde Google Sheets para trabajar con un Registro Inicial o un Reporte. La primera vez que se use, Google puede solicitar autorización para que el complemento acceda a los archivos necesarios.

## Flujo de trabajo

```text
Drive
  └─ Crear un Registro Inicial
       └─ Completar estudiantes, asignaturas, áreas y calendario
            └─ Abrir el Registro Inicial en Sheets
                 └─ Crear el archivo Reporte
                      ├─ Administrador: configura periodos y permisos
                      └─ Docentes: registran información en las hojas
                           └─ Administrador: genera los PDF
```

## Guía para administradores

### 1. Crear el Registro Inicial desde Drive

1. Instala el complemento interno, si aún no está instalado.
2. En Drive, selecciona la carpeta donde quieras guardar los archivos del grupo. También puedes seleccionar un archivo de esa carpeta.
3. Abre el complemento y completa el formulario **Registro Inicial de Grupos**.
4. Indica el nombre del grupo, si se registrará asistencia individual por asignatura y si se calcularán promedios por campo formativo.
5. Define el primer día de clases, los cierres del primer y segundo periodo y el último día del ciclo escolar.
6. Crea el Registro Inicial.

Los archivos de trabajo se guardan en la carpeta elegida. El Registro Inicial puede duplicarse desde Sheets para preparar otro grupo con una configuración similar.

![Creación de Registro Inicial](images/screenshots/Drive_setup.png)

### 2. Completar y revisar el Registro Inicial

Abre el archivo creado en Google Sheets y completa la información del grupo en las tablas preparadas para:

- Nombre del grupo y opciones generales.
- Estudiantes.
- Asignaturas y su organización por áreas formativas.
- Fechas del ciclo escolar y calendario.

Revisa el calendario generado y marca los días que no sean escolares, según corresponda. Si cambiaste las fechas del ciclo, puedes usar **Regenerar Calendario** desde el complemento; al regenerarlo se borrarán los días festivos o excepciones que se hayan marcado manualmente.

![Registro Inicial, información del grupo](images/screenshots/Setup_Form.png) | ![Registro Inicial, calendario](images/screenshots/Setup_Calendar.png) | ![Registro Inicial, complemento](images/screenshots/Setup.png)
:---:|:---:|:---:

### 3. Crear el archivo Reporte

Con el Registro Inicial abierto en Sheets:

1. Abre el complemento.
2. Verifica los datos del grupo y el calendario.
3. Usa **Crear Archivo de Calificaciones**.

El complemento crea el Reporte a partir de la configuración del grupo y lo guarda junto al Registro Inicial. Incluye una hoja por estudiante, hojas de asistencia y consulta, y hojas internas ocultas que el sistema necesita para funcionar.

![Pestañas del Reporte](images/screenshots/Report_Sheets.png)

### 4. Administrar el Reporte

Abre el archivo Reporte y usa el complemento de Sheets para las acciones disponibles, entre ellas:

- **Cambiar el periodo actual.** Al avanzar de periodo se actualiza el periodo mostrado en las hojas. Existe una opción para reiniciar las habilidades y los comentarios; al marcarla se borrarán esos datos de todos los estudiantes.
- **Proteger secciones.** Se pueden bloquear o desbloquear las habilidades, los comentarios y las calificaciones de cada periodo para controlar qué partes pueden editar los docentes.
- **Generar reportes PDF** individuales o para todo el grupo.

Comprueba el periodo y las protecciones antes de que los docentes comiencen a capturar información.

![Reporte, menú principal](images/screenshots/Report_MainMenu.png) | ![Reporte, generación de reportes](images/screenshots/Report_GenerateReports.png) | ![Reporte, manejo de periodos](images/screenshots/Report_Periods.png)
:--:|:--:|:--:

## Guía para docentes

Los docentes utilizan el archivo Reporte compartido por la escuela. **No necesitan instalar el complemento**: su trabajo se realiza completamente en Google Sheets.

Según las secciones que el administrador haya dejado editables, pueden:

- Registrar asistencia en la hoja **Asistencia**.
- Capturar evaluaciones y habilidades en la hoja individual de cada estudiante.
- Escribir comentarios y retroalimentación para el estudiante.
- Consultar **Concentrado** para dar seguimiento al grupo.
- Consultar **Estado** para identificar información que aún falta.

Las celdas protegidas no deben modificarse. Si una sección necesaria está bloqueada, el docente debe pedir al administrador que revise las protecciones.

## Hojas del archivo Reporte

Las pestañas visibles para el trabajo cotidiano son:

| Hoja | Uso |
| --- | --- |
| **Asistencia** | Registro de asistencia del grupo según el calendario configurado. Puede incluir asistencia individual por asignatura si se activó esa opción al crear el Registro Inicial. |
| **Concentrado** | Vista general de resultados y promedios del grupo. |
| **Estado** | Seguimiento de la información capturada y detección de datos pendientes por estudiante antes de generar reportes. |
| **Una hoja por estudiante** | Registro individual de información, habilidades, comentarios, evaluaciones y datos del estudiante. |

El archivo también contiene pestañas internas ocultas: **_Persistente**, **_AsistenciaTemplate** y **_EstudianteTemplate**. No las desocultes, renombres ni edites manualmente: hacerlo puede afectar el funcionamiento del Reporte.

<details>
<summary>Las diferentes pestañas del Reporte</summary>

#### Asistencia
![Asistencia](images/screenshots/Report_Assistance.png)
#### Concentrado
![Concentrado](images/screenshots/Report_Summary.png)
#### Estado
![Estado](images/screenshots/Report_Status.png)
#### Hoja de estudiante
![Estudiante](images/screenshots/Report_Student.png)

</details>

## Generación de reportes PDF

Desde la sección **Reportes** del complemento abierto en el Reporte, el administrador puede:

- **Generar un reporte individual:** seleccionar al estudiante y, si lo desea, indicar un nombre personalizado para el archivo.
- **Generar todos los reportes:** iniciar la generación en segundo plano para el grupo completo.

Los PDF se guardan en la misma carpeta de Drive que el archivo Reporte. La generación grupal puede tardar varios minutos; espera a que aparezcan los archivos antes de iniciar otro proceso.

![Reporte, generación de reportes](images/screenshots/Report_GenerateReports.png) | ![Reporte en PDF](images/screenshots/Report_pdf.png)
:---:|:---:

## Funciones pendientes

Las siguientes opciones corresponden a funciones planeadas, pero **todavía no están listas en la versión actual del código**:

- **Administrar estudiantes desde el Reporte:** agregar, eliminar o reordenar estudiantes usando el complemento. El tablero muestra una entrada de Alumnos, pero esa sección aún no ejecuta estas acciones.
- **Agregar o quitar fechas desde el Reporte:** los controles aparecen en el menú, pero sus acciones todavía no están conectadas.
- **Actualizar las fechas del ciclo escolar desde el Reporte:** la opción está visible, pero aún no guarda cambios.
- **Administrar materias, áreas formativas y ponderaciones desde el Reporte:** las entradas del tablero están previstas, pero todavía no tienen una pantalla funcional.

Mientras esas funciones no estén disponibles, configura los estudiantes, las asignaturas, las áreas y las fechas iniciales en el Registro Inicial antes de crear el Reporte.

## Créditos y licencia

- **Autor:** [Yamanqui García Rosales](https://github.com/Yamanqui)
- **Repositorio:** [YamanquiChacala/GeneradorDeReportes](https://github.com/YamanquiChacala/GeneradorDeReportes)
- **Licencia:** [MIT](./LICENSE)
