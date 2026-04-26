export const SetupSheetSchema = {
    templateId: "19WMef0XLfSNkK48IQDQ1WUxeLa4Ir1LfvHZ4W5xEiIk",
    sheets: {
        groupData: {
            sheetName: "Registro Inicial",
            ranges: {
                groupName: "datos_grupo",
                attendancePerClass: "datos_asistenciaIndividual",
                averagePerField: "datos_promedioPorCampos",
                dateStart: "datos_fechaInicio",
                dateTrim1: "datos_trimestre1",
                dateTrim2: "datos_trimestre2",
                dateEnd: "datos_fechaFin",
                subjects: "datos_asignaturas",
                students: "datos_estudiantes",
            },
        },
        calendarTemplate: {
            sheetName: "Template_Calendario",
            ranges: {
                monthNames1: "cal_mesAltura1",
                monthNames2: "cal_mesAltura2",
                monthNames3: "cal_mesAltura3",
                trimester1Day: "cal_diaTrimestre1",
                trimester2Day: "cal_diaTrimestre2",
                trimester3Day: "cal_diaTrimestre3",
                restDay: "cal_diaDescanso",
            },
        },
        calendar: {
            sheetName: "Calendario",
            ranges: {},
        },
    },
} as const;

export const SetupSheetSchema2 = {
    templateId: "19WMef0XLfSNkK48IQDQ1WUxeLa4Ir1LfvHZ4W5xEiIk",
    sheetNames: {
        groupData: "Registro Inicial",
        calendarTemplate: "Template_Calendario",
        calendar: "Calendario",
    },
    namedRanges: {
        groupName: "datos_grupo",
        attendancePerClass: "datos_asistenciaIndividual",
        averagePerField: "datos_promedioPorCampos",
        dateStart: "datos_fechaInicio",
        dateTrimester1: "datos_trimestre1",
        dateTrimester2: "datos_trimestre2",
        dateEnd: "datos_fechaFin",
        subjects: "datos_asignaturas",
        students: "datos_estudiantes",
        monthNames1: "cal_mesAltura1",
        monthNames2: "cal_mesAltura2",
        monthNames3: "cal_mesAltura3",
        trimester1Day: "cal_diaTrimestre1",
        trimester2Day: "cal_diaTrimestre2",
        trimester3Day: "cal_diaTrimestre3",
        restDay: "cal_diaDescanso",
    },
} as const;

export const ReportSheetSchema = {
    templateId: "1XVPjUP8Kw6TS9MuHt6TE2z09woqUg5Pp_WLZxW3NblY",
    sheetNames: {
        persistentData: "_Persistente",
        attendanceTemplate: "_AsistenciaTemplate",
        studentTemplate: "_EstudianteTemplate",
        attendance: "Asistencia",
        summary: "Concentrado",
        status: "Estado",
    },
    namedRanges: {
        // Persistent Data Sheet
        attendancePerClass: "persAsistenciaPorAsignatura",
        averagePerField: "persPromedioPorCampo",
        protectData: "persProtDatos",
        protectSkills: "persProtHabilidades",
        protectComments: "persProtObservaciones",
        protectPeriod1: "persProtPeriodo1",
        protectPeriod2: "persProtPeriodo2",
        protectPeriod3: "persProtPeriodo3",
        dateStart: "persFechaInicio",
        dateTrim1: "persFechaTrim1",
        dateTrim2: "persFechaTrim2",
        dateEnd: "persFechaFin",
        subjects: "persAsignaturas",
        students: "persEstudiantes",
        calendarDates: "persCalendario",
        // Attendance Template Sheet
        subjectRow: "faltasMateria",
        attendanceStudentRow: "faltasEstudiante",
        monthNames1: "faltasMes1",
        monthNames2: "faltasMes2",
        monthNames5: "faltasMes5",
        dayNames: "faltasDias",
        formatAttendanceCell: "faltasValidFormCond",
        formatPeriod1: "faltasFormPer1",
        formatPeriod2: "faltasFormPer2",
        formatPeriod3: "faltasFormPer3",
    },
} as const;
