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
            ranges: {
                start: "realCal_InicioEpoch",
                calendar: "realCal_CalendarioCompleto",
            },
        },
    },
} as const;

export const ReportSheetSchema = {
    templateId: "1XVPjUP8Kw6TS9MuHt6TE2z09woqUg5Pp_WLZxW3NblY",
    sheets: {
        persistentData: {
            sheetName: "_Persistente",
            ranges: {
                attendancePerClass: "pers_AsistenciaPorAsignatura",
                averagePerField: "pers_PromedioPorCampo",
                protectData: "pers_ProtDatos",
                protectSkills: "pers_ProtHabilidades",
                protectComments: "pers_ProtObservaciones",
                protectPeriod1: "pers_ProtPeriodo1",
                protectPeriod2: "pers_ProtPeriodo2",
                protectPeriod3: "pers_ProtPeriodo3",
                dateStart: "pers_FechaInicio",
                dateTrim1: "pers_FechaTrim1",
                dateTrim2: "pers_FechaTrim2",
                dateEnd: "pers_FechaFin",
                fields: "pers_CamposFormativos",
                subjects: "pers_Asignaturas",
                students: "pers_Estudiantes",
                calendarDates: "pers_Calendario",
            },
        },
        attendanceTemplate: {
            sheetName: "_AsistenciaTemplate",
            ranges: {
                subjectRow: "faltas_Materia",
                attendanceStudentRow: "faltas_Estudiante",
                monthNames1: "faltas_EtiquetaMes1",
                monthNames2: "faltas_EtiquetaMes2",
                monthNames5: "faltas_EtiquetaMes5",
                dayNames: "faltas_EtiquetaDias",
                formatAttendanceCell: "faltas_ValidacionFormatoGeneral",
                formatPeriod1: "faltas_FormatoPeriodo1",
                formatPeriod2: "faltas_FormatoPeriodo2",
                formatPeriod3: "faltas_FormatoPeriodo3",
            },
        },
    },
} as const;
