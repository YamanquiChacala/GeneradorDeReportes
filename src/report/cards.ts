import { decoratedTextWithCallback, headerIcon, iconImage } from "../common/gas-parts";
import { defineActionParameters, defineInputsSchema, InputType, ParamType } from "../common/gas-utils";
import { type ReportPersistentData, StudentRowType } from "../common/report-utils";
import { Icon } from "../common/utils";
import { onGenerateIndividualReport, onPushReportMenuCard } from "./callbacks";

export const ReportFileIdParam = defineActionParameters({
    reportFileId: ParamType.STRING,
} as const);

export const StudentReportInputs = defineInputsSchema({
    studentIndex: InputType.NUMBER,
    fileName: InputType.STRING,
} as const);

/**
 * Main Card for a report admin interface
 */
export function buildReportOptionsMainCard(): GoogleAppsScript.Card_Service.Card {
    const card = CardService.newCardBuilder().setHeader(headerIcon({ title: "Asistencia y Reportes", subtitle: "Montessori Chacala", iconName: Icon.CHART }));

    const spacer = CardService.newTextParagraph().setText("&nbsp;");
    const cardSection = CardService.newCardSection()
        .setHeader("Tablero de Administración")
        .addWidget(spacer)
        .addWidget(
            decoratedTextWithCallback({
                callback: onPushReportMenuCard.name,
                text: "Reportes",
                bottomText: "Generar reportes individuales o por grupo",
                startIcon: iconImage({ iconName: Icon.PAGE }),
                endIcon: iconImage({ iconName: Icon.FORWARD }),
            }),
        )
        .addWidget(spacer)
        .addWidget(
            decoratedTextWithCallback({
                callback: "", // TODO: Callback to card with details
                text: "Periodos y Calendario",
                bottomText: "Configurar fechas, permisos y ciclo actual.",
                startIcon: iconImage({ iconName: Icon.CALENDAR }),
                endIcon: iconImage({ iconName: Icon.FORWARD }),
            }),
        )
        .addWidget(spacer)
        .addWidget(
            decoratedTextWithCallback({
                callback: "", // TODO: Callback to card with details
                text: "Alumnos",
                bottomText: "Agregar, dar de baja o reordenar alumnos",
                startIcon: iconImage({ iconName: Icon.STUDENT }),
                endIcon: iconImage({ iconName: Icon.FORWARD }),
            }),
        )
        .addWidget(spacer)
        .addWidget(
            decoratedTextWithCallback({
                callback: "", // TODO: Callback to card with details
                text: "Materias y Áreas",
                bottomText: "Administrar el plan de estudios y posiciones",
                startIcon: iconImage({ iconName: Icon.BOOKS }),
                endIcon: iconImage({ iconName: Icon.FORWARD }),
            }),
        )
        .addWidget(spacer)
        .addWidget(
            decoratedTextWithCallback({
                callback: "", // TODO: Callback to card with details
                text: "Ponderaciones",
                bottomText: "Ajustar cálculo de califiaciones general y por materia",
                startIcon: iconImage({ iconName: Icon.WEIGHTS }),
                endIcon: iconImage({ iconName: Icon.FORWARD }),
            }),
        );

    const peekHeader = headerIcon({ title: "Tablero de Administración", subtitle: "Asistencia y Reportes", iconName: Icon.CHART });

    return card.addSection(cardSection).setPeekCardHeader(peekHeader).build();
}

/**
 * Report creation card
 */
export function reportCreateReportsCard(persistentData: ReportPersistentData): GoogleAppsScript.Card_Service.Card {
    const card = CardService.newCardBuilder().setHeader(headerIcon({ title: "Generación de Reportes", iconName: Icon.PAGE }));

    // Generate all reports

    const allReportsSection = CardService.newCardSection().setHeader("👥 &nbsp; Reporte Grupal");

    const groupExplanation = CardService.newTextParagraph().setText("⚠️Los reportes se guardarán en la misma carpeta que el archivo actual.");

    const generateAllAction = CardService.newAction().setFunctionName("");
    const generateAllButtonSet = CardService.newButtonSet().addButton(
        CardService.newTextButton().setText("🗂️ - Generar todos los reportes").setTextButtonStyle(CardService.TextButtonStyle.FILLED).setOnClickAction(generateAllAction),
    );

    const spacer = CardService.newTextParagraph().setText("<br><br>");

    allReportsSection.addWidget(spacer).addWidget(generateAllButtonSet).addWidget(spacer).addWidget(groupExplanation);

    // Generate individual reports

    const individualReportSection = CardService.newCardSection().setHeader("👤 &nbsp; Reporte Individual");

    const studentDropdown = CardService.newSelectionInput()
        .setType(CardService.SelectionInputType.DROPDOWN)
        .setTitle("🎓 Selecciona un alumno (Requerido)")
        .setFieldName(StudentReportInputs.fieldName("studentIndex"));

    persistentData.students.forEach((studentRow, index) => {
        let fullName = "";
        let value = "";
        if (studentRow.type === StudentRowType.STUDENT) {
            fullName = `${studentRow.firstName} ${studentRow.lastName}`;
            value = index.toString();
        }
        studentDropdown.addItem(fullName, value, false);
    });

    const fileNameInput = CardService.newTextInput()
        .setFieldName(StudentReportInputs.fieldName("fileName"))
        .setTitle("📝 Nombre del archivo")
        .setHint("Opcional: Si se deja en blanco, se usará el nombre del alumno");

    const generateIndividualAction = CardService.newAction()
        .setFunctionName(onGenerateIndividualReport.name)
        .addRequiredWidget(StudentReportInputs.fieldName("studentIndex"));
    const generateIndividualButton = CardService.newTextButton().setText("🎓 - Generar reporte").setOnClickAction(generateIndividualAction);

    // Agregar los widgets a la sección individual
    individualReportSection.addWidget(studentDropdown).addWidget(fileNameInput).addWidget(generateIndividualButton);

    return card.addSection(allReportsSection).addSection(individualReportSection).build();
}
