import { coloredParagraph, decoratedTextWithCallback, headerIcon, iconImage, spaceBetweenWidgets } from "../common/gas-parts";
import { Colors, defineActionParameters, defineInputsSchema, InputType, ParamType } from "../common/gas-utils";
import { Period, type ReportPersistentData, StudentRowType } from "../common/report-utils";
import { Icon } from "../common/utils";
import { onGenerateAllReports, onGenerateIndividualReport, onPushReportMenuCard } from "./callbacks";

export const ReportFileIdParam = defineActionParameters({
    reportFileId: ParamType.STRING,
} as const);

/**
 * Main Card for a report admin interface
 */
export function buildReportOptionsMainCard(): GoogleAppsScript.Card_Service.Card {
    const card = CardService.newCardBuilder().setHeader(headerIcon({ title: "Asistencia y Reportes", subtitle: "Montessori Chacala", iconName: Icon.CHART }));

    const cardSection = CardService.newCardSection()
        .setHeader("Tablero de Administración")
        .addWidget(spaceBetweenWidgets())
        .addWidget(
            decoratedTextWithCallback({
                callback: onPushReportMenuCard.name,
                text: "Reportes",
                bottomText: "Generar reportes individuales o por grupo",
                startIcon: iconImage({ iconName: Icon.PAGE }),
                endIcon: iconImage({ iconName: Icon.FORWARD }),
            }),
        )
        .addWidget(spaceBetweenWidgets())
        .addWidget(
            decoratedTextWithCallback({
                callback: "", // TODO: Callback to card with details
                text: "Periodos y Calendario",
                bottomText: "Configurar fechas, permisos y ciclo actual.",
                startIcon: iconImage({ iconName: Icon.CALENDAR }),
                endIcon: iconImage({ iconName: Icon.FORWARD }),
            }),
        )
        .addWidget(spaceBetweenWidgets())
        .addWidget(
            decoratedTextWithCallback({
                callback: "", // TODO: Callback to card with details
                text: "Alumnos",
                bottomText: "Agregar, dar de baja o reordenar alumnos",
                startIcon: iconImage({ iconName: Icon.STUDENT }),
                endIcon: iconImage({ iconName: Icon.FORWARD }),
            }),
        )
        .addWidget(spaceBetweenWidgets())
        .addWidget(
            decoratedTextWithCallback({
                callback: "", // TODO: Callback to card with details
                text: "Materias y Áreas",
                bottomText: "Administrar el plan de estudios y posiciones",
                startIcon: iconImage({ iconName: Icon.BOOKS }),
                endIcon: iconImage({ iconName: Icon.FORWARD }),
            }),
        )
        .addWidget(spaceBetweenWidgets())
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
 * The field inputs that are sent to create individual report.
 */
export const SingleStudentReportInputs = defineInputsSchema({
    studentIndex: InputType.NUMBER,
    fileName: InputType.STRING,
} as const);

/**
 * Report creation card
 */
export function buildReportCreateReportsCard(persistentData: ReportPersistentData): GoogleAppsScript.Card_Service.Card {
    const card = CardService.newCardBuilder().setHeader(headerIcon({ title: "Generación de Reportes", iconName: Icon.PAGE }));

    // Generate all reports

    const allReportsSection = CardService.newCardSection().setHeader("👥 &nbsp; Reporte Grupal");

    const groupExplanation = CardService.newTextParagraph().setText("⚠️Los reportes se guardarán en la misma carpeta que el archivo actual.");

    const generateAllAction = CardService.newAction().setFunctionName(onGenerateAllReports.name);
    const generateAllButtonSet = CardService.newButtonSet().addButton(
        CardService.newTextButton().setText("🗂️ - Generar todos los reportes").setTextButtonStyle(CardService.TextButtonStyle.FILLED).setOnClickAction(generateAllAction),
    );

    const spacer = CardService.newTextParagraph().setText("<br><br>");

    allReportsSection.addWidget(spacer).addWidget(generateAllButtonSet).addWidget(spacer).addWidget(groupExplanation);

    // Generate individual reports

    const individualReportSection = CardService.newCardSection().setHeader("👤 &nbsp; Reporte Individual");

    const studentDropdown = CardService.newSelectionInput()
        .setType(CardService.SelectionInputType.DROPDOWN)
        .setTitle("🎓 Selecciona un Alumno")
        .setFieldName(SingleStudentReportInputs.fieldName("studentIndex"));

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
        .setFieldName(SingleStudentReportInputs.fieldName("fileName"))
        .setTitle("📝 Nombre del archivo")
        .setHint("Opcional: Si se deja en blanco, se usará el nombre del alumno");

    const generateIndividualAction = CardService.newAction()
        .setFunctionName(onGenerateIndividualReport.name)
        .addRequiredWidget(SingleStudentReportInputs.fieldName("studentIndex"));
    const generateIndividualButton = CardService.newTextButton().setText("🎓 - Generar reporte").setOnClickAction(generateIndividualAction);

    // Agregar los widgets a la sección individual
    individualReportSection.addWidget(studentDropdown).addWidget(fileNameInput).addWidget(generateIndividualButton);

    return card.addSection(allReportsSection).addSection(individualReportSection).build();
}

/**
 * Periods and Calendar Card
 */
export function buildPeriodsAndCalendarCard(persistentData: ReportPersistentData): GoogleAppsScript.Card_Service.Card {
    const card = CardService.newCardBuilder().setHeader(headerIcon({ title: "Periodos y Calendario", iconName: Icon.CALENDAR }));

    return card.addSection(buildPeriodSection(persistentData.currentPeriod)).build();
}

/**
 * Inputs sent to change current period
 */
export const SwitchPeriodInputs = defineInputsSchema({
    period: InputType.NUMBER,
    reset: InputType.BOOLEAN,
} as const);

/**
 * Select period section
 */
function buildPeriodSection(currentPeriod: Period): GoogleAppsScript.Card_Service.CardSection {
    const section = CardService.newCardSection().setHeader("📅 &nbsp; Periodo Actual");

    const periodDropdown = CardService.newSelectionInput()
        .setType(CardService.SelectionInputType.DROPDOWN)
        .setTitle("Selecciona el periodo")
        .setFieldName(SwitchPeriodInputs.fieldName("period"))
        .addItem("1er Periodo", "0", currentPeriod === Period.FIRST)
        .addItem("2do Periodo", "1", currentPeriod === Period.SECOND)
        .addItem("3er Periodo", "2", currentPeriod === Period.THIRD);

    const resetCheckbox = CardService.newSelectionInput()
        .setType(CardService.SelectionInputType.CHECK_BOX)
        .setFieldName(SwitchPeriodInputs.fieldName("reset"))
        .addItem("🧹 Reiniciar habilidades y comentarios", "true", false);

    const resetWarning = coloredParagraph(Colors.ORANGE, "⚠️ &nbsp; Si marcas esta casilla, TODAS las habilidades y comentarios de los alumnos se perderán.");

    const applyAction = CardService.newAction()
        .setFunctionName("") // TODO: Create callback
        .addRequiredWidget(SwitchPeriodInputs.fieldName("period"));

    const applyButton = CardService.newTextButton().setText("✅ - Cambiar periodo").setTextButtonStyle(CardService.TextButtonStyle.FILLED).setOnClickAction(applyAction);

    return section
        .addWidget(periodDropdown)
        .addWidget(resetCheckbox)
        .addWidget(resetWarning)
        .addWidget(spaceBetweenWidgets())
        .addWidget(CardService.newButtonSet().addButton(applyButton));
}
