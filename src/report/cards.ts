import { coloredParagraph, decoratedTextWithCallback, headerIcon, iconImage, spaceBetweenWidgets } from "../common/gas-parts";
import { Colors, checkboxGroup, defineActionParameters, defineInputsSchema, InputType, ParamType } from "../common/gas-utils";
import { Period, type ProtectedSections, type ReportPersistentData, StudentRowType } from "../common/report-utils";
import { Icon } from "../common/utils";
import { onChangeProtectedSections, onGenerateAllReports, onGenerateIndividualReport, onPushMenuCard, onSwitchPeriod } from "./callbacks";

export const mainMenuCards = {
    reports: buildReportCreateReportsCard,
    periods: buildReportsPeriodsAndCalendarCard,
};

export type mainMenuCardName = keyof typeof mainMenuCards;

export const NavParam = defineActionParameters({ target: ParamType.STRING } as const);

/**
 * Main Card for a report admin interface
 */
export function buildReportOptionsMainCard(): GoogleAppsScript.Card_Service.Card {
    const navTo = (target: mainMenuCardName) => NavParam.build({ target });

    const card = CardService.newCardBuilder().setHeader(headerIcon({ title: "Asistencia y Reportes", subtitle: "Montessori Chacala", iconName: Icon.CHART }));

    const cardSection = CardService.newCardSection()
        .setHeader("Tablero de Administración")
        .addWidget(spaceBetweenWidgets())
        .addWidget(
            decoratedTextWithCallback({
                callback: onPushMenuCard.name,
                parameters: navTo("reports"),
                text: "Reportes",
                bottomText: "Generar reportes individuales o por grupo",
                startIcon: iconImage({ iconName: Icon.PAGE }),
                endIcon: iconImage({ iconName: Icon.FORWARD }),
            }),
        )
        .addWidget(spaceBetweenWidgets())
        .addWidget(
            decoratedTextWithCallback({
                callback: onPushMenuCard.name,
                parameters: navTo("periods"),
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
export function buildReportsPeriodsAndCalendarCard(persistentData: ReportPersistentData): GoogleAppsScript.Card_Service.Card {
    const card = CardService.newCardBuilder().setHeader(headerIcon({ title: "Periodos y Calendario", iconName: Icon.CALENDAR }));

    return card
        .addSection(buildPeriodSection(persistentData.currentPeriod))
        .addSection(buildProtectedSection(persistentData.protectedSections))
        .addSection(buildAddRemoveDatesSection())
        .addSection(buildPeriodDatesSection(persistentData.configData.dates))
        .build();
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

    const applyAction = CardService.newAction().setFunctionName(onSwitchPeriod.name).addRequiredWidget(SwitchPeriodInputs.fieldName("period"));

    const applyButton = CardService.newTextButton().setText("✅ - Cambiar periodo").setTextButtonStyle(CardService.TextButtonStyle.FILLED).setOnClickAction(applyAction);

    return section
        .addWidget(periodDropdown)
        .addWidget(resetCheckbox)
        .addWidget(resetWarning)
        .addWidget(spaceBetweenWidgets())
        .addWidget(CardService.newButtonSet().addButton(applyButton));
}

export const ProtectSectionsInputs = defineInputsSchema({
    myNum: InputType.NUMBER,
    protectedSections: checkboxGroup(["abilities", "comments", "period1", "period2", "period3"]),
} as const);

/**
 * Select protected sections
 */
function buildProtectedSection(protectedSections: ProtectedSections): GoogleAppsScript.Card_Service.CardSection {
    const section = CardService.newCardSection().setHeader("🗝️ &nbsp; Secciones Protegidas");

    const explanation = CardService.newTextParagraph().setText("🔒 Marca las casillas para <b>bloquear</b> la edición a los maestros.");

    const protectionCheckboxes = CardService.newSelectionInput()
        .setType(CardService.SelectionInputType.CHECK_BOX)
        .setTitle("Secciones a bloquear")
        .setFieldName(ProtectSectionsInputs.fieldName("protectedSections"))
        .addItem("🔒 Habilidades", ProtectSectionsInputs.option("protectedSections", "abilities"), protectedSections.abilities)
        .addItem("🔒 Comentarios", ProtectSectionsInputs.option("protectedSections", "comments"), protectedSections.comments)
        .addItem("🔒 Calificaciones del 1er Periodo", ProtectSectionsInputs.option("protectedSections", "period1"), protectedSections.trimesters[0])
        .addItem("🔒 Calificaciones del 2do Periodo", ProtectSectionsInputs.option("protectedSections", "period2"), protectedSections.trimesters[1])
        .addItem("🔒 Calificaciones del 3er Periodo", ProtectSectionsInputs.option("protectedSections", "period3"), protectedSections.trimesters[2]);

    const applyAction = CardService.newAction().setFunctionName(onChangeProtectedSections.name);

    const applyButton = CardService.newTextButton()
        .setText("🔐 - Aplicar protección")
        .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
        .setOnClickAction(applyAction);

    return section.addWidget(explanation).addWidget(protectionCheckboxes).addWidget(spaceBetweenWidgets()).addWidget(CardService.newButtonSet().addButton(applyButton));
}

export const AddRemoveDatesInputs = defineInputsSchema({
    rangeStart: InputType.DATE,
    rangeEnd: InputType.DATE,
} as const);

/**
 * Add or remove ranges of days
 */
function buildAddRemoveDatesSection(): GoogleAppsScript.Card_Service.CardSection {
    const section = CardService.newCardSection().setHeader("🗓️ &nbsp; Agregar / Quitar Fechas").setCollapsible(true).setNumUncollapsibleWidgets(0);

    const startPicker = CardService.newDatePicker().setFieldName(AddRemoveDatesInputs.fieldName("rangeStart")).setTitle("Fecha inicial");
    const endPicker = CardService.newDatePicker().setFieldName(AddRemoveDatesInputs.fieldName("rangeEnd")).setTitle("Fecha final (opcional)");

    const hint = CardService.newTextParagraph().setText("Si no eliges fecha final, solo se agregará o quitará el día inicial.");
    const warning = coloredParagraph(Colors.ORANGE, "⚠️ &nbsp; Las asistencias registradas en las fechas que se quiten se perderán para siempre.");

    const addAction = CardService.newAction()
        .setFunctionName("") // TODO: Create callback
        .addRequiredWidget(AddRemoveDatesInputs.fieldName("rangeStart"));

    const addButton = CardService.newTextButton().setText("➕ - Agregar").setTextButtonStyle(CardService.TextButtonStyle.FILLED).setOnClickAction(addAction);

    const removeAction = CardService.newAction()
        .setFunctionName("") // TODO: Create callback
        .addRequiredWidget(AddRemoveDatesInputs.fieldName("rangeStart"));

    const removeButton = CardService.newTextButton()
        .setText("🗑️ - Quitar")
        .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
        .setBackgroundColor(Colors.ORANGE)
        .setOnClickAction(removeAction);

    return section
        .addWidget(startPicker)
        .addWidget(endPicker)
        .addWidget(hint)
        .addWidget(warning)
        .addWidget(CardService.newButtonSet().addButton(addButton).addButton(removeButton));
}

export const ChangePeriodDatesInputs = defineInputsSchema({
    start: InputType.DATE,
    period1: InputType.DATE,
    period2: InputType.DATE,
    end: InputType.DATE,
} as const);

/**
 * Modify the period dates
 */
function buildPeriodDatesSection(dates: [number, number, number, number]): GoogleAppsScript.Card_Service.CardSection {
    const section = CardService.newCardSection().setHeader("🏫 &nbsp; Fechas del Ciclo Escolar").setCollapsible(true).setNumUncollapsibleWidgets(0);

    const picker = (fieldName: string, title: string, value?: number) => {
        const datePicker = CardService.newDatePicker().setFieldName(fieldName).setTitle(title);
        if (value !== undefined) datePicker.setValueInMsSinceEpoch(value);
        return datePicker;
    };

    const updateAction = CardService.newAction()
        .setFunctionName("") // TODO: Create callback
        .addRequiredWidget(ChangePeriodDatesInputs.fieldName("start"))
        .addRequiredWidget(ChangePeriodDatesInputs.fieldName("period1"))
        .addRequiredWidget(ChangePeriodDatesInputs.fieldName("period2"))
        .addRequiredWidget(ChangePeriodDatesInputs.fieldName("end"));

    const updateButton = CardService.newTextButton()
        .setText("💾 - Actualizar fechas")
        .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
        .setOnClickAction(updateAction);

    return section
        .addWidget(picker(ChangePeriodDatesInputs.fieldName("start"), "Inicio del ciclo escolar", dates[0]))
        .addWidget(picker(ChangePeriodDatesInputs.fieldName("period1"), "Fin del 1er periodo", dates[1]))
        .addWidget(picker(ChangePeriodDatesInputs.fieldName("period2"), "Fin del 2do periodo", dates[2]))
        .addWidget(picker(ChangePeriodDatesInputs.fieldName("end"), "Fin del ciclo escolar", dates[3]))
        .addWidget(spaceBetweenWidgets())
        .addWidget(CardService.newButtonSet().addButton(updateButton));
}
