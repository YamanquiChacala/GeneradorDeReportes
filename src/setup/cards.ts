import { Colors, headerIcon, headerImage, textButton } from "../common/gas-parts";
import { buildFieldsMask, defineActionParameters, defineInputsSchema } from "../common/gas-utils";
import { Icon } from "../common/utils";
import { onCopySetupFile, onCreateSetupFile, onGenerateCalendar, onInitializeReport } from "./callbacks";

export const CreateSetupFileInputs = defineInputsSchema({
    groupName: "string",
    attendancePerClass: "boolean",
    averagePerField: "boolean",
    dateStart: "date",
    dateEndTrimester1: "date",
    dateEndTrimester2: "date",
    dateEnd: "date",
} as const);

export const CreateSetupFileParams = defineActionParameters({
    folderId: "string",
} as const);

/**
 * Presents the user with a form to fill and a button to create the initialization file.
 */
export function buildCreateSetupFileCard(folderId: string): GoogleAppsScript.Card_Service.Card {
    const card = CardService.newCardBuilder().setHeader(headerImage({ title: "Registro Inicial de Grupos", subtitle: "Montessori Chacala" }));

    const generalSection = CardService.newCardSection().setHeader("Configuración General:");

    generalSection.addWidget(
        CardService.newTextInput()
            .setValue("Secundaria")
            .setFieldName(CreateSetupFileInputs.fieldName("groupName"))
            .setTitle("Nombre del Grupo")
            .setHint("Ejemplo: 5to y 6to"),
    );

    generalSection.addWidget(
        CardService.newDecoratedText()
            .setText("Asistencia individual por materia")
            .setSwitchControl(CardService.newSwitch().setFieldName(CreateSetupFileInputs.fieldName("attendancePerClass")).setValue("true").setSelected(false)),
    );

    generalSection.addWidget(
        CardService.newDecoratedText()
            .setText("Promedios por Campo Formativo")
            .setSwitchControl(CardService.newSwitch().setFieldName(CreateSetupFileInputs.fieldName("averagePerField")).setValue("true").setSelected(false)),
    );

    const datesSection = CardService.newCardSection().setHeader("Calendario Escolar");

    datesSection.addWidget(
        CardService.newDatePicker().setValueInMsSinceEpoch(1787961600000).setFieldName(CreateSetupFileInputs.fieldName("dateStart")).setTitle("Primer dia de clases"),
    );

    datesSection.addWidget(
        CardService.newDatePicker()
            .setValueInMsSinceEpoch(1793491200000)
            .setFieldName(CreateSetupFileInputs.fieldName("dateEndTrimester1"))
            .setTitle("Último día del primer trimestre"),
    );

    datesSection.addWidget(
        CardService.newDatePicker()
            .setValueInMsSinceEpoch(1797292800000)
            .setFieldName(CreateSetupFileInputs.fieldName("dateEndTrimester2"))
            .setTitle("Último día del segundo trimestre"),
    );

    datesSection.addWidget(
        CardService.newDatePicker().setValueInMsSinceEpoch(1812326400000).setFieldName(CreateSetupFileInputs.fieldName("dateEnd")).setTitle("Último día de clases"),
    );

    const createAction = CardService.newAction()
        .setFunctionName(onCreateSetupFile.name)
        .setParameters(CreateSetupFileParams.build({ folderId }))
        .addRequiredWidget(CreateSetupFileInputs.fieldName("groupName"))
        .addRequiredWidget(CreateSetupFileInputs.fieldName("dateStart"))
        .addRequiredWidget(CreateSetupFileInputs.fieldName("dateEndTrimester1"))
        .addRequiredWidget(CreateSetupFileInputs.fieldName("dateEndTrimester2"))
        .addRequiredWidget(CreateSetupFileInputs.fieldName("dateEnd"));

    const submitButton = textButton({ text: "📋 Crear Registro Inicial del Grupo", action: createAction, style: CardService.TextButtonStyle.FILLED });

    const footer = CardService.newFixedFooter().setPrimaryButton(submitButton);

    return card.addSection(generalSection).addSection(datesSection).setFixedFooter(footer).build();
}

export const GenerateCalendarParams = defineActionParameters({
    setupFileId: "string",
} as const);

export const CopySetupFileInputs = defineInputsSchema({
    groupName: "string",
    folderId: "string",
} as const);

export const CopySetupFileParams = defineActionParameters({
    setupFileId: "string",
} as const);

export const InitializeReportParams = defineActionParameters({
    setupFileId: "string",
    parentId: "string",
} as const);

/**
 * Builds the Sheets card to show when the Setup file is open.
 */
export function buildEditSetupFileCard(setupFileId: string): GoogleAppsScript.Card_Service.Card {
    const card = CardService.newCardBuilder().setHeader(headerIcon({ title: "Registro Inicial de Grupos", subtitle: "Montessori Chacala", iconName: Icon.CLIPBOARD }));

    // --- Step 1: Manage the Calendar ---
    const calendarSection = CardService.newCardSection().setHeader("📅 1. Actualizar Calendario");

    calendarSection.addWidget(
        CardService.newTextParagraph().setText(
            "Si modificaste las fechas del ciclo escolar, actualiza el calendario aquí.<br><br>" +
                `<font color='${Colors.ORANGE}'><b>⚠️ Advertencia:</b></font> Regenerar el calendario borrará los días festivos que ya hayas seleccionado manualmente.`,
        ),
    );

    const calendarAction = CardService.newAction().setFunctionName(onGenerateCalendar.name).setParameters(GenerateCalendarParams.build({ setupFileId }));
    calendarSection.addWidget(
        CardService.newTextButton().setText("Regenerar Calendario").setTextButtonStyle(CardService.TextButtonStyle.OUTLINED).setOnClickAction(calendarAction),
    );

    // --- Step 2: Re-use / Copy ---
    const copySection = CardService.newCardSection().setHeader("📄 2. Duplicar Configuración");

    copySection.addWidget(CardService.newTextParagraph().setText("¿Tienes otro grupo con el mismo calendario? Haz una copia de este registro para ahorrar tiempo."));

    copySection.addWidget(
        CardService.newTextInput().setFieldName(CopySetupFileInputs.fieldName("groupName")).setTitle("Nombre del nuevo grupo").setHint("Ejemplo: 5to y 6to"),
    );

    const folderDropDown = CardService.newSelectionInput()
        .setType(CardService.SelectionInputType.DROPDOWN)
        .setTitle("Ubicación de la copia")
        .setFieldName(CopySetupFileInputs.fieldName("folderId"))
        .addItem("", "", false);

    let parentId: string;

    try {
        const fileData = Drive?.Files.get(setupFileId, { fields: buildFieldsMask<GoogleAppsScript.Drive_v3.Drive.V3.Schema.File>("parents"), supportsAllDrives: true });
        if (!fileData?.parents?.length || !fileData.parents[0]) throw new Error("No parent for the Setup file");
        parentId = fileData.parents[0];
        const parentData = Drive?.Files.get(fileData.parents[0], {
            fields: buildFieldsMask<GoogleAppsScript.Drive_v3.Drive.V3.Schema.File>("parents"),
            supportsAllDrives: true,
        });
        if (!parentData?.parents?.length || !parentData.parents[0]) throw new Error("No grandparent for the Setup file");
        const searchFolder = parentData.parents[0];

        const query = `'${searchFolder}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false`;
        const folderList = Drive?.Files.list({
            corpora: "allDrives",
            q: query,
            fields: "files(id,name)", // Can't use buildFieldMask, since Drive uses an old sintax that doesn't use dot sintax.
            orderBy: "name",
            supportsAllDrives: true,
            includeItemsFromAllDrives: true,
        });
        if (!folderList?.files) throw new Error("No list of folders.");
        folderList.files.forEach((folder) => {
            folderDropDown.addItem(`📁 ${folder.name}`, folder.id, false);
        });
    } catch (error) {
        const errorMessage = error instanceof Error ? error.message : String(error);
        parentId = "root";
        console.error(errorMessage);
        folderDropDown.addItem("📁 Carpeta actual", "null", false);
    }

    copySection.addWidget(folderDropDown);

    const copyAction = CardService.newAction()
        .setFunctionName(onCopySetupFile.name)
        .setParameters(CopySetupFileParams.build({ setupFileId }))
        .addRequiredWidget(CopySetupFileInputs.fieldName("groupName"))
        .addRequiredWidget(CopySetupFileInputs.fieldName("folderId"));

    copySection.addWidget(
        CardService.newTextButton().setText("Copiar Registro Inicial").setTextButtonStyle(CardService.TextButtonStyle.OUTLINED).setOnClickAction(copyAction),
    );

    // --- Step 3: Finalize (Context Only) ---
    const finalizeSection = CardService.newCardSection().setHeader("🚀 3. Finalizar y Empezar");

    // We adjust the text slightly to point the user toward the footer
    finalizeSection.addWidget(
        CardService.newTextParagraph().setText(
            "Si ya terminaste de revisar el calendario y los días festivos, utiliza el botón en la parte inferior para generar tu archivo principal de trabajo.",
        ),
    );

    // --- Fixed Footer (The Primary Action) ---
    const initializeAction = CardService.newAction().setFunctionName(onInitializeReport.name).setParameters(InitializeReportParams.build({ setupFileId, parentId }));
    const footerButton = CardService.newTextButton()
        .setText("📊 Crear Archivo de Calificaciones")
        .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
        .setOnClickAction(initializeAction);

    const footer = CardService.newFixedFooter().setPrimaryButton(footerButton);

    return card
        .addSection(calendarSection)
        .addSection(copySection)
        .addSection(finalizeSection)
        .setFixedFooter(footer) // Pins the main action to the bottom!
        .build();
}
