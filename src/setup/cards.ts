import { headerIcon, headerImage, textButton } from "../common/cardParts";
import { Colors, Icon } from "../common/enums";
import { onCreateInitializationFile, onGenerateCalendar } from "./callbacks";

/**
 * Presents the user with a form to fill and a button to create the initialization file.
 */
export function buildCreateInitializationFileCard(folderId: string): GoogleAppsScript.Card_Service.Card {
    const card = CardService.newCardBuilder().setHeader(headerImage({ title: "Registro Inicial de Grupos", subtitle: "Montessori Chacala" }));

    const generalSection = CardService.newCardSection().setHeader("Configuración General:");

    generalSection.addWidget(CardService.newTextInput().setValue("Secundaria").setFieldName("groupName").setTitle("Nombre del Grupo").setHint("Ejemplo: 5to y 6to"));

    generalSection.addWidget(
        CardService.newDecoratedText()
            .setText("Asistencia individual por materia")
            .setSwitchControl(CardService.newSwitch().setFieldName("attendancePerClass").setValue("true").setSelected(false)),
    );

    generalSection.addWidget(
        CardService.newDecoratedText()
            .setText("Promedios por Campo Formativo")
            .setSwitchControl(CardService.newSwitch().setFieldName("averagePerField").setValue("true").setSelected(false)),
    );

    const datesSection = CardService.newCardSection().setHeader("Calendario Escolar");

    datesSection.addWidget(CardService.newDatePicker().setValueInMsSinceEpoch(1787961600000).setFieldName("dateStart").setTitle("Primer dia de clases"));

    datesSection.addWidget(
        CardService.newDatePicker().setValueInMsSinceEpoch(1793491200000).setFieldName("dateEndTrimester1").setTitle("Último día del primer trimestre"),
    );

    datesSection.addWidget(
        CardService.newDatePicker().setValueInMsSinceEpoch(1797292800000).setFieldName("dateEndTrimester2").setTitle("Último día del segundo trimestre"),
    );

    datesSection.addWidget(CardService.newDatePicker().setValueInMsSinceEpoch(1812326400000).setFieldName("dateEnd").setTitle("Último día de clases"));

    const createAction = CardService.newAction()
        .setFunctionName(onCreateInitializationFile.name)
        .setParameters({ folderId })
        .addRequiredWidget("groupName")
        .addRequiredWidget("dateStart")
        .addRequiredWidget("dateEndTrimester1")
        .addRequiredWidget("dateEndTrimester2")
        .addRequiredWidget("dateEnd");

    const submitButton = textButton({ text: "📋 Crear Registro Inicial del Grupo", action: createAction, style: CardService.TextButtonStyle.FILLED });

    const footer = CardService.newFixedFooter().setPrimaryButton(submitButton);

    return card.addSection(generalSection).addSection(datesSection).setFixedFooter(footer).build();
}

/**
 *
 */
export function buildInitializationFileEditCard(fileId: string): GoogleAppsScript.Card_Service.Card {
    const card = CardService.newCardBuilder().setHeader(headerIcon({ title: "Registro Inicial de Grupos", subtitle: "Montessori Chacala", iconName: Icon.CLIPBOARD }));

    // --- Step 1: Manage the Calendar ---
    const calendarSection = CardService.newCardSection().setHeader("📅 1. Actualizar Calendario");

    calendarSection.addWidget(
        CardService.newTextParagraph().setText(
            "Si modificaste las fechas del ciclo escolar, actualiza el calendario aquí.<br><br>" +
                `<font color='${Colors.ORANGE}'><b>⚠️ Advertencia:</b></font> Regenerar el calendario borrará los días festivos que ya hayas seleccionado manualmente.`,
        ),
    );

    const calendarAction = CardService.newAction().setFunctionName(onGenerateCalendar.name).setParameters({ fileId });
    calendarSection.addWidget(
        CardService.newTextButton().setText("Regenerar Calendario").setTextButtonStyle(CardService.TextButtonStyle.TEXT).setOnClickAction(calendarAction),
    );

    // --- Step 2: Re-use / Copy ---
    const copySection = CardService.newCardSection().setHeader("📄 2. Duplicar Configuración");

    copySection.addWidget(CardService.newTextParagraph().setText("¿Tienes otro grupo con el mismo calendario? Haz una copia de este registro para ahorrar tiempo."));

    copySection.addWidget(CardService.newTextInput().setFieldName("groupName").setTitle("Nombre del nuevo grupo").setHint("Ejemplo: 5to y 6to"));

    const copyAction = CardService.newAction().setFunctionName("onCopySetup").setParameters({ fileId }).addRequiredWidget("groupName");

    copySection.addWidget(
        CardService.newTextButton().setText("Copiar Registro Inicial").setTextButtonStyle(CardService.TextButtonStyle.TEXT).setOnClickAction(copyAction),
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
    const initializeAction = CardService.newAction().setFunctionName("onInitializeReport").setParameters({ fileId });
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
