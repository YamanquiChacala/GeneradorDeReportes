import { onPopCardStack } from "../common/callbacks";
import { headerImage } from "../common/cardParts";
import { flattenFormInputs, sanitizeFileName } from "../common/utils";

interface InitFileData {
    folderId: string;
    groupName: string;
    attendancePerClass: boolean;
    averagePerField: boolean;
    dateStart: number;
    dateEndTrimester1: number;
    dateEndTrimester2: number;
    dateEnd: number;
}

/**
 * Callback to the button to create a new Initialization Group File.
 */
export function onCreateInitializationFile(e: GoogleAppsScript.Addons.EventObject): GoogleAppsScript.Card_Service.ActionResponse {
    const { destinationFolder: folderId } = e.commonEventObject.parameters;

    const { groupName, attendancePerClass, averagePerField, dateStart, dateEndTrimester1, dateEndTrimester2, dateEnd } = flattenFormInputs<InitFileData>(
        e.commonEventObject.formInputs,
    );

    if (!folderId) {
        return CardService.newActionResponseBuilder()
            .setNotification(CardService.newNotification().setText("❌ Error creando Registro de Grupo: Falta carpeta de destino."))
            .build();
    }
    if (!groupName || !dateStart || !dateEndTrimester1 || !dateEndTrimester2 || !dateEnd) {
        return CardService.newActionResponseBuilder().setNotification(CardService.newNotification().setText("❌ Error creando Registro de Grupo: Faltan datos.")).build();
    }

    const groupNameSanitized = sanitizeFileName(groupName);

    if (!(dateStart < dateEndTrimester1 && dateEndTrimester1 < dateEndTrimester2 && dateEndTrimester2 < dateEnd)) {
        return CardService.newActionResponseBuilder()
            .setNotification(CardService.newNotification().setText("❌ Error creando Registro de Grupo: Las fechas deben estar en orden ascendente."))
            .build();
    }

    const moreThanAYear = 400 * 24 * 60 * 60 * 1000;
    if (dateEnd - dateStart > moreThanAYear) {
        return CardService.newActionResponseBuilder()
            .setNotification(CardService.newNotification().setText("❌ Error creando Registro de Grupo: Periodo demasiado largo."))
            .build();
    }

    const initData: InitFileData = {
        folderId,
        groupName: groupNameSanitized,
        attendancePerClass: attendancePerClass ?? false,
        averagePerField: averagePerField ?? false,
        dateStart,
        dateEndTrimester1,
        dateEndTrimester2,
        dateEnd,
    };

    try {
        Initialization.createInitializationFile(initData);
    } catch (error) {
        const errorMessage = error instanceof Error ? error.message : String(error);
        console.error(errorMessage);
        return CardService.newActionResponseBuilder()
            .setNotification(CardService.newNotification().setText(`❌ Error creando Registro de Grupo: ${errorMessage}`))
            .build();
    }

    // Trigger creation example
    // const trigger = ScriptApp.newTrigger(fireCreateInitializationFile.name)
    //     .timeBased()
    //     .after(1)
    //     .create();

    // const triggerId = trigger.getUniqueId();
    // PropertiesService.getUserProperties().setProperty(triggerId, JSON.stringify(initData));

    const successCard = CardService.newCardBuilder()
        .setHeader(headerImage({ title: "Registro Inicial de Grupos", subtitle: "Montessory Chacala", image: "school" }))
        .addSection(
            CardService.newCardSection()
                .addWidget(
                    CardService.newTextParagraph().setText(
                        `✅ Creación del Registro para "<b>${groupName}</b>" en proceso.<br><br>El archivo aparecerá la carpeta de Drive en un momento.<br><br>(Por favor espera al menos 1 minúto antes de intentarlo de nuevo.)`,
                    ),
                )
                .addWidget(CardService.newTextButton().setText("Regresar al inicio").setOnClickAction(CardService.newAction().setFunctionName(onPopCardStack.name))),
        )
        .build();

    return CardService.newActionResponseBuilder().setNavigation(CardService.newNavigation().pushCard(successCard)).build();
}
