/**
 * Callback to the button to create a new Initialization Group File.
 */
export function onCreateInitializationFile(e: GoogleAppsScript.Addons.EventObject): GoogleAppsScript.Card_Service.ActionResponse {
    const formInput = e.commonEventObject.formInputs;
    const folderId = e.commonEventObject.parameters.destinationFolder;

    const groupName = Utils.sanitizeFileName(formInput.groupName.stringInputs?.value[0]);

    const attendancePerClass = formInput.attendancePerClass ? true : false;
    const averagePerField = formInput.averagePerField ? true : false;

    const dateStart = parseInt(formInput.dateStart.dateInput?.msSinceEpoch ?? "");
    const dateEndTrimester1 = parseInt(formInput.dateEndTrimester1.dateInput?.msSinceEpoch ?? "");
    const dateEndTrimester2 = parseInt(formInput.dateEndTrimester2.dateInput?.msSinceEpoch ?? "");
    const dateEnd = parseInt(formInput.dateEnd.dateInput?.msSinceEpoch ?? "");

    if (!(dateStart < dateEndTrimester1 && dateEndTrimester1 < dateEndTrimester2 && dateEndTrimester2 < dateEnd)) {
        return CardService.newActionResponseBuilder().setNotification(CardService.newNotification().setText("❌ Las fechas deben estar en orden ascendente.")).build();
    }

    const moreThanAYear = 400 * 24 * 60 * 60 * 1000;
    if (dateEnd - dateStart > moreThanAYear) {
        return CardService.newActionResponseBuilder().setNotification(CardService.newNotification().setText("❌ Periodo demasiado largo.")).build();
    }

    /** @type {InitFileData} */
    const initData = {
        folderId,
        groupName,
        attendancePerClass,
        averagePerField,
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
        .setHeader(CardParts.headerImage({ title: "Registro Inicial de Grupos", subtitle: "Montessory Chacala", image: "school" }))
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
