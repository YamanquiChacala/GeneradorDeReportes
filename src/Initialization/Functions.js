/**
 * Callback to the button to create a new Initialization Group File.
 * @param {GoogleAppsScript.Addons.EventObject} e 
 * @returns {GoogleAppsScript.Card_Service.ActionResponse}
 */
function onCreateInitializationFile(e) {
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
        return CardService.newActionResponseBuilder()
            .setNotification(CardService.newNotification()
                .setText("❌ Las fechas deben estar en orden ascendente."))
            .build();
    }

    const moreThanAYear = 400 * 24 * 60 * 60 * 1000;
    if (dateEnd - dateStart > moreThanAYear) {
        return CardService.newActionResponseBuilder()
            .setNotification(CardService.newNotification()
                .setText("❌ Periodo demasiado largo."))
            .build();
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
    }

    // TODO: Surrond by try-catch, show a card with the error if error.
    Initialization.createInitializationFile(initData);


    // Trigger creation example
    // const trigger = ScriptApp.newTrigger(fireCreateInitializationFile.name)
    //     .timeBased()
    //     .after(1)
    //     .create();

    // const triggerId = trigger.getUniqueId();
    // PropertiesService.getUserProperties().setProperty(triggerId, JSON.stringify(initData));

    // TODO: Extract this into a generic function buildParagraphCard(header, htmlText)
    const successCard = CardService.newCardBuilder()
        .setHeader(CardParts.headerImage({ title: "Registro Inicial de Grupos", subtitle: "Montessory Chacala", image: "school" }))
        .addSection(CardService.newCardSection()
            .addWidget(CardService.newTextParagraph()
                .setText(`✅ Creación del Registro para "<b>${groupName}</b>" en proceso.<br><br>El archivo aparecerá la carpeta de Drive en un momento.<br><br>(Por favor espera al menos 1 minúto antes de intentarlo de nuevo.)`))
            .addWidget(CardService.newTextButton()
                .setText("Regresar al inicio")
                .setOnClickAction(CardService.newAction().setFunctionName(onPopCardStack.name))))
        .build();

    return CardService.newActionResponseBuilder()
        .setNavigation(CardService.newNavigation().pushCard(successCard))
        .build();
}

// TODO: Move to file with generic reusable functions.
/**
 * Pops the Card Stack to the root.
 * @returns {GoogleAppsScript.Card_Service.ActionResponse}
 */
function onPopCardStack() {
    return CardService.newActionResponseBuilder().setNavigation(CardService.newNavigation().popToRoot()).build();
}


// Trigger callback example.
// /**
//  * Trigger callback to create a new Initialization Group File.
//  * @param {GoogleAppsScript.Events.TimeDriven} e 
//  */
// function fireCreateInitializationFile(e) {
//     const triggerId = e.triggerUid;
//     if (!triggerId) return;

//     const userProperties = PropertiesService.getUserProperties();
//     const payloadString = userProperties.getProperty(triggerId);

//     userProperties.deleteProperty(triggerId);
//     ScriptApp.getProjectTriggers().forEach(trigger => {
//         if (trigger.getUniqueId() === triggerId) {
//             ScriptApp.deleteTrigger(trigger);
//         }
//     });

//     if (!payloadString) return;
//     /** @type {InitFileData} */
//     const initData = JSON.parse(payloadString);

//     try {
//         Initialization.createInitializationFile(initData);
//     } catch (error) {
//         const errorMessage = error instanceof Error ? error.message : String(error);
//         console.error(`Error creando Registro para "${initData.groupName}": ${errorMessage}`);
//     }
// }

