
/**
 * A card asking the user to give authorization for the current file.
 * 
 * @returns {GoogleAppsScript.Card_Service.Card}
 */
function buildRequestAuthorizationCard() {
    const warning = CardService.newDecoratedText()
        .setStartIcon(CardParts.icon({ name: Icon.warning, color: "red", height: 64 }))
        .setText("Verifica antes de continuar");
    const explanation = CardService.newTextParagraph()
        .setText("Confirma que estás en la hoja correcta:<br/><br/>• 📋 <b>Registro inicial de grupos</b><br/><br/>• 📊 <b>Calificaciones, asistencias y reportes</b><br/><br/>🔒 Autorizar permitirá <b>editar</b> esta hoja.");
    const askPermissionAction = CardService.newAction().setFunctionName(onAskPermission.name);
    const askPermissionButton = CardService.newTextButton()
        .setText('🔑 Dar permiso')
        .setOnClickAction(askPermissionAction);

    const mainSection = CardService.newCardSection()
        .addWidget(warning)
        .addWidget(explanation)
        .addWidget(askPermissionButton);
    return CardService.newCardBuilder()
        .setHeader(CardParts.header({ title: "Montessori Chacala", subtitle: "Permiso de edición", icon: "school" }))
        .addSection(mainSection)
        .build();
}

/**
 * Callback to request authorization for the current file.
 * 
 * @param {GoogleAppsScript.Addons.EventObject} e 
 * @returns {GoogleAppsScript.Card_Service.EditorFileScopeActionResponse}
 */
function onAskPermission(e) {
    return CardService.newEditorFileScopeActionResponseBuilder()
        .requestFileScopeForActiveDocument()
        .build();
}