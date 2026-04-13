import * as CardParts from "./common/cardParts";
import { FileType, Icon } from "./common/enums";
import { getFileType } from "./common/fileValidation";
import { buildWrongSelectionCard } from "./common/premadeCards";

/**
 *
 */
export function buildSheetsCard(e: GoogleAppsScript.Addons.EventObject): GoogleAppsScript.Card_Service.Card {
    if (!e.sheets?.addonHasFileScopePermission) {
        return buildRequestAuthorizationCard();
    }

    const sheetId = e?.sheets?.id;

    if (!sheetId) return buildWrongSelectionCard();

    const sheetFileType = getFileType(sheetId);

    switch (sheetFileType) {
        case FileType.SETUP:
        //TODO: Activate this
        // return InitializationCards.buildInitializationFileEditCard(sheetId);
        case FileType.REPORT:
        // TODO: Return Card "Report options"
    }

    return buildWrongSelectionCard();
}

/**
 * A card asking the user to give authorization for the current file.
 */
function buildRequestAuthorizationCard(): GoogleAppsScript.Card_Service.Card {
    const warning = CardService.newDecoratedText()
        .setStartIcon(CardParts.icon({ iconName: Icon.WARNING, color: "red", height: 64 }))
        .setText("Verifica antes de continuar");
    const explanation = CardService.newTextParagraph().setText(
        "Confirma que estás en la hoja correcta:<br/><br/>• 📋 <b>Registro inicial de grupos</b><br/><br/>• 📊 <b>Calificaciones, asistencias y reportes</b><br/><br/>🔒 Autorizar permitirá <b>editar</b> esta hoja.",
    );
    const askPermissionAction = CardService.newAction().setFunctionName(onAskPermission.name);
    const askPermissionButton = CardService.newTextButton().setText("🔑 Dar permiso").setOnClickAction(askPermissionAction);

    const mainSection = CardService.newCardSection().addWidget(warning).addWidget(explanation).addWidget(askPermissionButton);
    return CardService.newCardBuilder()
        .setHeader(CardParts.headerImage({ title: "Montessori Chacala", subtitle: "Permiso de edición", image: "school" }))
        .addSection(mainSection)
        .build();
}

/**
 * Callback to request authorization for the current file.
 */
function onAskPermission(): GoogleAppsScript.Card_Service.EditorFileScopeActionResponse {
    return CardService.newEditorFileScopeActionResponseBuilder().requestFileScopeForActiveDocument().build();
}
