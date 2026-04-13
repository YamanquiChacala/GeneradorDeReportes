import { headerImage, icon } from "./cardParts";
import { Icon } from "./enums";

/**
 * A card indicating the user has selected an invalid file or folder.
 * Reminds them of the specific files/folders the add-on can manage.
 *
 * @returns {GoogleAppsScript.Card_Service.Card}
 */

export function buildWrongSelectionCard(): GoogleAppsScript.Card_Service.Card {
    const warning = CardService.newDecoratedText()
        .setStartIcon(icon({ iconName: Icon.FOLDER_QUESTION, color: "orange", height: 64 }))
        .setText("Selección no válida");

    const explanation = CardService.newTextParagraph().setText(
        "El archivo o carpeta seleccionado no es compatible.<br/><br/>Por favor, selecciona una de las siguientes opciones para continuar:<br/><br/>• 📁 Una <b>carpeta vacía</b> (para crear la configuración inicial)<br/><br/>• 📋 Un archivo de <b>Registro inicial de grupos</b><br/><br/>• 📊 Un archivo de <b>Calificaciones, asistencias y reportes</b>",
    );

    const mainSection = CardService.newCardSection().addWidget(warning).addWidget(explanation);

    return CardService.newCardBuilder()
        .setHeader(headerImage({ title: "Montessori Chacala", subtitle: "Archivo no reconocido" }))
        .addSection(mainSection)
        .build();
}
