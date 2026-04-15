import { headerImage, icon } from "./cardParts";
import { Icon } from "./enums";

interface UtilityCardParams {
    header: GoogleAppsScript.Card_Service.CardHeader;
    title?: string;
    message?: string;
    points?: string[];
    button?: GoogleAppsScript.Card_Service.Button;
}

/**
 * A card indicating the user has selected an invalid file or folder.
 * Reminds them of the specific files/folders the add-on can manage.
 */
export function buildWrongSelectionCard(): GoogleAppsScript.Card_Service.Card {
    // 1. Highlight the warning clearly
    const warning = CardService.newDecoratedText()
        .setStartIcon(icon({ iconName: Icon.FOLDER_QUESTION, color: "#EA4335", height: 48 })) // Switched to a slightly smaller red icon for error state
        .setText("<font color='#EA4335'><b>Selección no válida</b></font>");

    const explanation = CardService.newTextParagraph().setText(
        "El elemento actual no es compatible. Para continuar, por favor abre o selecciona una de las siguientes opciones:",
    );

    const mainSection = CardService.newCardSection().addWidget(warning).addWidget(explanation);

    // 2. Use distinct widgets for the list. This creates a native, scannable layout.
    const optionsSection = CardService.newCardSection()
        .addWidget(CardService.newDecoratedText().setText("📁 <b>Carpeta de Drive</b>").setBottomLabel("Para crear la configuración inicial"))
        .addWidget(CardService.newDecoratedText().setText("📋 <b>Registro inicial de grupos</b>").setBottomLabel("Archivo de configuración base"))
        .addWidget(
            CardService.newDecoratedText().setText("📊 <b>Calificaciones, asistencias y reportes</b>").setBottomLabel("Archivo principal de gestión").setWrapText(true), // Ensures the longer title doesn't get cut off
        );

    // 3. Build and return the card
    return CardService.newCardBuilder()
        .setHeader(headerImage({ title: "Montessori Chacala", subtitle: "Archivo no reconocido" }))
        .addSection(mainSection)
        .addSection(optionsSection)
        .build();
}

/**
 * Builds a flexible utility card for statuses, errors, or simple info.
 * Dynamically adds widgets only if the data is provided in the config.
 */
export function buildUtilityCard({ header, title, message, points, button }: UtilityCardParams): GoogleAppsScript.Card_Service.Card {
    const section = CardService.newCardSection();

    // 1. Optional Title
    if (title) {
        section.addWidget(
            // Using a slightly larger, bold font to establish hierarchy
            CardService.newTextParagraph().setText(`<font size="large"><b>${title}</b></font>`),
        );
    }

    // 2. Optional Main Message
    if (message) {
        section.addWidget(CardService.newTextParagraph().setText(message));
    }

    // 3. Optional List of Points
    if (points && points.length > 0) {
        // Adding each point as its own widget avoids clunky <br/> tags
        // and gives the list natural, readable spacing.
        points.forEach((point) => {
            section.addWidget(CardService.newTextParagraph().setText(`• ${point}`));
        });
    }

    // 4. Optional Button
    if (button) {
        // Wrapping the button in a ButtonSet prevents layout stretching
        // and aligns it nicely according to standard Material Design rules.
        const buttonSet = CardService.newButtonSet().addButton(button);
        section.addWidget(buttonSet);
    }

    return CardService.newCardBuilder().setHeader(header).addSection(section).build();
}
