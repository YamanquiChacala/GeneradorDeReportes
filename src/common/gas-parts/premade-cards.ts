import { Colors, Icon } from "../constants";
import { onAskPermission } from "./callbacks";
import { headerImage, icon } from "./card-parts";

interface UtilityCardParams {
    header: GoogleAppsScript.Card_Service.CardHeader;
    title?: string;
    message?: string;
    points?: string[];
    button?: GoogleAppsScript.Card_Service.Button;
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
            CardService.newTextParagraph().setText(title),
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

/**
 * A card indicating the user has selected an invalid file or folder.
 * Reminds them of the specific files/folders the add-on can manage.
 */
export function buildWrongSelectionCard(): GoogleAppsScript.Card_Service.Card {
    // 1. Highlight the warning clearly
    const warning = CardService.newDecoratedText()
        .setStartIcon(icon({ iconName: Icon.FOLDER_QUESTION, color: Colors.ORANGE, height: 48 }))
        .setText(`<font color="${Colors.ORANGE}"><b>Selección no válida</b></font>`);

    const explanation = CardService.newTextParagraph().setText(
        "El elemento actual no es compatible. Para continuar, por favor abre o selecciona una de las siguientes opciones:",
    );

    const mainSection = CardService.newCardSection().addWidget(warning).addWidget(explanation);

    // 2. Use distinct widgets for the list with matching IconImages
    const optionsSection = CardService.newCardSection()
        .addWidget(
            CardService.newDecoratedText()
                .setStartIcon(icon({ iconName: Icon.FOLDER }))
                .setText("<b>Carpeta de Drive</b>")
                .setBottomLabel("Para crear la configuración inicial"),
        )
        .addWidget(
            CardService.newDecoratedText()
                .setStartIcon(icon({ iconName: Icon.CLIPBOARD }))
                .setText("<b>Registro inicial de grupos</b>")
                .setBottomLabel("Archivo de configuración base"),
        )
        .addWidget(
            CardService.newDecoratedText()
                .setStartIcon(icon({ iconName: Icon.CHART }))
                .setText("<b>Calificaciones, asistencias y reportes</b>")
                .setBottomLabel("Archivo principal de gestión")
                .setWrapText(true), // Ensures the longer title doesn't get cut off
        );

    // 3. Build and return the card
    return CardService.newCardBuilder()
        .setHeader(headerImage({ title: "Archivo no reconocido", subtitle: "Montessori Chacala" }))
        .addSection(mainSection)
        .addSection(optionsSection)
        .build();
}

/**
 * A card asking the user to give authorization for the current file.
 * Acts as a gatekeeper to prevent users from authorizing the wrong sheets.
 */
export function buildRequestAuthorizationCard(): GoogleAppsScript.Card_Service.Card {
    const card = CardService.newCardBuilder().setHeader(headerImage({ title: "Permiso de edición", subtitle: "Montessori Chacala" }));

    const section = CardService.newCardSection();

    // 1. Prominent Warning Header (Now using the Enum)
    section.addWidget(
        CardService.newDecoratedText()
            .setStartIcon(icon({ iconName: Icon.WARNING, color: Colors.ORANGE, height: 48 }))
            // Using template literals to inject the Enum color directly into the HTML
            .setText(`<font color='${Colors.ORANGE}'><b>Verifica antes de continuar</b></font>`),
    );

    // 2. Clear Context
    section.addWidget(
        CardService.newTextParagraph().setText("El complemento necesita permiso para interactuar con este archivo. Confirma que estás en la hoja correcta:"),
    );

    // 3. The Valid Files (Now using IconImages for consistent rendering)
    section.addWidget(
        CardService.newDecoratedText()
            .setStartIcon(icon({ iconName: Icon.CLIPBOARD }))
            .setText("<b>Registro inicial de grupos</b>"),
    );
    section.addWidget(
        CardService.newDecoratedText()
            .setStartIcon(icon({ iconName: Icon.CHART }))
            .setText("<b>Calificaciones, asistencias y reportes</b>")
            .setWrapText(true),
    );

    // 4. Consequence/Security message
    section.addWidget(CardService.newTextParagraph().setText("🔒 <i>Al autorizar, el complemento podrá <b>editar</b> este documento.</i>"));

    // 5. Action Button
    const askPermissionAction = CardService.newAction().setFunctionName(onAskPermission.name);
    const askPermissionButton = CardService.newTextButton()
        .setText("🔑 Dar permiso")
        .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
        .setOnClickAction(askPermissionAction);

    section.addWidget(CardService.newButtonSet().addButton(askPermissionButton));

    return card.addSection(section).build();
}
