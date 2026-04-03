/**
 * This function builds the homepage card for the Add-on.
 * It will run whenever the Add-on is opened in Drive or Sheets.
 */
function buildHelloWorldCard(e) {
    // 1. Create a text widget
    var textParagraph = CardService.newTextParagraph()
        .setText("Hello World! This is my first Workspace Add-on.");

    // 2. Create a section and add the widget to it
    var section = CardService.newCardSection()
        .addWidget(textParagraph);

    // 3. Build the card, add a header, and attach the section
    var card = CardService.newCardBuilder()
        .setHeader(CardService.newCardHeader().setTitle("Report Card Automator"))
        .addSection(section)
        .build();

    // 4. Return the card so Google Workspace knows what to display
    return card;
}