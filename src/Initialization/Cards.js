const InitializationCards = {
    /**
     * Presents the user with a form to fill and a button to create the initialization file.
     * 
     * @param {string} folderId The Drive folder where the Initialization file will be created.
     * @returns {GoogleAppsScript.Card_Service.Card}
     */
    buildCreateInitializationFileCard(folderId) {
        const card = CardService.newCardBuilder()
            .setHeader(CardParts.headerImage({ title: "Registro Inicial de Grupos", subtitle: "Montessory Chacala", image: "school" }));

        const section = CardService.newCardSection().setHeader("Información del Grupo:");

        section.addWidget(CardService.newTextInput()
            .setValue("Secundaria")
            .setFieldName("groupName")
            .setTitle("Nombre del Grupo")
            .setHint("Ejemplo: 5to y 6to"));

        section.addWidget(CardService.newDecoratedText()
            .setText("Asistencia individual por materia")
            .setSwitchControl(CardService.newSwitch()
                .setFieldName("attendancePerClass")
                .setValue("attendancePerClass")
                .setSelected(false)));

        section.addWidget(CardService.newDecoratedText()
            .setText("Promedios por Campo Formativo")
            .setSwitchControl(CardService.newSwitch()
                .setFieldName("averagePerField")
                .setValue("averagePerField")
                .setSelected(false)));

        section.addWidget(CardService.newDatePicker()
            .setValueInMsSinceEpoch(1787961600000)
            .setFieldName("dateStart")
            .setTitle("Primer dia de clases"));

        section.addWidget(CardService.newDatePicker()
            .setValueInMsSinceEpoch(1793491200000)
            .setFieldName("dateEndTrimester1")
            .setTitle("Último día del primer trimestre"));

        section.addWidget(CardService.newDatePicker()
            .setValueInMsSinceEpoch(1797292800000)
            .setFieldName("dateEndTrimester2")
            .setTitle("Último día del segundo trimestre"));

        section.addWidget(CardService.newDatePicker()
            .setValueInMsSinceEpoch(1812326400000)
            .setFieldName("dateEnd")
            .setTitle("Último día de clases"));

        const createAction = CardService.newAction()
            .setFunctionName(onCreateInitializationFile.name)
            .setParameters({ destinationFolder: folderId })
            .addRequiredWidget("groupName")
            .addRequiredWidget("dateStart")
            .addRequiredWidget("dateEndTrimester1")
            .addRequiredWidget("dateEndTrimester2")
            .addRequiredWidget("dateEnd");

        section.addWidget(CardService.newTextButton()
            .setText("📋 Crear Registro Inicial del Grupo")
            .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
            .setOnClickAction(createAction));

        return card.addSection(section).build();
    },

    /**
     * 
     * @param {string} fileId 
     */
    buildInitializationFileEditCard(fileId) {
        const card = CardService.newCardBuilder()
            .setHeader(CardParts.headerIcon({ title: "Registro Inicial de Grupos", subtitle: "Montessory Chacala", iconName: Icon.clipboard }));

        const section = CardService.newCardSection();

        const calendarAction = CardService.newAction()
            .setFunctionName(onGenerateCalendar.name)
            .setParameters({ fileId });

        section.addWidget(CardService.newTextButton()
            .setText("Regenerar Calendario")
            .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
            .setOnClickAction(calendarAction));

        return card.addSection(section).build();
    }
}
