var Evento = window.Evento || {};

(function () {


    this.OnLoad = function (executionContext) {
        debugger;
        formContext = executionContext.getFormContext();
        formType = formContext.ui.getFormType();

        if (formType == CrmFormType.update) {

        }


        if (formType == CrmFormType.create) {

        }

    }

    this.OnSave = function (executionContext) {
        debugger;
        formContext = executionContext.getFormContext();

        var chosenDate = formContext.getAttribute("aicgusto_data").getValue();

        var currentDate = new Date();

        if (chosenDate < currentDate) {

            formContext.ui.setFormNotification("Chosen date cannot be in the past.", "ERROR");

            // Prevent the form from saving
            executionContext.getEventArgs().preventDefault();

        }
    }


}).call(Evento);