function onNewMessageComposeHandler(event) {
    var setting = Office.context.document.settings.get('openApiToken');
    if (!setting) {
        Office.context.ui.displayDialogAsync('https://localhost:3000/popups/tokenpopup.html', { height: 30, width: 20 },
            function (asyncResult) {
                const dialog = asyncResult.value;
                dialog.addEventHandler(Office.EventType.DialogMessageReceived, (args) => {
                    Office.context.document.settings.set('openApiToken', args);
                    dialog.close();
                });
            }
        );
    }
}