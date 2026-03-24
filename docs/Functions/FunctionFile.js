// Die Initialisierungsfunktion muss bei jedem Laden einer neuen Seite ausgeführt werden.
Office.onReady(() => {
        // Wenn eine Initialisierung erfolgen muss, kann dies hier geschehen.
});

async function sampleFunction(event) {
try {
        await PowerPoint.run(async (context) => {
                const textRange = context.presentation.getSelectedTextRange();
                textRange.font.color = "green";
                await context.sync();
                });            
} catch (error) {
        console.error(error);
}

// Das Aufrufen von event.completed ist erforderlich. event.completed teilt der Plattform mit, dass die Verarbeitung abgeschlossen wurde.
event.completed();
}

Office.actions.associate("sampleFunction", sampleFunction);
