// core/utils.js

let messageBanner;

/**
 * Initialisiert das Office UI Fabric MessageBanner,
 * das für Benachrichtigungen im Taskpane verwendet wird.
 */
function initMessageBanner() {
    const element = document.querySelector(".MessageBanner");

    // MessageBanner-Komponente kommt aus MessageBanner.js
    if (element && window.components && components.MessageBanner) {
        messageBanner = new components.MessageBanner(element);
        messageBanner.hideBanner();
    }
}

/**
 * Zeigt eine Benachrichtigung im Banner am unteren Rand des Taskpane.
 * @param {string} header  Überschrift der Meldung
 * @param {string} content Text der Meldung
 */
function showNotification(header, content) {
    const headerEl = document.getElementById("notification-header");
    const bodyEl = document.getElementById("notification-body");

    if (!headerEl || !bodyEl || !messageBanner) {
        console.warn("Notification could not be shown, MessageBanner not initialized.");
        return;
    }

    headerEl.textContent = header;
    bodyEl.textContent = content;

    messageBanner.showBanner();
    messageBanner.toggleExpansion();
}

/**
 * Schreibt einen Fehler in die Konsole und zeigt eine Benachrichtigung.
 * @param {string} source  Kontext/Quelle des Fehlers (z.B. "Apply border radius")
 * @param {any}    error   Error-Objekt oder Nachricht
 */
function logError(source, error) {
    console.error(source, error);

    const msg = error && error.message ? error.message : String(error);
    showNotification("Error", `${source}: ${msg}`);
}