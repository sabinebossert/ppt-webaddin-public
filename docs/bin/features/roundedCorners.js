var RoundedCorners = (function () {

    function init() {
        initMessageBanner();

        document.getElementById("normalize-corners-button")
            .addEventListener("click", onNormalizeButtonClick);

        document.getElementById("undo-normalize-button")
            .addEventListener("click", onUndoButtonClick);

        document.getElementById("preset-radius-2")
            .addEventListener("click", () => applyPresetRadius(2));

        document.getElementById("preset-radius-4")
            .addEventListener("click", () => applyPresetRadius(4));

        document.getElementById("preset-radius-8")
            .addEventListener("click", () => applyPresetRadius(8));
    }

    function getSelectedScope() {
        const radios = document.querySelectorAll("input[name='scopeRadio']");
        for (const r of radios) if (r.checked) return r.value;
        return "selected";
    }

    function applyPresetRadius(radius) {
        const input = document.getElementById("radiusInput");
        input.value = radius;
        showNotification("Preset selected", `Radius = ${radius} pt`);
    }

    async function onNormalizeButtonClick() {
        const inputEl = document.getElementById("radiusInput");
        const val = inputEl ? parseFloat(inputEl.value) : NaN;

        if (isNaN(val) || val < 0) {
            showNotification(
                "Invalid value",
                "Please enter a non-negative numeric value for the border radius (in points)."
            );
            return;
        }

        const scope = getSelectedScope();

        // Optional: Hinweis wie im alten Code
        try {
            const docUrl = Office.context && Office.context.document && Office.context.document.url;
            if (!docUrl) {
                showNotification(
                    "Hint",
                    "The presentation does not seem to be saved yet. " +
                    "Please save the file before applying extensive changes."
                );
            }
        } catch (e) {
            console.warn("Could not read document URL.", e);
        }

        try {
            showNotification(
                "Working...",
                "Applying border radius. This may take a moment for large presentations."
            );

            const summary = await normalizeRectangleCorners(val, {
                applyChanges: true,
                scope
            });

            if (!summary || summary.aborted) {
                return;
            }

            showNotification(
                "Done",
                `Slides processed: ${summary.slidesProcessed}, ` +
                `shapes checked: ${summary.shapesProcessed}, ` +
                `shapes adjusted: ${summary.shapesModified}, ` +
                `shapes skipped: ${summary.shapesSkipped}.`
            );
        } catch (e) {
            logError("Apply border radius", e);
        }
    }

    async function onUndoButtonClick() {
        await undoLastNormalization();
    }

    return {
        init
    };
})();