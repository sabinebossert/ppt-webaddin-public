Office.onReady().then(() => {
    initTabs();

    // Feature initialisieren
    RoundedCorners.init();
});

function initTabs() {
    const buttons = document.querySelectorAll(".tab-button");
    const panels = document.querySelectorAll(".tab-panel");

    buttons.forEach(btn => {
        btn.addEventListener("click", () => {
            const target = btn.dataset.target;

            buttons.forEach(b => b.classList.remove("active"));
            btn.classList.add("active");

            panels.forEach(p => {
                p.classList.toggle("active", p.id === target);
            });
        });
    });
}
