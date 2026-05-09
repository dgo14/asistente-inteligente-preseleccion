document.addEventListener("DOMContentLoaded", () => {
    const input = document.getElementById("cvs");
    const label = document.getElementById("file-label");

    if (input && label) {
        input.addEventListener("change", () => {
            const count = input.files.length;

            if (count === 0) {
                label.innerText = "Selecciona tus CVs en PDF";
            } else if (count === 1) {
                label.innerText = input.files[0].name;
            } else {
                label.innerText = `${count} archivos PDF seleccionados`;
            }
        });
    }

    const sidebarToggle = document.getElementById("sidebar-toggle");
    const dashboardLayout = document.getElementById("dashboard-layout");

    if (sidebarToggle && dashboardLayout) {
        sidebarToggle.addEventListener("click", () => {
            dashboardLayout.classList.toggle("sidebar-collapsed");
        });
    }
});

