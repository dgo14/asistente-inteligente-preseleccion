document.addEventListener("DOMContentLoaded", () => {
    const fileInput = document.querySelector("#cvs");
    const fileLabel = document.querySelector("#file-label");

    if (fileInput && fileLabel) {
        fileInput.addEventListener("change", () => {
            const count = fileInput.files.length;
            if (count === 0) {
                fileLabel.textContent = "Selecciona tus CVs en PDF";
            } else if (count === 1) {
                fileLabel.textContent = "1 archivo seleccionado";
            } else {
                fileLabel.textContent = `${count} archivos seleccionados`;
            }
        });
    }
});
