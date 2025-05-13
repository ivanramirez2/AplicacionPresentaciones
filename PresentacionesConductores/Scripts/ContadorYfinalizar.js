document.addEventListener("DOMContentLoaded", function () {
    console.log("Script ContadorYfinalizar.js cargado y DOM listo.");

    // 30 seg ..
    let timeLeft = 30;
    const timerElement = document.querySelector(".presentation-timer");

    if (timerElement) {
        console.log("Temporizador encontrado. Iniciando conteo...");
        function updateTimer() {
            const minutes = Math.floor(timeLeft / 60);
            const seconds = timeLeft % 60;
            timerElement.textContent = `${minutes}:${seconds < 10 ? "0" : ""}${seconds}`;

            if (timeLeft <= 0) {
                console.log("Temporizador finalizado. Redirigiendo a /Home/Index...");
                window.location.href = "/Home/Index";
                window.close();
            } else {
                timeLeft--;
                setTimeout(updateTimer, 1000);
            }
        }

        updateTimer();
    } else {
        console.log("Temporizador NO encontrado.");
    }

   
    const finalizeButton = document.querySelector(".presentation-finalize-button");
    if (finalizeButton) {
        console.log("Botón FINALIZAR encontrado. Asignando evento click...");
        finalizeButton.addEventListener("click", function () {
            console.log("Botón FINALIZAR clicado. Redirigiendo a /Home/Index...");
            window.location.href = "/Home/Index";
            window.close();
        });
    } else {
        console.log("Botón FINALIZAR NO encontrado.");
    }

    console.log("Esto está dentro de ContadorYFinalizar.js");

});