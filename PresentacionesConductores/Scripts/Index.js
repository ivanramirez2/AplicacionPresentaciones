document.addEventListener('DOMContentLoaded', function () {
    // Manteniendo cerrarPopup() en su lugar
    window.cerrarPopup = function () {
        console.log("✅ cerrarPopup() ha sido llamada.");
        const popup = document.getElementById("popup");
        if (popup) {
            popup.classList.remove("active");
            console.log("✅ Pop-up ocultado correctamente.");
        } else {
            console.error("❌ Error: No se encontró el pop-up en el DOM.");
        }
    };

    // Obtener los elementos necesarios del DOM
    const displayNumeroEmpleado = document.getElementById('displayNumeroEmpleado');
    const popup = document.getElementById('popup');
    const popupMessage = document.getElementById('popupMessage');

    if (!displayNumeroEmpleado) {
        console.error('Error: No se encontró el elemento displayNumeroEmpleado en el DOM.');
        return;
    }

    let currentNumeroEmpleado = displayNumeroEmpleado.textContent || '';

    displayNumeroEmpleado.textContent = currentNumeroEmpleado || " ";


    // Verificar si hay mensaje de error y activar el pop-up si es necesario
    const errorMensaje = popupMessage.textContent.trim();
    console.log("DEBUG ERROR: ", errorMensaje);

    if (errorMensaje && errorMensaje !== "") {
        console.log("✅ Activando pop-up después del submit...");
        popup.classList.add("active");

        setTimeout(() => {
            popup.classList.remove("active");
            console.log("✅ Pop-up ocultado automáticamente.");
        }, 6500);
    }

    function updateDateTime() {
        const now = new Date();
        const daysOfWeek = ["DOMINGO", "LUNES", "MARTES", "MIÉRCOLES", "JUEVES", "VIERNES", "SÁBADO"];
        const months = ["ENE", "FEB", "MAR", "ABR", "MAY", "JUN", "JUL", "AGO", "SEP", "OCT", "NOV", "DIC"];
        const dayName = daysOfWeek[now.getDay()];
        const day = now.getDate();
        const month = months[now.getMonth()];
        const hours = String(now.getHours()).padStart(2, "0");
        const minutes = String(now.getMinutes()).padStart(2, "0");
        const formattedDateTime = `${dayName} ${day} ${month} ${hours}:${minutes}`;
        const dateElement = document.querySelector(".date");
        if (dateElement) {
            dateElement.textContent = formattedDateTime;
        }
    }

    updateDateTime();
    setInterval(updateDateTime, 1000);

    // Configurar teclado numérico para número de empleado
    const numberButtons = ["btn0", "btn1", "btn2", "btn3", "btn4", "btn5", "btn6", "btn7", "btn8", "btn9"];

numberButtons.forEach(id => {
    const btnElement = document.getElementById(id);
    if (btnElement) {
        btnElement.addEventListener('click', function () {
            currentNumeroEmpleado += btnElement.textContent.trim(); 
            displayNumeroEmpleado.innerText = currentNumeroEmpleado; 
        });
    }
});



    const btnClear = document.getElementById('btnClear');
    if (btnClear) {
        btnClear.addEventListener('click', function () {
            currentNumeroEmpleado = '';
            displayNumeroEmpleado.textContent = '';
        });
    }

    const btnSubmit = document.getElementById('btnSubmit');
    if (btnSubmit) {
        btnSubmit.addEventListener('click', function () {
            const hiddenNumeroEmpleado = document.getElementById('numeroEmpleado');
            if (hiddenNumeroEmpleado) {
                if (!currentNumeroEmpleado.trim()) {
                    alert("Debe ingresar un número de empleado.");
                    return;
                }
                hiddenNumeroEmpleado.value = currentNumeroEmpleado;
                document.getElementById('dataForm').submit();
            }
        });
    }
});
