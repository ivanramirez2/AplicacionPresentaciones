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
    const displayPin = document.getElementById('displayPin'); // Nuevo elemento para el PIN
    const popup = document.getElementById('popup');
    const popupMessage = document.getElementById('popupMessage');

    // Verificar que los elementos existan
    if (!displayNumeroEmpleado || !displayPin) {
        console.error('Error: No se encontraron los elementos displayNumeroEmpleado o displayPin en el DOM.');
        return;
    }

    // Estado inicial
    let activeDisplay = 'numeroEmpleado'; // Display activo por defecto
    let currentNumeroEmpleado = displayNumeroEmpleado.textContent || '';
    let currentPin = ''; // Variable para almacenar el PIN

    // Mostrar valores iniciales
    displayNumeroEmpleado.textContent = currentNumeroEmpleado || " ";
    displayPin.textContent = currentPin ? '*'.repeat(currentPin.length) : ''; // Mostrar PIN como asteriscos

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

    // Función para actualizar la fecha y hora
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

 
    const numberButtons = ["btn0", "btn1", "btn2", "btn3", "btn4", "btn5", "btn6", "btn7", "btn8", "btn9"];

    numberButtons.forEach(id => {
        const btnElement = document.getElementById(id);
        if (btnElement) {
            btnElement.addEventListener('click', function () {
                appendNumber(btnElement.textContent.trim());
            });
        }
    });

    // Botón para limpiar ambos campos
    const btnClear = document.getElementById('btnClear');
    if (btnClear) {
        btnClear.addEventListener('click', function () {
            currentNumeroEmpleado = '';
            currentPin = '';
            displayNumeroEmpleado.textContent = '';
            displayPin.textContent = '';
        });
    }

    // Botón para enviar el formulario
    const btnSubmit = document.getElementById('btnSubmit');
    if (btnSubmit) {
        btnSubmit.addEventListener('click', function () {
            const hiddenNumeroEmpleado = document.getElementById('numeroEmpleado');
            const hiddenPin = document.getElementById('pin'); // Nuevo campo oculto para el PIN

            if (hiddenNumeroEmpleado && hiddenPin) {
             
                hiddenNumeroEmpleado.value = currentNumeroEmpleado;
                hiddenPin.value = currentPin;
                document.getElementById('dataForm').submit();
            } else {
                console.error('Error: No se encontraron los campos ocultos numeroEmpleado o pin en el DOM.');
            }
        });
    }

    // Alternar entre displays al hacer clic
    displayNumeroEmpleado.addEventListener('click', function () {
        setActiveDisplay('numeroEmpleado');
    });

    displayPin.addEventListener('click', function () {
        setActiveDisplay('pin');
    });

    // Establecer display activo por defecto
    setActiveDisplay('numeroEmpleado');

    // Función para alternar el display activo
    function setActiveDisplay(displayId) {
        activeDisplay = displayId;
        displayNumeroEmpleado.classList.toggle('active', displayId === 'numeroEmpleado');
        displayPin.classList.toggle('active', displayId === 'pin');
    }

    // Función para añadir números al display activo
    function appendNumber(number) {
        if (activeDisplay === 'numeroEmpleado') {
            currentNumeroEmpleado += number;
            displayNumeroEmpleado.textContent = currentNumeroEmpleado;
        } else if (activeDisplay === 'pin') {
            currentPin += number;
            displayPin.textContent = '*'.repeat(currentPin.length); // Mostrar asteriscos
        }
    }
});