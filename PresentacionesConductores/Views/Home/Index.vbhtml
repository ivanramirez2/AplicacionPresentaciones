@Code
    Layout = "~/Views/Shared/_Layout.vbhtml"
    ViewBag.Title = "Inicio"
    Dim numeroEmpleado As String = ""
    Dim pin As String = ""
End Code





<div class="home-page">
    <div class="inicio-container">
        <div class="header">
            <img src="@Url.Content("~/Content/Images/mlo_logo.png")" alt="Metroligero Oeste" class="logo" />
            <p class="date"></p>
        </div>

        <h1>CONTROL DE PRESENTACIÓN DE PERSONAL DE CONDUCCIÓN</h1>

        <div class="form-container">
            <form id="dataForm" method="post" action="@Url.Action("ProcesarFormulario", "Home")">
                <div class="keypad-and-inputs">
                    <div class="keypad">
                        <button type="button" id="btn7">7</button>
                        <button type="button" id="btn8">8</button>
                        <button type="button" id="btn9">9</button>
                        <button type="button" id="btn4">4</button>
                        <button type="button" id="btn5">5</button>
                        <button type="button" id="btn6">6</button>
                        <button type="button" id="btn3">3</button>
                        <button type="button" id="btn2">2</button>
                        <button type="button" id="btn1">1</button>
                        <button type="button" class="clear-btn" id="btnClear">C</button>
                        <button type="button" id="btn0">0</button>
                        <button type="submit" class="submit-btn" id="btnSubmit">→</button>
                    </div>
                    <div class="inputs">
                        <div class="form-group">
                            <label for="numeroEmpleado">NÚMERO DE EMPLEADO</label>
                            <div class="display" id="displayNumeroEmpleado" data-is-password="false">@numeroEmpleado</div>
                            <input type="hidden" id="numeroEmpleado" name="numeroEmpleado" value="@numeroEmpleado" />
                        </div>
                        <div class="form-group">
                            <label for="pin">PIN</label>
                            <div class="display" id="displayPin" data-is-password="true"></div>
                            <input type="hidden" id="pin" name="pin" value="@pin" />
                        </div>
                    </div>
                </div>
            </form>
        </div>
    </div>
</div>

<!--  Pop-up para mostrar errores -->
<div id="popup" class="popup-container">
    <div class="popup-content">
        <p id="popupMessage">@ViewBag.ErrorMensaje</p>
        <button onclick="cerrarPopup()">Aceptar</button>

    </div>
</div>

@Section Styles
    <link href="@Url.Content("~/Content/Home.css")" rel="stylesheet" type="text/css" />
End Section

@Section Scripts
    <script>
        function cerrarPopup() {
            console.log("✅ cerrarPopup() ha sido llamada desde Index.vbhtml...");
            const popup = document.getElementById("popup");
            if (popup) {
                popup.classList.remove("active");
                console.log("✅ Pop-up ocultado correctamente desde Index.vbhtml.");
            } else {
                console.error("❌ Error: No se encontró el pop-up en el DOM.");
            }
        }
    </script>
    <script src="@Url.Content("~/Scripts/Index.js")" defer></script>
End Section


