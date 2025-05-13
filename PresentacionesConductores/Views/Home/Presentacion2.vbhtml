@Code
    Layout = "~/Views/Shared/_Layout.vbhtml"
    ViewBag.Title = "Presentacion Realizada"
End Code

<div class="presentation-page">
    <div class="presentation-container">
        <div class="presentation-header">
            <img src="@Url.Content("~/Content/Images/mlo_logo.png")" alt="Metro Ligero Oeste Logo" class="presentation-logo" />
            <div class="presentation-welcome">
                BIENVENIDO,@ViewBag.NombreEmpleado
            </div>
            <div class="presentation-user-profile">
                <span class="presentation-timer">0:30</span>

                <img src="/Home/ObtenerImagen?numeroEmpleado=@ViewBag.NumeroEmpleado"
                     alt="User" class="presentation-user-image" />
            </div>
        </div>


            <div class="presentation-main-content">
                <div class="presentation-message">
                    <span class="message-text">YA TE HAS PRESENTADO</span>
                </div>
                <div class="presentation-section">
                    <label>TU SERVICIO PARA HOY ES:</label>
                    <textarea class="presentation-textarea" readonly>@ViewBag.Servicio</textarea>
                </div>
                <div class="presentation-section">
                    <label>INCIDENCIAS A CUBRIR:</label>
                    <textarea class="presentation-textarea" readonly>@ViewBag.Incidencias</textarea>
                </div>
            </div>
            <div class="presentation-button-container">
                <div class="presentation-button-wrapper">
                    <form method="get" action="/Home/DescargarServicioExcel">
                        <button class="presentation-button" type="submit">IMPRIMIR SERVICIO</button>
                    </form>

                    <form method="get" action="/Home/GenerarPDF">
                        <button class="presentation-button" type="submit">IMPRIMIR INCIDENCIA</button>
                    </form>
                </div>
                <button class="presentation-finalize-button">FINALIZAR</button>
            </div>
        </div>
</div>

@Section Styles
    <link rel="stylesheet" href="~/Content/Presentacionrealizada.css" type="text/css" />
End Section


<script>
    console.log("Prueba directa desde Presentacion2");
</script>
<script src="@Url.Content("~/Scripts/ContadorYFinalizar.js")" defer></script>

@Section Scripts
    <script>
        console.log("SECCIÓN Scripts en Presentacion2 funcionando");
    </script>
    <script src="@Url.Content("~/Scripts/ContadorYFinalizar.js")" defer></script>
End Section
