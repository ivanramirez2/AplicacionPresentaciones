@Code
    ViewData("Title") = "Dashboard - Metro Ligero Oeste"
    Layout = "~/Views/Shared/_Layout.vbhtml"
End Code

<div class="presentation-page">
    <div class="presentation-container">
        <div class="presentation-header">
            <img src="@Url.Content("~/Content/img/MLO-removebg-preview.png")" alt="Metro Ligero Oeste Logo" class="presentation-logo" />
            <div class="presentation-welcome">
                BIENVENIDO 
            </div>
            <div class="presentation-user-profile">
                <span class="presentation-timer">0:05</span>
                <img src="@Url.Content("~/Content/img/usuario_cuadrado.png")" alt="User" class="presentation-user-image" />
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
                <button class="presentation-button">IMPRIMIR SERVICIO</button>
                <button class="presentation-button">IMPRIMIR INCIDENCIAS</button>
            </div>
            <a href="#" class="presentation-finalize-button">FINALIZAR</a>
        </div>
    </div>
</div>

@Section Styles
    <link rel="stylesheet" href="@Url.Content("~/Content/css/PresentacionRealizada.css")" type="text/css" />
End Section