@ModelType Object
@Code
    ViewData("Title") = "Error"
    Layout = "~/Views/Shared/_Layout.vbhtml"
End Code

<div class="error-page">
    <h2>Error</h2>
    <p>@Model.Mensaje</p>
    <a href="@Url.Action("Index", "Home")">Volver a intentar</a>
</div>