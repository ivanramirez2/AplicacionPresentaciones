Imports System.Web.Mvc
Imports System.Configuration
Imports System.Web
Imports System.Data.SqlClient
Imports iTextSharp.text
Imports iTextSharp.text.pdf
Imports System.IO
Imports System.Drawing
Imports System.Net.Mime.MediaTypeNames
Imports OfficeOpenXml
Imports System.IO.Packaging

Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet

Public Class HomeController
    Inherits System.Web.Mvc.Controller

    Function Index() As ActionResult
        Return View()
    End Function

    <HttpPost>
    Function ProcesarFormulario(numeroEmpleado As String, pin As String) As ActionResult
        Dim connectionString1 As String = ConfigurationManager.ConnectionStrings("SQLServerConnection1").ConnectionString 'MLORRHH'
        Dim connectionString2 As String = ConfigurationManager.ConnectionStrings("SQLServerConnection2").ConnectionString 'MLOOPER'

        ' Verificar si el empleado existe en el sistema (MLORRHH)
        Using conn1 As New SqlConnection(connectionString1)
            Try
                conn1.Open()

                Dim empleadoExisteQuery As String = "SELECT COUNT(*) FROM RRHH_PF_EMPLEADOS WHERE COD_EMPLEADO = @cod_empleado"
                Using cmd As New SqlCommand(empleadoExisteQuery, conn1)
                    cmd.Parameters.AddWithValue("@cod_empleado", numeroEmpleado)
                    Dim empleadoExiste As Integer = Convert.ToInt32(cmd.ExecuteScalar())

                    If empleadoExiste = 0 Then
                        ViewBag.ErrorMensaje = "LO SIENTO, NÚMERO DE EMPLEADO NO EXISTE EN EL SISTEMA"
                        Return View("Index")
                    End If
                End Using
            Catch ex As Exception
                ViewBag.ErrorMensaje = "Error inesperado al verificar el empleado: " & ex.Message
                Return View("Index")
            End Try
        End Using

        ' Verificar si el empleado está registrado como conductor (MLOOPER)
        Using conn2 As New SqlConnection(connectionString2)
            Try
                conn2.Open()

                Dim esConductorQuery As String = "SELECT COUNT(*) FROM TbPRESENTACIONES WHERE EMPLEADO = @cod_empleado AND CONDUCTOR IS NOT NULL"
                Using cmd As New SqlCommand(esConductorQuery, conn2)
                    cmd.Parameters.AddWithValue("@cod_empleado", numeroEmpleado)
                    Dim esConductor As Integer = Convert.ToInt32(cmd.ExecuteScalar())

                    If esConductor = 0 Then
                        ViewBag.ErrorMensaje = "LO SIENTO, EL EMPLEADO NO ESTÁ REGISTRADO COMO CONDUCTOR HABILITADO"
                        Return View("Index")
                    End If
                End Using
            Catch ex As Exception
                ViewBag.ErrorMensaje = "Error inesperado al verificar el conductor: " & ex.Message
                Return View("Index")
            End Try
        End Using

        ' Verificar el PIN del empleado (MLORRHH)
        Using conn1 As New SqlConnection(connectionString1)
            Try
                conn1.Open()

                If String.IsNullOrWhiteSpace(pin) Then
                    ViewBag.ErrorMensaje = "El PIN no puede estar vacío. Por favor, ingrese un PIN válido."
                    Return View("Index")
                End If

                Dim pinQuery As String = "SELECT PIN FROM RRHH_PF_EMPLEADOS WHERE COD_EMPLEADO = @cod_empleado"
                Using cmd As New SqlCommand(pinQuery, conn1)
                    cmd.Parameters.AddWithValue("@cod_empleado", numeroEmpleado)
                    Dim pinAlmacenado As Object = cmd.ExecuteScalar()

                    If pinAlmacenado Is Nothing Then
                        ViewBag.ErrorMensaje = "No se encontró un PIN registrado para este empleado. Contacte con su responsable."
                        Return View("Index")
                    End If

                    Dim pinAlmacenadoStr As String = pinAlmacenado.ToString().Trim()
                    Dim pinIngresado As String = pin.Trim()

                    If Not String.Equals(pinAlmacenadoStr, pinIngresado, StringComparison.OrdinalIgnoreCase) Then
                        ViewBag.ErrorMensaje = "PIN INCORRECTO. Por favor, intenta de nuevo."
                        Return View("Index")
                    End If
                End Using
            Catch ex As Exception
                ViewBag.ErrorMensaje = "Error inesperado al verificar el PIN: " & ex.Message
                Return View("Index")
            End Try
        End Using

        ' Verificar si el empleado tiene un servicio asignado (MLOOPER)
        Using conn2 As New SqlConnection(connectionString2)
            Try
                conn2.Open()

                Dim servicioQuery As String = "SELECT TOP 1 SERVICIO FROM TbPRESENTACIONES WHERE EMPLEADO = @cod_empleado ORDER BY Id DESC"
                Using cmd As New SqlCommand(servicioQuery, conn2)
                    cmd.Parameters.AddWithValue("@cod_empleado", numeroEmpleado)
                    Dim servicio As String = TryCast(cmd.ExecuteScalar(), String)

                    If String.IsNullOrEmpty(servicio) Then
                        ViewBag.ErrorMensaje = "SIN SERVICIO ASIGNADO, CONTACTE CON SU RESPONSABLE"
                        Return View("Index")
                    End If
                End Using

                ' Obtener la hora programada desde la base de datos
                Dim horaProgramadaQuery As String = "SELECT HPROGRAMADA FROM TbPRESENTACIONES WHERE EMPLEADO = @cod_empleado"
                Using cmd As New SqlCommand(horaProgramadaQuery, conn2)
                    cmd.Parameters.AddWithValue("@cod_empleado", numeroEmpleado)
                    Dim horaProgramada As String = TryCast(cmd.ExecuteScalar(), String)

                    ' Obtener la hora actual del sistema
                    Dim horaActual As String = DateTime.Now.ToString("HH:mm")

                    ' Verificar si el empleado ya se presentó
                    Dim presentadoQuery As String = "SELECT PRESENTADO FROM TbPRESENTACIONES WHERE EMPLEADO = @cod_empleado"
                    Using cmdPres As New SqlCommand(presentadoQuery, conn2)
                        cmdPres.Parameters.AddWithValue("@cod_empleado", numeroEmpleado)
                        Dim presentado As Integer = Convert.ToInt32(cmdPres.ExecuteScalar())

                        If presentado = 1 Then
                            TempData("NumeroEmpleado") = numeroEmpleado
                            Using conn1 As New SqlConnection(connectionString1)
                                conn1.Open()
                                Dim nombreEmpleadoQuery As String = "SELECT NOMBRE FROM RRHH_PF_EMPLEADOS WHERE COD_EMPLEADO = @cod_empleado"
                                Using cmdNombre As New SqlCommand(nombreEmpleadoQuery, conn1)
                                    cmdNombre.Parameters.AddWithValue("@cod_empleado", numeroEmpleado)
                                    Dim nombreEmpleado As String = TryCast(cmdNombre.ExecuteScalar(), String)
                                    TempData("NombreEmpleado") = If(String.IsNullOrEmpty(nombreEmpleado), "Empleado desconocido", nombreEmpleado)
                                End Using
                            End Using
                            Return RedirectToAction("Presentacion")
                        End If

                        ' Verificar llegada adelantada (2 horas o más)
                        If Not String.IsNullOrEmpty(horaProgramada) Then
                            Dim horaProgramadaTime As TimeSpan = TimeSpan.Parse(horaProgramada)
                            Dim horaActualTime As TimeSpan = TimeSpan.Parse(horaActual)
                            Dim diferenciaHoras As Double = (horaProgramadaTime - horaActualTime).TotalHours
                            If horaActualTime < horaProgramadaTime AndAlso diferenciaHoras > 2 Then
                                Dim diferencia As TimeSpan = horaProgramadaTime - horaActualTime
                                Dim mensajeDiferencia As String = String.Format("{0} horas y {1} minutos", CInt(diferencia.TotalHours), diferencia.Minutes)
                                ViewBag.ErrorMensaje = "LLEGADA ADELANTADA. Hora programada: " & horaProgramada & ", Hora de llegada: " & horaActual & ". Llegó " & mensajeDiferencia & " antes. ESPERE Y PRESÉNTESE A LA HORA INDICADA."
                                Return View("Index")
                            End If
                        End If

                        ' Verificar llegada tarde (2 horas o más)
                        If Not String.IsNullOrEmpty(horaProgramada) Then
                            Dim horaProgramadaTime As TimeSpan = TimeSpan.Parse(horaProgramada)
                            Dim horaActualTime As TimeSpan = TimeSpan.Parse(horaActual)
                            Dim diferenciaHoras As Double = (horaActualTime - horaProgramadaTime).TotalHours
                            If horaActualTime > horaProgramadaTime AndAlso diferenciaHoras > 2 Then
                                Dim diferencia As TimeSpan = horaActualTime - horaProgramadaTime
                                Dim mensajeDiferencia As String = String.Format("{0} horas y {1} minutos", CInt(diferencia.TotalHours), diferencia.Minutes)
                                ViewBag.ErrorMensaje = "LLEGADA TARDE. Hora programada: " & horaProgramada & ", Hora de llegada: " & horaActual & ". Llegó " & mensajeDiferencia & " tarde. NO SE PERMITE LA ENTRADA."
                                Return View("Index")
                            End If
                        End If

                        ' Actualizar HREAL si llega a tiempo y PRESENTADO = 0
                        Dim actualizarHoraQuery As String = "UPDATE TbPRESENTACIONES SET HREAL = @hora WHERE EMPLEADO = @cod_empleado"
                        Using updateCmd As New SqlCommand(actualizarHoraQuery, conn2)
                            updateCmd.Parameters.AddWithValue("@hora", DateTime.Now.ToString("HH:mm"))
                            updateCmd.Parameters.AddWithValue("@cod_empleado", numeroEmpleado)
                            updateCmd.ExecuteNonQuery()
                        End Using
                    End Using
                End Using

                ' Obtener el nombre del empleado (MLORRHH)
                Using conn1 As New SqlConnection(connectionString1)
                    Try
                        conn1.Open()
                        Dim nombreEmpleadoQuery As String = "SELECT NOMBRE FROM RRHH_PF_EMPLEADOS WHERE COD_EMPLEADO = @cod_empleado"
                        Using cmd As New SqlCommand(nombreEmpleadoQuery, conn1)
                            cmd.Parameters.AddWithValue("@cod_empleado", numeroEmpleado)
                            Dim nombreEmpleado As String = TryCast(cmd.ExecuteScalar(), String)
                            TempData("NombreEmpleado") = If(String.IsNullOrEmpty(nombreEmpleado), "Empleado desconocido", nombreEmpleado)
                        End Using
                    Catch ex As Exception
                        ViewBag.ErrorMensaje = "Error al obtener el nombre del empleado: " & ex.Message
                        Return View("Index")
                    End Try
                End Using

                TempData("NumeroEmpleado") = numeroEmpleado
                Return RedirectToAction("Presentacion")
            Catch ex As Exception
                ViewBag.ErrorMensaje = "Error inesperado: " & ex.Message
                Return View("Index")
            End Try
        End Using
    End Function



    ' Función que obtiene la imagen del empleado'
    Function ObtenerImagen(numeroEmpleado As Integer) As ActionResult
        Dim connectionString1 As String = ConfigurationManager.ConnectionStrings("SQLServerConnection1").ConnectionString

        Using conn As New SqlConnection(connectionString1)
            conn.Open()
            Dim query As String = "SELECT FOTO FROM RRHH_PF_EMPLEADOS WHERE COD_EMPLEADO = @cod_empleado"

            Using cmd As New SqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@cod_empleado", numeroEmpleado)
                Dim imagenBytes As Byte() = TryCast(cmd.ExecuteScalar(), Byte())

                If imagenBytes IsNot Nothing Then
                    Return File(imagenBytes, "image/jpeg")
                Else
                    Return File("~/Content/Images/usuario_cuadrado.png", "image/png")
                End If
            End Using
        End Using
    End Function

    Function Presentacion() As ActionResult
        If TempData("NumeroEmpleado") IsNot Nothing Then
            ViewBag.NumeroEmpleado = TempData("NumeroEmpleado")
            TempData("NumeroEmpleado") = ViewBag.NumeroEmpleado
        Else
            ViewBag.NumeroEmpleado = 0 ' Valor por defecto
        End If

        If TempData("NombreEmpleado") IsNot Nothing Then
            ViewBag.NombreEmpleado = TempData("NombreEmpleado")
            TempData("NombreEmpleado") = ViewBag.NombreEmpleado
        Else
            ViewBag.NombreEmpleado = "Empleado desconocido"
        End If

        Debug.WriteLine("Número de empleado enviado a la vista: " & ViewBag.NumeroEmpleado)

        ' Verificar el valor de PRESENTADO en TbPRESENTACIONES y obtener SERVICIO e INCIDENCIAS
        Dim connectionString2 As String = ConfigurationManager.ConnectionStrings("SQLServerConnection2").ConnectionString
        Using conn As New SqlConnection(connectionString2)
            Try
                conn.Open()

                ' Obtener el servicio asignado al conductor (PRIMERO, antes de cualquier redirección)
                Dim servicioQuery As String = "SELECT TOP 1 SERVICIO FROM TbPRESENTACIONES WHERE EMPLEADO = @cod_empleado ORDER BY Id DESC"
                Using cmd As New SqlCommand(servicioQuery, conn)
                    cmd.Parameters.AddWithValue("@cod_empleado", ViewBag.NumeroEmpleado)
                    Dim servicioResult As Object = cmd.ExecuteScalar()
                    Debug.WriteLine("Presentacion: Resultado de SERVICIO (sin convertir): " & If(servicioResult Is Nothing, "Nothing", servicioResult.ToString()))
                    ViewBag.Servicio = TryCast(servicioResult, String)
                End Using

                ' Obtener las incidencias asignadas al conductor (si las hay)
                Dim incidenciasQuery As String = "SELECT TOP 1 INCIDENCIAS FROM TbPRESENTACIONES WHERE EMPLEADO = @cod_empleado ORDER BY Id DESC"
                Using cmd As New SqlCommand(incidenciasQuery, conn)
                    cmd.Parameters.AddWithValue("@cod_empleado", ViewBag.NumeroEmpleado)
                    Dim incidenciasResult As Object = cmd.ExecuteScalar()
                    Debug.WriteLine("Presentacion: Resultado de INCIDENCIAS (sin convertir): " & If(incidenciasResult Is Nothing, "Nothing", incidenciasResult.ToString()))
                    ViewBag.Incidencias = TryCast(incidenciasResult, String)
                End Using

                ' Verificar el valor de PRESENTADO en TbPRESENTACIONES (DESPUÉS de obtener los datos)
                Dim presentadoQuery As String = "SELECT PRESENTADO FROM TbPRESENTACIONES WHERE EMPLEADO = @cod_empleado"
                Using cmd As New SqlCommand(presentadoQuery, conn)
                    cmd.Parameters.AddWithValue("@cod_empleado", ViewBag.NumeroEmpleado)
                    Dim result As Object = cmd.ExecuteScalar()
                    Debug.WriteLine("Resultado de PRESENTADO (sin convertir): " & If(result Is Nothing, "Nothing", result.ToString()))

                    ' Verificar si se obtuvo un valor
                    If result Is Nothing Then
                        ViewBag.ErrorMensaje = "No se pudo obtener el estado de presentación del empleado."
                        Debug.WriteLine("PRESENTADO es Nothing para EMPLEADO: " & ViewBag.NumeroEmpleado)
                        Return View("Index")
                    End If

                    ' Convertir el valor de PRESENTADO 
                    Dim presentado As Integer
                    Try
                        If TypeOf result Is Boolean Then
                            presentado = If(Convert.ToBoolean(result), 1, 0)
                        Else
                            presentado = Convert.ToInt32(result)
                        End If
                        Debug.WriteLine("Valor de PRESENTADO (convertido): " & presentado)
                    Catch ex As Exception
                        ViewBag.ErrorMensaje = "Error al convertir el estado de presentación: " & ex.Message
                        Debug.WriteLine("Error al convertir PRESENTADO: " & ex.Message)
                        Return View("Index")
                    End Try

                    ' Redirigir según el valor de PRESENTADO
                    If presentado = 0 Then
                        Dim updateQuery As String = "UPDATE TbPRESENTACIONES SET PRESENTADO = 1 WHERE EMPLEADO = @cod_empleado"

                        Using updateCmd As New SqlCommand(updateQuery, conn)
                            updateCmd.Parameters.AddWithValue("@cod_empleado", ViewBag.NumeroEmpleado)
                            updateCmd.ExecuteNonQuery()
                        End Using
                    ElseIf presentado = 1 Then
                        ' Antes de redirigir, pasar los datos a TempData para que Presentacion2() los use
                        TempData("Servicio") = ViewBag.Servicio
                        TempData("Incidencias") = ViewBag.Incidencias
                        Return RedirectToAction("Presentacion2")
                    Else
                        ViewBag.ErrorMensaje = "Estado de presentación no válido: " & presentado
                        Return View("Index")
                    End If
                End Using

                ' Registrar la hora de presentación (solo si PRESENTADO era 0)
                Dim horaActual As String = DateTime.Now.ToString("HH:mm")
                Dim actualizarHoraQuery As String = "UPDATE TbPRESENTACIONES SET HREAL = @hora WHERE EMPLEADO = @cod_empleado"
                Using cmd As New SqlCommand(actualizarHoraQuery, conn)
                    cmd.Parameters.AddWithValue("@hora", horaActual)
                    cmd.Parameters.AddWithValue("@cod_empleado", ViewBag.NumeroEmpleado)

                    Debug.WriteLine("Actualizando hora para empleado: " & ViewBag.NumeroEmpleado)
                    Dim filasAfectadas As Integer = cmd.ExecuteNonQuery()
                    Debug.WriteLine("Filas afectadas al actualizar la hora: " & filasAfectadas)

                    If filasAfectadas > 0 Then
                        ViewBag.HoraPresentacion = horaActual ' Guarda la hora en ViewBag para mostrarla en la vista
                    Else
                        ViewBag.HoraPresentacion = "No se pudo registrar la hora"
                    End If
                End Using

            Catch ex As SqlException
                ViewBag.ErrorMensaje = "Error en la base de datos: " & ex.Message
                Debug.WriteLine("Error en SQL Server: " & ex.Message)
                Return View("Index")
            Catch ex As Exception
                ViewBag.ErrorMensaje = "Error inesperado: " & ex.Message
                Debug.WriteLine("Error general: " & ex.Message)
                Return View("Index")
            End Try
        End Using

        Return View()
    End Function

    Function Presentacion2() As ActionResult
        If TempData("NumeroEmpleado") IsNot Nothing Then
            ViewBag.NumeroEmpleado = TempData("NumeroEmpleado")
            TempData("NumeroEmpleado") = ViewBag.NumeroEmpleado
        Else
            ViewBag.NumeroEmpleado = 0
        End If

        If TempData("NombreEmpleado") IsNot Nothing Then
            ViewBag.NombreEmpleado = TempData("NombreEmpleado")
            TempData("NombreEmpleado") = ViewBag.NombreEmpleado
        Else
            ViewBag.NombreEmpleado = "Empleado desconocido"
        End If

        Debug.WriteLine("Número de empleado enviado a la vista: " & ViewBag.NumeroEmpleado)

        ' Usar los datos de TempData si están disponibles, si no, consultar la base de datos
        If TempData("Servicio") IsNot Nothing Then
            ViewBag.Servicio = TempData("Servicio")
            TempData("Servicio") = ViewBag.Servicio
        End If

        If TempData("Incidencias") IsNot Nothing Then
            ViewBag.Incidencias = TempData("Incidencias")
            TempData("Incidencias") = ViewBag.Incidencias
        End If

        ' Obtener el servicio e incidencias para el conductor solo si no están en TempData
        If ViewBag.Servicio Is Nothing Or ViewBag.Incidencias Is Nothing Then
            Dim connectionString2 As String = ConfigurationManager.ConnectionStrings("SQLServerConnection2").ConnectionString
            Using conn As New SqlConnection(connectionString2)
                Try
                    conn.Open()

                    ' Obtener el servicio asignado al conductor
                    Dim servicioQuery As String = "SELECT TOP 1 SERVICIO FROM TbPRESENTACIONES WHERE EMPLEADO = @cod_empleado ORDER BY Id DESC"
                    Using cmd As New SqlCommand(servicioQuery, conn)
                        cmd.Parameters.AddWithValue("@cod_empleado", ViewBag.NumeroEmpleado)
                        Dim servicioResult As Object = cmd.ExecuteScalar()
                        ViewBag.Servicio = TryCast(servicioResult, String)
                    End Using

                    ' Obtener las incidencias asignadas al conductor (si las hay)
                    Dim incidenciasQuery As String = "SELECT TOP 1 INCIDENCIAS FROM TbPRESENTACIONES WHERE EMPLEADO = @cod_empleado ORDER BY Id DESC"
                    Using cmd As New SqlCommand(incidenciasQuery, conn)
                        cmd.Parameters.AddWithValue("@cod_empleado", ViewBag.NumeroEmpleado)
                        Dim incidenciasResult As Object = cmd.ExecuteScalar()
                        ViewBag.Incidencias = TryCast(incidenciasResult, String)
                    End Using

                Catch ex As Exception
                    ViewBag.ErrorMensaje = "Error al obtener los datos: " & ex.Message
                End Try
            End Using
        End If

        Return View()
    End Function





    Function ImprimirServicioSinIncidenciaExcel() As ActionResult
        Try
            ' 1. Obtener datos del servicio y empleado
            Dim servicio As String = If(TempData("Servicio"), "SIN SERVICIO ASIGNADO")
            Dim employeeId As String = If(TempData("NumeroEmpleado") IsNot Nothing, TempData("NumeroEmpleado").ToString(), "")

            ' 2. Consultar tabla MLO para obtener DESC_TIPO_DIA
            Dim descTipoDia As String = ""
            Dim connectionString As String = ConfigurationManager.ConnectionStrings("SQLServerConnection3").ConnectionString

            Using conn As New SqlConnection(connectionString)
                Try
                    conn.Open()
                    ' Consulta para obtener DESC_TIPO_DIA haciendo JOIN entre las tablas
                    Dim queryMLO As String = "SELECT TOP 1 td.DESC_TIPO_DIA " &
                                        "FROM [MLO].[dbo].[VIGENCIAS] v " &
                                        "INNER JOIN [MLO].[dbo].[TIPOS_DIA] td ON v.ID_TIPO_DIA = td.ID_TIPO_DIA " &
                                        "INNER JOIN [MLO].[dbo].[TIPOS_DIA_RELACION_TORNOS] tdrt ON td.ID_RELACION_TORNO = tdrt.ID_RELACION_TORNO " &
                                        "WHERE v.FECHA_OPERACION <= GETDATE() " &
                                        "ORDER BY v.FECHA_OPERACION DESC"

                    Using cmd As New SqlCommand(queryMLO, conn)
                        Dim result = cmd.ExecuteScalar()
                        If result IsNot Nothing Then
                            descTipoDia = result.ToString().Trim()
                        End If
                    End Using
                Catch ex As Exception
                    Debug.WriteLine("Error al obtener DESC_TIPO_DIA: " & ex.Message)
                    Throw New Exception("Error al consultar la tabla MLO: " & ex.Message)
                End Try
            End Using

            ' 3. Seleccionar plantilla basada en DESC_TIPO_DIA
            Dim plantillaPath As String = ""
            Dim plantillasDirectory As String = Server.MapPath("~/Content/Plantillas/")

            If Not String.IsNullOrEmpty(descTipoDia) AndAlso descTipoDia.Length >= 2 Then
                ' Obtener las dos primeras letras de DESC_TIPO_DIA
                Dim prefijo As String = descTipoDia.Substring(0, 2).ToUpper()

                ' Buscar archivo Excel que comience con esas dos letras
                Dim archivosExcel() As String = Directory.GetFiles(plantillasDirectory, "*.xlsx")

                For Each archivo In archivosExcel
                    Dim nombreArchivo As String = Path.GetFileName(archivo).ToUpper()
                    If nombreArchivo.StartsWith(prefijo) Then
                        plantillaPath = archivo
                        Exit For
                    End If
                Next

                ' Si no se encuentra archivo con el prefijo, usar plantilla por defecto
                If String.IsNullOrEmpty(plantillaPath) Then
                    Debug.WriteLine($"No se encontró plantilla que comience con '{prefijo}'. Usando plantilla por defecto.")
                    plantillaPath = Server.MapPath("~/Content/Plantillas/Plantilla.xlsx")
                End If
            Else
                ' Si no hay DESC_TIPO_DIA válido, usar lógica por día de la semana como respaldo
                Debug.WriteLine("DESC_TIPO_DIA no válido. Usando selección por día de la semana.")
                Dim diaSemana As DayOfWeek = DateTime.Now.DayOfWeek

                Select Case diaSemana
                    Case DayOfWeek.Thursday, DayOfWeek.Friday
                        plantillaPath = Server.MapPath("~/Content/Plantillas/Plantilla.xlsx")
                End Select
            End If

            ' Verificar que la plantilla seleccionada existe
            If Not System.IO.File.Exists(plantillaPath) Then
                Throw New FileNotFoundException($"No se encontró la plantilla: {plantillaPath}")
            End If

            Debug.WriteLine($"Plantilla seleccionada: {Path.GetFileName(plantillaPath)} basada en DESC_TIPO_DIA: '{descTipoDia}'")

            Dim tempExcelPath As String = Server.MapPath("~/Temp/" & Guid.NewGuid().ToString() & ".xlsx")
            Dim tempPdfPath As String = Server.MapPath("~/Temp/" & Guid.NewGuid().ToString() & ".pdf")

            ' Asegurarse de que la carpeta Temp exista
            If Not Directory.Exists(Server.MapPath("~/Temp")) Then
                Directory.CreateDirectory(Server.MapPath("~/Temp"))
            End If

            ' Copiar la plantilla seleccionada al archivo temporal
            System.IO.File.Copy(plantillaPath, tempExcelPath)

            ' 4. Manipular el archivo Excel con OpenXML
            Using package As SpreadsheetDocument = SpreadsheetDocument.Open(tempExcelPath, True)
                Dim workbookPart As WorkbookPart = package.WorkbookPart
                Dim sheets As Sheets = workbookPart.Workbook.Sheets

                ' Buscar la hoja con el nombre del servicio
                Dim hojaServicio As Sheet = sheets.Elements(Of Sheet)().FirstOrDefault(Function(s) s.Name = servicio)

                If hojaServicio IsNot Nothing Then
                    ' Obtener la hoja seleccionada
                    Dim wsPartServicio As WorksheetPart = workbookPart.GetPartById(hojaServicio.Id)
                    Dim wsServicio As Worksheet = wsPartServicio.Worksheet
                    Dim sheetData As SheetData = wsServicio.GetFirstChild(Of SheetData)()

                    ' Crear o recuperar el WorkbookStylesPart y el Stylesheet mínimo
                    Dim stylesPart As WorkbookStylesPart = workbookPart.WorkbookStylesPart
                    If stylesPart Is Nothing Then
                        stylesPart = workbookPart.AddNewPart(Of WorkbookStylesPart)()
                        Dim ss As New Stylesheet()
                        ss.Fonts = New Fonts(New DocumentFormat.OpenXml.Spreadsheet.Font())
                        ss.Fills = New Fills(
                    New Fill(New PatternFill() With {.PatternType = PatternValues.None}),
                    New Fill(New PatternFill() With {.PatternType = PatternValues.Gray125}))
                        ss.Borders = New Borders(New Border())
                        ss.CellFormats = New CellFormats(New CellFormat())
                        ss.CellStyles = New CellStyles(New CellStyle() With {.Name = "Normal", .FormatId = 0, .BuiltinId = 0})
                        ss.Save()
                        stylesPart.Stylesheet = ss
                    End If

                    Dim ssheet As Stylesheet = stylesPart.Stylesheet
                    If ssheet.Fills Is Nothing Then
                        ssheet.Fills = New Fills()
                    End If

                    Dim yellowFillIndex As Integer = -1
                    Dim countFills As Integer = ssheet.Fills.ChildElements.Count
                    For i As Integer = 0 To countFills - 1
                        Dim f As Fill = CType(ssheet.Fills.ElementAt(i), Fill)
                        Dim pf As PatternFill = f.GetFirstChild(Of PatternFill)()
                        If pf IsNot Nothing AndAlso pf.PatternType = PatternValues.Solid Then
                            Dim fc As ForegroundColor = pf.ForegroundColor
                            If fc IsNot Nothing AndAlso fc.Rgb IsNot Nothing AndAlso fc.Rgb.Value.ToUpper() = "FFFF00" Then
                                yellowFillIndex = i
                                Exit For
                            End If
                        End If
                    Next

                    If yellowFillIndex = -1 Then
                        Dim yellowFill As New Fill(New PatternFill() With {
                    .PatternType = PatternValues.Solid,
                    .ForegroundColor = New ForegroundColor() With {.Rgb = "FFFF00"},
                    .BackgroundColor = New BackgroundColor() With {.Rgb = "FFFF00"}
                })
                        ssheet.Fills.Append(yellowFill)
                        ssheet.Fills.Count = CLng(ssheet.Fills.ChildElements.Count)
                        yellowFillIndex = ssheet.Fills.ChildElements.Count - 1
                    End If

                    ' ----- Consulta SQL: Si el empleado tiene incidencia -----
                    Dim incidenceCode As String = ""
                    Using conn As New SqlConnection(connectionString)
                        Try
                            conn.Open()
                            Dim query As String = "SELECT TOP 1 INCIDENCIAS FROM TbPRESENTACIONES WHERE EMPLEADO = @cod_empleado"
                            Using cmd As New SqlCommand(query, conn)
                                cmd.Parameters.AddWithValue("@cod_empleado", employeeId)
                                Dim result = cmd.ExecuteScalar()
                                If result IsNot Nothing Then
                                    incidenceCode = result.ToString()
                                End If
                            End Using
                        Catch ex As Exception
                            Debug.WriteLine("Error al obtener incidencia: " & ex.Message)
                        End Try
                    End Using

                    If Not String.IsNullOrEmpty(incidenceCode) Then
                        ' Recorrer todas las celdas de la hoja y, si su valor contiene el código de incidencia, aplicar el relleno amarillo
                        For Each row In sheetData.Elements(Of Row)()
                            For Each cell In row.Elements(Of Cell)()
                                ' Implementación inline de GetCellValue
                                Dim cellValue As String = ""
                                If cell IsNot Nothing AndAlso cell.CellValue IsNot Nothing Then
                                    cellValue = cell.CellValue.InnerText
                                    If cell.DataType IsNot Nothing AndAlso cell.DataType = CellValues.SharedString Then
                                        Dim sst = workbookPart.GetPartsOfType(Of SharedStringTablePart)().FirstOrDefault()
                                        If sst IsNot Nothing Then
                                            cellValue = sst.SharedStringTable.ElementAt(Integer.Parse(cellValue)).InnerText
                                        End If
                                    End If
                                End If

                                If Not String.IsNullOrEmpty(cellValue) AndAlso cellValue.Contains(incidenceCode) Then
                                    ' Preservar el formato existente y agregar solo el relleno amarillo
                                    Dim currentStyleIndex As UInt32 = 0
                                    If cell.StyleIndex IsNot Nothing Then
                                        currentStyleIndex = cell.StyleIndex
                                    End If

                                    ' Obtener el estilo actual
                                    Dim currentCellFormat As CellFormat = Nothing
                                    If currentStyleIndex < ssheet.CellFormats.Count.Value Then
                                        currentCellFormat = CType(ssheet.CellFormats.ElementAt(CInt(currentStyleIndex)), CellFormat)
                                    End If
                                    ' Crear nuevo formato basado en el formato actual
                                    Dim newCellFormat As New CellFormat()

                                    ' Copiar todas las propiedades del formato actual
                                    If currentCellFormat IsNot Nothing Then
                                        ' Copiar propiedades de numeración
                                        If currentCellFormat.NumberFormatId IsNot Nothing Then
                                            newCellFormat.NumberFormatId = currentCellFormat.NumberFormatId
                                        End If

                                        If currentCellFormat.ApplyNumberFormat IsNot Nothing Then
                                            newCellFormat.ApplyNumberFormat = currentCellFormat.ApplyNumberFormat
                                        End If

                                        ' Copiar propiedades de fuente
                                        If currentCellFormat.FontId IsNot Nothing Then
                                            newCellFormat.FontId = currentCellFormat.FontId
                                        End If

                                        If currentCellFormat.ApplyFont IsNot Nothing Then
                                            newCellFormat.ApplyFont = currentCellFormat.ApplyFont
                                        End If

                                        ' Copiar propiedades de bordes
                                        If currentCellFormat.BorderId IsNot Nothing Then
                                            newCellFormat.BorderId = currentCellFormat.BorderId
                                        End If

                                        If currentCellFormat.ApplyBorder IsNot Nothing Then
                                            newCellFormat.ApplyBorder = currentCellFormat.ApplyBorder
                                        End If

                                        ' Copiar propiedades de alineación
                                        If currentCellFormat.Alignment IsNot Nothing Then
                                            newCellFormat.Alignment = CType(currentCellFormat.Alignment.CloneNode(True), Alignment)
                                        End If

                                        If currentCellFormat.ApplyAlignment IsNot Nothing Then
                                            newCellFormat.ApplyAlignment = currentCellFormat.ApplyAlignment
                                        End If

                                        ' Copiar propiedades de protección
                                        If currentCellFormat.Protection IsNot Nothing Then
                                            newCellFormat.Protection = CType(currentCellFormat.Protection.CloneNode(True), Protection)
                                        End If

                                        If currentCellFormat.ApplyProtection IsNot Nothing Then
                                            newCellFormat.ApplyProtection = currentCellFormat.ApplyProtection
                                        End If
                                    Else
                                        ' Propiedades por defecto si no hay formato previo
                                        newCellFormat.FontId = 0
                                        newCellFormat.BorderId = 0
                                    End If

                                    ' Establecer el relleno amarillo
                                    newCellFormat.FillId = CUInt(yellowFillIndex)
                                    newCellFormat.ApplyFill = True

                                    ' Añadir el formato a la hoja de estilos
                                    ssheet.CellFormats.Append(newCellFormat)
                                    ssheet.CellFormats.Count = CLng(ssheet.CellFormats.ChildElements.Count)
                                    Dim newCellFormatIndex As UInt32 = CUInt(ssheet.CellFormats.ChildElements.Count - 1)

                                    ' Aplicar el nuevo estilo a la celda
                                    cell.StyleIndex = newCellFormatIndex
                                End If
                            Next
                        Next
                        stylesPart.Stylesheet.Save()
                        wsServicio.Save()
                        workbookPart.Workbook.Save()
                    End If
                Else
                    Dim sheetNames As String = String.Join(", ", sheets.Elements(Of Sheet)().Select(Function(s) s.Name))
                    Throw New Exception($"Las hojas disponibles en el archivo son: {sheetNames}. No se encontró una hoja con el nombre '{servicio}'")
                End If
            End Using

            ' 5. Convertir Excel a PDF usando una instancia de Excel, especificando la hoja del servicio
            ConvertExcelToPdf(tempExcelPath, tempPdfPath, servicio)

            ' 6. Imprimir el PDF generado
            ImprimirPdf(tempPdfPath)

            ' 7. Limpiar archivos temporales
            Try
                If System.IO.File.Exists(tempExcelPath) Then
                    System.IO.File.Delete(tempExcelPath)
                End If
                If System.IO.File.Exists(tempPdfPath) Then
                    System.IO.File.Delete(tempPdfPath)
                End If
            Catch ex As Exception
                ' Registrar error de limpieza pero continuar
                Debug.WriteLine("Error al eliminar archivos temporales: " & ex.Message)
            End Try

            Return Content("Servicio impreso correctamente en PDF")
        Catch ex As Exception
            Return Content("Error al procesar e imprimir el servicio: " & ex.Message)
        End Try
    End Function


    Function ImprimirIncidenciaExcel() As ActionResult
        Try
            ' 1. Obtener datos del servicio y empleado
            Dim servicio As String = If(TempData("Servicio"), "SIN SERVICIO ASIGNADO")
            Dim employeeId As String = If(TempData("NumeroEmpleado") IsNot Nothing, TempData("NumeroEmpleado").ToString(), "")

            ' 2. Generar nombre del archivo de incidencias basado en la fecha actual
            Dim fechaActual As DateTime = DateTime.Now
            Dim nombresDias As String() = {"DOMINGO", "LUNES", "MARTES", "MIÉRCOLES", "JUEVES", "VIERNES", "SÁBADO"}
            Dim nombresMeses As String() = {"", "ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"}

            Dim diaNombre As String = nombresDias(fechaActual.DayOfWeek)
            Dim diaNumero As String = fechaActual.Day.ToString()
            Dim mesNombre As String = nombresMeses(fechaActual.Month)

            Dim nombreArchivoIncidencias As String = $"INCIDENCIAS {diaNombre} {diaNumero} {mesNombre}.xlsx"

            ' 3. Buscar el archivo con el nombre generado en la carpeta de plantillas
            Dim plantillaPath As String = Server.MapPath($"~/Content/Plantillas/{nombreArchivoIncidencias}")

            ' Si no existe el archivo específico, usar plantilla por defecto
            If Not System.IO.File.Exists(plantillaPath) Then
                Debug.WriteLine($"No se encontró el archivo: {nombreArchivoIncidencias}. Usando plantilla por defecto.")
                plantillaPath = Server.MapPath("~/Content/Plantillas/Plantilla.xlsx")
            Else
                Debug.WriteLine($"Usando archivo: {nombreArchivoIncidencias}")
            End If

            Dim tempExcelPath As String = Server.MapPath("~/Temp/" & Guid.NewGuid().ToString() & ".xlsx")
            Dim tempPdfPath As String = Server.MapPath("~/Temp/" & Guid.NewGuid().ToString() & ".pdf")

            ' Asegurarse de que la carpeta Temp exista
            If Not Directory.Exists(Server.MapPath("~/Temp")) Then
                Directory.CreateDirectory(Server.MapPath("~/Temp"))
            End If

            ' Copiar plantilla al archivo temporal
            System.IO.File.Copy(plantillaPath, tempExcelPath)

            ' 4. Manipular el archivo Excel con OpenXML
            Using package As SpreadsheetDocument = SpreadsheetDocument.Open(tempExcelPath, True)
                Dim workbookPart As WorkbookPart = package.WorkbookPart
                Dim sheets As Sheets = workbookPart.Workbook.Sheets

                ' Buscar la hoja con el nombre del servicio
                Dim hojaServicio As Sheet = sheets.Elements(Of Sheet)().FirstOrDefault(Function(s) s.Name = servicio)

                If hojaServicio IsNot Nothing Then
                    ' Obtener la hoja seleccionada
                    Dim wsPartServicio As WorksheetPart = workbookPart.GetPartById(hojaServicio.Id)
                    Dim wsServicio As Worksheet = wsPartServicio.Worksheet
                    Dim sheetData As SheetData = wsServicio.GetFirstChild(Of SheetData)()

                    ' Crear o recuperar el WorkbookStylesPart y el Stylesheet mínimo
                    Dim stylesPart As WorkbookStylesPart = workbookPart.WorkbookStylesPart
                    If stylesPart Is Nothing Then
                        stylesPart = workbookPart.AddNewPart(Of WorkbookStylesPart)()
                        Dim ss As New Stylesheet()
                        ss.Fonts = New Fonts(New DocumentFormat.OpenXml.Spreadsheet.Font())
                        ss.Fills = New Fills(
                    New Fill(New PatternFill() With {.PatternType = PatternValues.None}),
                    New Fill(New PatternFill() With {.PatternType = PatternValues.Gray125}))
                        ss.Borders = New Borders(New Border())
                        ss.CellFormats = New CellFormats(New CellFormat())
                        ss.CellStyles = New CellStyles(New CellStyle() With {.Name = "Normal", .FormatId = 0, .BuiltinId = 0})
                        ss.Save()
                        stylesPart.Stylesheet = ss
                    End If

                    Dim ssheet As Stylesheet = stylesPart.Stylesheet
                    If ssheet.Fills Is Nothing Then
                        ssheet.Fills = New Fills()
                    End If

                    Dim yellowFillIndex As Integer = -1
                    Dim countFills As Integer = ssheet.Fills.ChildElements.Count
                    For i As Integer = 0 To countFills - 1
                        Dim f As Fill = CType(ssheet.Fills.ElementAt(i), Fill)
                        Dim pf As PatternFill = f.GetFirstChild(Of PatternFill)()
                        If pf IsNot Nothing AndAlso pf.PatternType = PatternValues.Solid Then
                            Dim fc As ForegroundColor = pf.ForegroundColor
                            If fc IsNot Nothing AndAlso fc.Rgb IsNot Nothing AndAlso fc.Rgb.Value.ToUpper() = "FFFF00" Then
                                yellowFillIndex = i
                                Exit For
                            End If
                        End If
                    Next

                    If yellowFillIndex = -1 Then
                        Dim yellowFill As New Fill(New PatternFill() With {
                    .PatternType = PatternValues.Solid,
                    .ForegroundColor = New ForegroundColor() With {.Rgb = "FFFF00"},
                    .BackgroundColor = New BackgroundColor() With {.Rgb = "FFFF00"}
                })
                        ssheet.Fills.Append(yellowFill)
                        ssheet.Fills.Count = CLng(ssheet.Fills.ChildElements.Count)
                        yellowFillIndex = ssheet.Fills.ChildElements.Count - 1
                    End If

                    ' ----- Consulta SQL: Si el empleado tiene incidencia -----
                    Dim incidenceCode As String = ""
                    Dim connectionString As String = ConfigurationManager.ConnectionStrings("SQLServerConnection2").ConnectionString
                    Using conn As New SqlConnection(connectionString)
                        Try
                            conn.Open()
                            Dim query As String = "SELECT TOP 1 INCIDENCIAS FROM TbPRESENTACIONES WHERE EMPLEADO = @cod_empleado"
                            Using cmd As New SqlCommand(query, conn)
                                cmd.Parameters.AddWithValue("@cod_empleado", employeeId)
                                Dim result = cmd.ExecuteScalar()
                                If result IsNot Nothing Then
                                    incidenceCode = result.ToString()
                                End If
                            End Using
                        Catch ex As Exception
                            Debug.WriteLine("Error al obtener incidencia: " & ex.Message)
                        End Try
                    End Using

                    If Not String.IsNullOrEmpty(incidenceCode) Then
                        ' Recorrer todas las celdas de la hoja y, si su valor contiene el código de incidencia, aplicar el relleno amarillo
                        For Each row In sheetData.Elements(Of Row)()
                            For Each cell In row.Elements(Of Cell)()
                                ' Implementación inline de GetCellValue
                                Dim cellValue As String = ""
                                If cell IsNot Nothing AndAlso cell.CellValue IsNot Nothing Then
                                    cellValue = cell.CellValue.InnerText
                                    If cell.DataType IsNot Nothing AndAlso cell.DataType = CellValues.SharedString Then
                                        Dim sst = workbookPart.GetPartsOfType(Of SharedStringTablePart)().FirstOrDefault()
                                        If sst IsNot Nothing Then
                                            cellValue = sst.SharedStringTable.ElementAt(Integer.Parse(cellValue)).InnerText
                                        End If
                                    End If
                                End If

                                If Not String.IsNullOrEmpty(cellValue) AndAlso cellValue.Contains(incidenceCode) Then
                                    ' Preservar el formato existente y agregar solo el relleno amarillo
                                    Dim currentStyleIndex As UInt32 = 0
                                    If cell.StyleIndex IsNot Nothing Then
                                        currentStyleIndex = cell.StyleIndex
                                    End If

                                    ' Obtener el estilo actual
                                    Dim currentCellFormat As CellFormat = Nothing
                                    If currentStyleIndex < ssheet.CellFormats.Count.Value Then
                                        currentCellFormat = CType(ssheet.CellFormats.ElementAt(CInt(currentStyleIndex)), CellFormat)
                                    End If
                                    ' Crear nuevo formato basado en el formato actual
                                    Dim newCellFormat As New CellFormat()

                                    ' Copiar todas las propiedades del formato actual
                                    If currentCellFormat IsNot Nothing Then
                                        ' Copiar propiedades de numeración
                                        If currentCellFormat.NumberFormatId IsNot Nothing Then
                                            newCellFormat.NumberFormatId = currentCellFormat.NumberFormatId
                                        End If

                                        If currentCellFormat.ApplyNumberFormat IsNot Nothing Then
                                            newCellFormat.ApplyNumberFormat = currentCellFormat.ApplyNumberFormat
                                        End If

                                        ' Copiar propiedades de fuente
                                        If currentCellFormat.FontId IsNot Nothing Then
                                            newCellFormat.FontId = currentCellFormat.FontId
                                        End If

                                        If currentCellFormat.ApplyFont IsNot Nothing Then
                                            newCellFormat.ApplyFont = currentCellFormat.ApplyFont
                                        End If

                                        ' Copiar propiedades de bordes
                                        If currentCellFormat.BorderId IsNot Nothing Then
                                            newCellFormat.BorderId = currentCellFormat.BorderId
                                        End If

                                        If currentCellFormat.ApplyBorder IsNot Nothing Then
                                            newCellFormat.ApplyBorder = currentCellFormat.ApplyBorder
                                        End If

                                        ' Copiar propiedades de alineación
                                        If currentCellFormat.Alignment IsNot Nothing Then
                                            newCellFormat.Alignment = CType(currentCellFormat.Alignment.CloneNode(True), Alignment)
                                        End If

                                        If currentCellFormat.ApplyAlignment IsNot Nothing Then
                                            newCellFormat.ApplyAlignment = currentCellFormat.ApplyAlignment
                                        End If

                                        ' Copiar propiedades de protección
                                        If currentCellFormat.Protection IsNot Nothing Then
                                            newCellFormat.Protection = CType(currentCellFormat.Protection.CloneNode(True), Protection)
                                        End If

                                        If currentCellFormat.ApplyProtection IsNot Nothing Then
                                            newCellFormat.ApplyProtection = currentCellFormat.ApplyProtection
                                        End If
                                    Else
                                        ' Propiedades por defecto si no hay formato previo
                                        newCellFormat.FontId = 0
                                        newCellFormat.BorderId = 0
                                    End If

                                    ' Aplicar el relleno amarillo
                                    newCellFormat.FillId = CUInt(yellowFillIndex)
                                    newCellFormat.ApplyFill = New BooleanValue(True)

                                    ' Añadir el formato a la hoja de estilos
                                    ssheet.CellFormats.Append(newCellFormat)
                                    ssheet.CellFormats.Count = CLng(ssheet.CellFormats.ChildElements.Count)
                                    Dim newCellFormatIndex As UInt32 = CUInt(ssheet.CellFormats.ChildElements.Count - 1)

                                    ' Aplicar el nuevo estilo a la celda
                                    cell.StyleIndex = newCellFormatIndex
                                End If
                            Next
                        Next
                        stylesPart.Stylesheet.Save()
                        wsServicio.Save()
                        workbookPart.Workbook.Save()
                    End If
                Else
                    Dim sheetNames As String = String.Join(", ", sheets.Elements(Of Sheet)().Select(Function(s) s.Name))
                    Throw New Exception($"Las hojas disponibles en el archivo son: {sheetNames}. No se encontró una hoja con el nombre '{servicio}'")
                End If
            End Using

            ' 5. Convertir Excel a PDF usando una instancia de Excel, especificando la hoja del servicio
            ConvertExcelToPdf(tempExcelPath, tempPdfPath, servicio)

            ' 6. Imprimir el PDF generado
            ImprimirPdf(tempPdfPath)

            ' 7. Limpiar archivos temporales
            Try
                If System.IO.File.Exists(tempExcelPath) Then
                    System.IO.File.Delete(tempExcelPath)
                End If
                If System.IO.File.Exists(tempPdfPath) Then
                    System.IO.File.Delete(tempPdfPath)
                End If
            Catch ex As Exception
                ' Registrar error de limpieza pero continuar
                Debug.WriteLine("Error al eliminar archivos temporales: " & ex.Message)
            End Try

            Return Content("Incidencia impresa correctamente en PDF")
        Catch ex As Exception
            Return Content("Error al procesar e imprimir la incidencia: " & ex.Message)
        End Try
    End Function


    ' Convertir solo la hoja especificada de Excel a PDF usando Microsoft.Office.Interop.Excel
    Private Sub ConvertExcelToPdf(excelPath As String, pdfPath As String, sheetName As String)
        Dim excelApp As Object = Nothing
        Dim workbooks As Object = Nothing
        Dim workbook As Object = Nothing
        Dim worksheet As Object = Nothing

        Try
            ' Crear instancia de Excel
            excelApp = CreateObject("Excel.Application")
            excelApp.Visible = False
            excelApp.DisplayAlerts = False

            ' Abrir el libro de trabajo
            workbooks = excelApp.Workbooks
            workbook = workbooks.Open(excelPath)

            ' Obtener la hoja específica por nombre
            worksheet = workbook.Sheets(sheetName)

            ' Exportar solo esa hoja como PDF
            worksheet.ExportAsFixedFormat(0, pdfPath) ' 0 = xlTypePDF
        Catch ex As Exception
            Throw New Exception("Error al convertir Excel a PDF: " & ex.Message)
        Finally
            ' Limpiar objetos COM
            If worksheet IsNot Nothing Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet)
                worksheet = Nothing
            End If

            If workbook IsNot Nothing Then
                workbook.Close(False)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook)
                workbook = Nothing
            End If

            If workbooks IsNot Nothing Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbooks)
                workbooks = Nothing
            End If

            If excelApp IsNot Nothing Then
                excelApp.Quit()
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp)
                excelApp = Nothing
            End If

            ' Forzar la recolección de basura
            GC.Collect()
            GC.WaitForPendingFinalizers()
        End Try
    End Sub

    ' Método para imprimir PDF utilizando SumatraPDF (sin cambios)
    Private Sub ImprimirPdf(pdfPath As String)
        ' Ruta real al ejecutable de SumatraPDF
        Dim sumatraPath As String = "C:\Users\ivan.ramirez\AppData\Local\SumatraPDF\SumatraPDF.exe"
        If Not System.IO.File.Exists(sumatraPath) Then
            Throw New FileNotFoundException("No se encontró SumatraPDF en la ruta: " & sumatraPath)
        End If

        ' Configurar el proceso para imprimir usando SumatraPDF en la impresora predeterminada
        Dim psi As New ProcessStartInfo()
        psi.FileName = sumatraPath
        psi.Arguments = String.Format("-print-to-default ""{0}""", pdfPath)
        psi.CreateNoWindow = True
        psi.WindowStyle = ProcessWindowStyle.Hidden
        psi.UseShellExecute = False

        Dim proc As Process = Process.Start(psi)
        ' Esperar hasta 10 segundos para que finalice el proceso de impresión
        proc.WaitForExit(10000)

        ' Si el proceso aún no ha finalizado, lo cerramos
        If Not proc.HasExited Then
            proc.Kill()
        End If
    End Sub
End Class