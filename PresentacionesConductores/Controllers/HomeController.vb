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
    Function ProcesarFormulario(numeroEmpleado As String) As ActionResult
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
                ViewBag.ErrorMensaje = "Error inesperado: " & ex.Message
                Return View("Index")
            End Try
        End Using

        ' Verificar conductor, estado de presentación y servicio
        Using conn2 As New SqlConnection(connectionString2)
            Try
                conn2.Open()

                ' Verificar si el empleado está registrado como conductor
                Dim esConductorQuery As String = "SELECT COUNT(*) FROM TbPRESENTACIONES WHERE EMPLEADO = @cod_empleado AND CONDUCTOR IS NOT NULL"
                Using cmd As New SqlCommand(esConductorQuery, conn2)
                    cmd.Parameters.AddWithValue("@cod_empleado", numeroEmpleado)
                    Dim esConductor As Integer = Convert.ToInt32(cmd.ExecuteScalar())

                    If esConductor = 0 Then
                        ViewBag.ErrorMensaje = "LO SIENTO, EL EMPLEADO NO ESTÁ REGISTRADO COMO CONDUCTOR HABILITADO"
                        Return View("Index")
                    End If
                End Using

                ' Verificar si el empleado tiene un servicio asignado
                Dim servicioQuery As String = "SELECT TOP 1 SERVICIO FROM TbPRESENTACIONES WHERE EMPLEADO = @cod_empleado ORDER BY Id DESC"
                Using cmd As New SqlCommand(servicioQuery, conn2)
                    cmd.Parameters.AddWithValue("@cod_empleado", numeroEmpleado)
                    Dim servicio As String = TryCast(cmd.ExecuteScalar(), String)

                    If String.IsNullOrEmpty(servicio) Then
                        ViewBag.ErrorMensaje = "SIN SERVICIO ASIGNADO, CONTACTE CON SU RESPONSABLE"
                        Return View("Index")
                    End If
                End Using

                ' Obtener la hora programada desde la base de datos antes de actualizar cualquier campo
                Dim horaProgramadaQuery As String = "SELECT HPROGRAMADA FROM TbPRESENTACIONES WHERE EMPLEADO = @cod_empleado"
                Using cmd As New SqlCommand(horaProgramadaQuery, conn2)
                    cmd.Parameters.AddWithValue("@cod_empleado", numeroEmpleado)
                    Dim horaProgramada As String = TryCast(cmd.ExecuteScalar(), String)

                    ' Obtener la hora actual del sistema
                    Dim horaActual As String = DateTime.Now.ToString("HH:mm")

                    ' Verificar si el empleado ya se presentó antes de comprobar la hora
                    Dim presentadoQuery As String = "SELECT PRESENTADO FROM TbPRESENTACIONES WHERE EMPLEADO = @cod_empleado"
                    Using cmdPres As New SqlCommand(presentadoQuery, conn2)
                        cmdPres.Parameters.AddWithValue("@cod_empleado", numeroEmpleado)
                        Dim presentado As Integer = Convert.ToInt32(cmdPres.ExecuteScalar())

                        ' Si `PRESENTADO = 1`, permitir el acceso sin mostrar mensajes de error
                        If presentado = 1 Then
                            TempData("NumeroEmpleado") = numeroEmpleado
                            ' Obtener el nombre del empleado antes de redirigir
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


                        ' SI LLEGA 2H ANTES JUSTAS ENTRA, SI LLEGA 2H Y 1MIN O MAS YA NO ENTRA
                        ' SI LLEGA 2H TARDE JUSTAS ENTRA, SI LLEGA 2H Y 1MIN TARDE O MAS NO ENTRA

                        ' Comparar si el empleado llega demasiado temprano
                        ' SI LLEGA 2H ANTES JUSTAS ENTRA, SI LLEGA 2H Y 1MIN O MAS YA NO ENTRA
                        If Not String.IsNullOrEmpty(horaProgramada) Then
                            Dim horaProgramadaTime As TimeSpan = TimeSpan.Parse(horaProgramada)
                            Dim horaActualTime As TimeSpan = TimeSpan.Parse(horaActual)
                            Dim diferenciaHoras As Double = (horaProgramadaTime - horaActualTime).TotalHours
                            If horaActualTime < horaProgramadaTime AndAlso diferenciaHoras > 2 Then
                                Dim diferencia As TimeSpan = horaProgramadaTime - horaActualTime
                                Dim mensajeDiferencia As String = String.Format("{0} horas y {1} minutos", CInt(diferencia.TotalHours), diferencia.Minutes)
                                ViewBag.ErrorMensaje = "LLEGADA ADELANTADA. Hora programada: " & horaProgramada & ", Hora de llegada: " & horaActual & ". Llegó " & mensajeDiferencia & " antes. ESPERE Y PRESÉNTESE A LA HORA INDICADA."
                                Return View("Index") ' No guarda HREAL ni redirige
                            End If
                        End If

                        ' Comparar si el empleado llega demasiado tarde
                        ' SI LLEGA 2H TARDE JUSTAS ENTRA, SI LLEGA 2H Y 1MIN TARDE O MAS NO ENTRA
                        If Not String.IsNullOrEmpty(horaProgramada) Then
                            Dim horaProgramadaTime As TimeSpan = TimeSpan.Parse(horaProgramada)
                            Dim horaActualTime As TimeSpan = TimeSpan.Parse(horaActual)
                            Dim diferenciaHoras As Double = (horaActualTime - horaProgramadaTime).TotalHours
                            If horaActualTime > horaProgramadaTime AndAlso diferenciaHoras > 2 Then
                                Dim diferencia As TimeSpan = horaActualTime - horaProgramadaTime
                                Dim mensajeDiferencia As String = String.Format("{0} horas y {1} minutos", CInt(diferencia.TotalHours), diferencia.Minutes)
                                ViewBag.ErrorMensaje = "LLEGADA TARDE. Hora programada: " & horaProgramada & ", Hora de llegada: " & horaActual & ". Llegó " & mensajeDiferencia & " tarde. NO SE PERMITE LA ENTRADA."
                                Return View("Index") ' No guarda HREAL ni redirige
                            End If
                        End If

                        ' SOLO SI LLEGA A TIEMPO Y PRESENTADO = 0, GUARDAR `HREAL`
                        Dim actualizarHoraQuery As String = "UPDATE TbPRESENTACIONES SET HREAL = @hora WHERE EMPLEADO = @cod_empleado"
                        Using updateCmd As New SqlCommand(actualizarHoraQuery, conn2)
                            updateCmd.Parameters.AddWithValue("@hora", DateTime.Now.ToString("HH:mm"))
                            updateCmd.Parameters.AddWithValue("@cod_empleado", numeroEmpleado)
                            updateCmd.ExecuteNonQuery()
                        End Using
                    End Using
                End Using

                ' Obtener el nombre del empleado antes de redirigir (MLORRHH)
                Using conn1 As New SqlConnection(connectionString1)
                    Try
                        conn1.Open()
                        Dim nombreEmpleadoQuery As String = "SELECT NOMBRE FROM RRHH_PF_EMPLEADOS WHERE COD_EMPLEADO = @cod_empleado"
                        Using cmd As New SqlCommand(nombreEmpleadoQuery, conn1)
                            cmd.Parameters.AddWithValue("@cod_empleado", numeroEmpleado)
                            Dim nombreEmpleado As String = TryCast(cmd.ExecuteScalar(), String)

                            If Not String.IsNullOrEmpty(nombreEmpleado) Then
                                TempData("NombreEmpleado") = nombreEmpleado
                            Else
                                TempData("NombreEmpleado") = "Empleado desconocido"
                            End If
                        End Using
                    Catch ex As Exception
                        ViewBag.ErrorMensaje = "Error inesperado: " & ex.Message
                        Return View("Index")
                    End Try
                End Using

                ' Guardar número de empleado y permitir el acceso si todo está correcto
                TempData("NumeroEmpleado") = numeroEmpleado
                Return RedirectToAction("Presentacion")

            Catch ex As Exception
                ViewBag.ErrorMensaje = "Error inesperado: " & ex.Message
                Return View("Index")
            End Try
        End Using
    End Function

    Function ObtenerImagen(numeroEmpleado As Integer) As ActionResult
        Dim connectionString1 As String = ConfigurationManager.ConnectionStrings("SQLServerConnection1").ConnectionString

        Using conn As New SqlConnection(connectionString1)
            conn.Open()
            Dim query As String = "SELECT FOTO FROM RRHH_PF_EMPLEADOS WHERE COD_EMPLEADO = @cod_empleado"

            Using cmd As New SqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@cod_empleado", numeroEmpleado)
                Dim imagenBytes As Byte() = TryCast(cmd.ExecuteScalar(), Byte())

                If imagenBytes IsNot Nothing Then
                    Return File(imagenBytes, "image/jpeg") ' formato JPEG
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

    Function DescargarServicioExcel() As ActionResult
        ' 1. Obtener datos
        Dim servicio As String = If(TempData("Servicio"), "SIN SERVICIO ASIGNADO")

        ' 2. Copiar plantilla a un stream temporal
        Dim plantillaPath As String = Server.MapPath("~/Content/Plantillas/Plantilla.xlsx")
        Dim tempStream As New MemoryStream()
        Using fileStream As New FileStream(plantillaPath, FileMode.Open)
            fileStream.CopyTo(tempStream)
        End Using
        tempStream.Position = 0

        ' 3. Manipular el archivo con OpenXML
        Using package As SpreadsheetDocument = SpreadsheetDocument.Open(tempStream, True)
            Dim workbookPart As WorkbookPart = package.WorkbookPart
            Dim sheets As Sheets = workbookPart.Workbook.Sheets

            ' Buscar la hoja con el nombre del servicio
            Dim hojaServicio As Sheet = sheets.Elements(Of Sheet)().FirstOrDefault(Function(s) s.Name = servicio)

            If hojaServicio IsNot Nothing Then
                ' Identificar hojas que deben eliminarse
                Dim hojasAEliminar = sheets.Elements(Of Sheet)().Where(Function(s) s.Name <> hojaServicio.Name).ToList()

                For Each sheet In hojasAEliminar
                    Dim wsPart As WorksheetPart = workbookPart.GetPartById(sheet.Id)

                    ' **Eliminar completamente la hoja del libro**
                    workbookPart.DeletePart(wsPart)
                    sheet.Remove()
                Next

                ' Guardar cambios
                workbookPart.Workbook.Save()
            Else
                Dim sheetNames As String = String.Join(", ", sheets.Elements(Of Sheet)().Select(Function(s) s.Name))
                Throw New Exception($"Las hojas disponibles en el archivo son: {sheetNames}. No se encontró una hoja con el nombre '{servicio}'")
            End If
        End Using

        ' 4. Devolver el archivo modificado
        tempStream.Position = 0
        Return File(tempStream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Servicio.xlsx")
    End Function
End Class