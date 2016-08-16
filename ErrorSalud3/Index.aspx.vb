Imports System.Data
Imports MySql.Data.MySqlClient
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Runtime.InteropServices
Imports System.IO
Imports ClosedXML.Excel

Public Class Index
    Inherits System.Web.UI.Page

    Dim claseprocedure As New Codprocedure

    Dim Nomre_Archivo As New DataTable

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    End Sub


    Protected Sub Button3_Click1(sender As Object, e As EventArgs) Handles Button3.Click
        claseprocedure.Ceros_izquierda()
        claseprocedure.Sexo_Hombre_Mal()
        claseprocedure.Sexo_Mujer_Mal()
        claseprocedure.Edadnodocumento()
        claseprocedure.Identificacionrepetida()
        claseprocedure.Identificacionvacia()
        claseprocedure.Primerapellidovacio()
        claseprocedure.Pnombrevacio()
        claseprocedure.Sexonoexiste()
        claseprocedure.Tipodocnoexiste()
        claseprocedure.Umenoexiste()
        claseprocedure.Hij_papellido()
        claseprocedure.Hij_sapellido()
        claseprocedure.Hij_snombre()
        Llenar_Grid()
        'Longitudmax()
    End Sub


    Protected Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim conect As New Conexion
        Try
            claseprocedure.Delete_tablas()
            FileUpload1.PostedFile.SaveAs(Server.MapPath("~/") + FileUpload1.FileName)
            Dim path As String = FileUpload1.PostedFile.FileName
            Dim x As String = Server.MapPath("~/")
            Dim source As String = Replace(x, "\", "/")
            If Not String.IsNullOrEmpty(path) Then
                Dim Conectar_ As New MySqlConnection(conect.CrearConexion.ConnectionString)
                Dim cmd As New MySqlCommand("LOAD DATA LOCAL INFILE " + "'" + source + path + "'" + " INTO TABLE datosusutext FIELDS TERMINATED BY ','", Conectar_)
                Conectar_.Open()
                cmd.ExecuteNonQuery()
                Conectar_.Close()
                claseprocedure.insertar_fichero(path)
            ElseIf String.IsNullOrEmpty(path) Then
                ClientScript.RegisterStartupScript(Me.GetType(), "alert", "alert('Debe seleccionar un archivo .txt');", True)
            End If
            My.Computer.FileSystem.DeleteFile(Server.MapPath("~/") + path)

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub


    Private Sub Llenar_Grid()
        Dim dt As New DataTable()
        'columnas
        dt.Columns.Add("Nombre Archivo")
        dt.Columns.Add("Numero de Registros")
        dt.Columns.Add("Registros Erroneos")
        '  GridViewDatos.DataSource =
        Dim llenar_grid_temp As New DataSet
        llenar_grid_temp = claseprocedure.Llenar
        Dim row As DataRow = dt.NewRow()
        Dim fi As String = claseprocedure.nombre_fichero()
        Dim fich As String = Replace(fi, ".txt", ".xlsx")
        row("Nombre Archivo") = fich
        row("Numero de Registros") = llenar_grid_temp.Tables(0).Rows(0).Item(0).ToString()
        row("Registros Erroneos") = llenar_grid_temp.Tables(0).Rows(0).Item(1).ToString()
        dt.Rows.Add(row)
        GridView1.DataSource = dt
        GridView1.DataBind()

    End Sub


    'Protected Sub btnExcel_Click()

    'Response.Clear()
    'Response.AddHeader("Content-Disposition", "attachment;filename=data.xls")
    'Response.ContentType = "application/vnd.ms-excel"
    'Dim sw As StringWriter = New StringWriter()
    'Dim htw As HtmlTextWriter = New HtmlTextWriter(sw)
    'GridView1.RenderControl(htw)
    'Response.Write(sw.ToString())
    'Response.End()
    'End Sub


    Private Sub GridView1_RowEditing(sender As Object, e As GridViewEditEventArgs) Handles GridView1.RowEditing

    End Sub


    Public Sub ExportExcel()
        Using wb As New XLWorkbook()
            wb.Worksheets.Add(claseprocedure.tabla_Errores, "datosusuerrores")
            Response.Clear()
            Response.Buffer = True
            Response.Charset = ""
            Dim ex As String = claseprocedure.nombre_fichero()
            Dim exc As String = Replace(ex, ".txt", ".xlsx")
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            Response.AddHeader("content-disposition", "attachment;filename=" + exc)
            Using MyMemoryStream As New MemoryStream()
                wb.SaveAs(MyMemoryStream)
                MyMemoryStream.WriteTo(Response.OutputStream)
                Response.Flush()
                Response.End()
            End Using
        End Using

    End Sub

    Protected Sub btnDescarga_Click(sender As Object, e As EventArgs)
        'claseprocedure.Truncatedatoserrores()
        ' claseprocedure.Truncatedatostext()
        'claseprocedure.Truncatefichero()
        ExportExcel()
    End Sub
End Class