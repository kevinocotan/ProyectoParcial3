Imports System.Data.SqlClient

Public Class Devoluciones
    Public conn As New SqlConnection(My.Settings.conexion)
    Public lector As SqlDataReader

    '--validacion de campos

    Private Sub TextBox2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox2.KeyPress

        If Char.IsNumber(e.KeyChar) Then
            e.Handled = False
        ElseIf Char.IsControl(e.KeyChar) Then
            e.Handled = False
        ElseIf Char.IsSeparator(e.KeyChar) Then
            e.Handled = False
        Else
            e.Handled = True
            MsgBox("En este campo solo se permiten numeros. ", MsgBoxStyle.Critical, "Error, no se han ingresado el dato ")
        End If

    End Sub

    '---Botón Cargar

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        conn.Open()

        Dim comando As New SqlCommand("", conn)

        Try
            comando.CommandType = CommandType.Text
            comando.CommandText = "select * from Devoluciones"

            lector = comando.ExecuteReader
            lector.Read()

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        If (lector.HasRows()) Then
            TextBox1.Text = lector(0)
            TextBox2.Text = lector(1)
            TextBox3.Text = lector(2)
            ListBox1.Items.Clear()

            Do
                ListBox1.Items.Add(lector("Id_Devolucion") & "                          |              " & lector("Id_Prestamo") & "                |         " & lector("Fecha_Devolucion"))

                TextBox1.Text = ""
                TextBox2.Text = ""
                TextBox3.Text = ""
                TextBox1.Focus()

            Loop While (lector.Read())
        Else
            MsgBox("Para cargar los datos, primero ingreselos en el botón de ""Nuevo""  ", MsgBoxStyle.Critical, "Error, no se pudo cargar los datos")
        End If

        conn.Close()
    End Sub

    '---Botón Buscar

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        conn.Open()
        Dim comando As New SqlCommand("", conn)

        Try
            comando.CommandType = CommandType.Text
            comando.CommandText = "select * from Devoluciones where id_devolucion=" & Integer.Parse(TextBox1.Text)

            lector = comando.ExecuteReader
            lector.Read()
            If (lector.HasRows()) Then
                TextBox1.Text = lector(0)
                TextBox2.Text = lector(1)
                TextBox3.Text = lector(2)

            Else
                MsgBox("El dato que se ha ingresado, no existe ", MsgBoxStyle.Critical, "Error")

            End If
        Catch ex As Exception
            MsgBox("Para buscar un dato, por favor ingrese el Codigo primero", MsgBoxStyle.Critical, "Error, no se han ingresado datos")
        End Try
        conn.Close()
    End Sub

    '---Botón Nuevo

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        If (Button3.Text = "Nuevo") Then
            Button3.Text = "Guardar"
            TextBox1.Text = ""
            TextBox2.Text = ""
            TextBox3.Text = ""
            TextBox1.Focus()

        Else
            If (TextBox2.Text <> "") And (TextBox3.Text <> "") Then

                conn.Open()
                Button3.Text = "Nuevo"
                Dim comando As New SqlCommand("", conn)

                Try
                    Dim validacion As Integer
                    validacion = TextBox1.Text

                    If validacion >= 1 Then

                        comando.CommandType = CommandType.Text
                        comando.CommandText = "insert into devoluciones Values (" & Integer.Parse(TextBox1.Text) & " , '" & TextBox2.Text & "' , '" & TextBox3.Text & "')"

                        comando.ExecuteNonQuery()
                        MsgBox("Se ingreso la nueva Devolución: " & TextBox2.Text & " " & TextBox3.Text & " ", MsgBoxStyle.Information, "Datos Ingresados con Éxito")

                        TextBox1.Text = ""
                        TextBox2.Text = ""
                        TextBox3.Text = ""
                        TextBox1.Focus()

                    Else
                        MsgBox("El código de Devolución no puede ser negativo. ", MsgBoxStyle.Critical, "Error, no se han ingresado el dato")
                    End If

                Catch ex As Exception
                    MsgBox("Por favor ingrese los datos que correspondan o revise si el Código de Devolucion o de Prestamo existe ", MsgBoxStyle.Critical, "Error al ingresar la Devolución")
                End Try
            Else
                MsgBox("Hay campos obligatorios que se encuentran vacios. Revise e intente de nuevo por favor. ", MsgBoxStyle.Critical, "Error al ingresar los datos")
            End If
        End If
        conn.Close()
    End Sub

    '---Boton Borrar

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        conn.Open()
        Dim eliminar As New SqlCommand("", conn)

        Try
            Dim validacion As Integer
            validacion = TextBox1.Text

            If validacion < 1 Then
                MsgBox("El código de Autor no puede ser negativo. ", MsgBoxStyle.Critical, "Error, no se han ingresado el dato")
            Else
                Try
                    eliminar.CommandType = CommandType.Text
                    eliminar.CommandText = "DELETE from Devoluciones where id_devolucion = " & Integer.Parse(TextBox1.Text)

                    eliminar.ExecuteNonQuery()
                    MsgBox("Se eliminó la Devolución con código: " & TextBox1.Text, MsgBoxStyle.Information, "Datos Eliminados con Éxito")
                Catch ex As Exception
                    MsgBox("Asegurese que este campo no está referenciado en otro Formulario.", MsgBoxStyle.Critical, "Error, al intentar borrar los datos")
                End Try
            End If

        Catch ex As Exception
            MsgBox("Este campo es obligatorio y debe de ser un número entero. ", MsgBoxStyle.Critical, "Error, no se pudo eliminar la Devolución")
        End Try
        conn.Close()
    End Sub

    '---Boton Actualizar

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        If (Button5.Text = "Actualizar") Then
            Button5.Text = "Guardar"
            TextBox1.Focus()
        Else
            conn.Open()
            Button5.Text = "Nuevo"
            Dim comando As New SqlCommand("", conn)

            Try
                Dim validacion As Integer
                validacion = TextBox1.Text

                If validacion >= 1 Then

                    comando.CommandType = CommandType.Text
                    comando.CommandText = "UPDATE devolciones SET id_prestamo = '" & TextBox2.Text & "' , Fecha_Devolucion = '" & TextBox3.Text & "'  where Id_devolucion = " & Integer.Parse(TextBox1.Text)

                    comando.ExecuteNonQuery()
                    MsgBox("Se actualizaron los datos de la devolución " & TextBox2.Text & " ", MsgBoxStyle.Information, "Datos Ingresados con Éxito")

                    TextBox1.Text = ""
                    TextBox2.Text = ""
                    TextBox3.Text = ""
                    TextBox1.Focus()
                Else
                    MsgBox("El código de Devolcuion no puede ser negativo. ", MsgBoxStyle.Critical, "Error, no se han ingresado el dato")
                End If

            Catch ex As Exception
                MsgBox("Por favor ingrese los datos que correspondan", MsgBoxStyle.Critical, "Error, no se han ingresado datos")
            End Try
            conn.Close()
        End If
    End Sub

    '---Botón Regresar

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Me.Close()
    End Sub

End Class


