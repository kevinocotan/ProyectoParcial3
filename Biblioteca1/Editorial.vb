Imports System.Data.SqlClient

Public Class Editorial
    Public conn As New SqlConnection(My.Settings.conexion)
    Public lector As SqlDataReader

    '--validacion de campos

    Private Sub TextBox2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox2.KeyPress

        If Char.IsLetter(e.KeyChar) Then
            e.Handled = False
        ElseIf Char.IsControl(e.KeyChar) Then
            e.Handled = False
        ElseIf Char.IsSeparator(e.KeyChar) Then
            e.Handled = False
        Else
            e.Handled = True
            MsgBox("Este campo es de solo texto, no puede ingresar otro tipo de datos. Intente de nuevo por favor. ", MsgBoxStyle.Critical, "Error al ingresar el dato.")
        End If

    End Sub

    '---Botón Cargar

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        conn.Open()

        Dim comando As New SqlCommand("", conn)

        'comando.CommandType = CommandType.Text
        'comando.CommandText = "select * from Editorial
        Try
            comando.CommandType = CommandType.Text
            comando.CommandText = "select * from Editorial"
            lector = comando.ExecuteReader
            lector.Read()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        If (lector.HasRows()) Then
            TextBox1.Text = lector(0)
            TextBox2.Text = lector(1)
            ListBox1.Items.Clear()

            Do
                ListBox1.Items.Add(lector("Id_Editorial") & "        |     " & lector("Nombre"))

                TextBox1.Text = ""
                TextBox2.Text = ""
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

        'comando.CommandType = CommandType.Text
        'comando.CommandText = "Select * from Editorial where id_Editorial=" & Integer.Parse(TextBox1.Text)
        'verfificar cuando no esta lo que se busca

        Try
            comando.CommandType = CommandType.Text
            comando.CommandText = "Select * from Editorial where id_Editorial=" & Integer.Parse(TextBox1.Text)

            lector = comando.ExecuteReader
            lector.Read()
            If (lector.HasRows()) Then
                TextBox1.Text = lector(0)
                TextBox2.Text = lector(1)

            Else
                MsgBox("El dato que se ha ingresado, no existe ", MsgBoxStyle.Critical, "Error")

            End If
        Catch ex As Exception
            'MsgBox(ex.Message)
            MsgBox("Para buscar un dato, por favor ingrese el Codigo primero. ", MsgBoxStyle.Critical, "Error, no se han encontrado datos")
        End Try
        conn.Close()
    End Sub

    '---Botón Nuevo

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        If (Button3.Text = "Nuevo") Then
            Button3.Text = "Guardar"
            TextBox1.Text = ""
            TextBox2.Text = ""
            TextBox1.Focus()

        Else
            If (TextBox2.Text <> "") Then

                conn.Open()
                Button3.Text = "Nuevo"
                Dim comando As New SqlCommand("", conn)
                'comando.CommandType = CommandType.Text
                'comando.CommandText = "insert into Editorial Values (" & Integer.Parse(TextBox1.Text) & " , '" & TextBox2.Text & "')"

                Try
                    Dim validacion As Integer
                    validacion = TextBox1.Text

                    If validacion >= 1 Then
                        comando.CommandType = CommandType.Text
                        comando.CommandText = "insert into Editorial Values (" & Integer.Parse(TextBox1.Text) & " , '" & TextBox2.Text & "')"

                        comando.ExecuteNonQuery()
                        MsgBox("Se ingreso la nueva Editorial: " & TextBox2.Text & " ", MsgBoxStyle.Information, "Datos Ingresados con Éxito")

                        TextBox1.Text = ""
                        TextBox2.Text = ""
                        TextBox1.Focus()

                    Else
                        MsgBox("El código de Editorial no puede ser negativo. ", MsgBoxStyle.Critical, "Error, no se han ingresado el dato")
                    End If

                Catch ex As Exception
                    'MsgBox(ex.Message)
                    MsgBox("Por favor ingrese los datos que correspondan. Recurde que el Código Editorial no debe repetirse y debe de ser un numero entero.", MsgBoxStyle.Critical, "Error al ingresar la Editorial")
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

                MsgBox("El código de Editorial no puede ser negativo. ", MsgBoxStyle.Critical, "Error, no se han ingresado el dato")
            Else

                Try
                    eliminar.CommandType = CommandType.Text
                    eliminar.CommandText = "DELETE from editorial where id_editorial = " & Integer.Parse(TextBox1.Text)

                    eliminar.ExecuteNonQuery()
                    MsgBox("Se eliminó la Editorial con código: " & TextBox1.Text, MsgBoxStyle.Information, "Datos Eliminados con Éxito")

                Catch ex As Exception
                    'MsgBox(ex.Message)
                    'si funciona que solo elimine los que existan puedo agregar al mensaje que "tiene que existir"
                    MsgBox("Asegurese que este campo no está referenciado en otro Formulario.", MsgBoxStyle.Critical, "Error, al intentar borrar los datos")
                End Try
            End If

        Catch ex As Exception
            MsgBox("Este campo es obligatorio y debe de ser un número entero. ", MsgBoxStyle.Critical, "Error, no se pudo eliminar la Editorial")
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
            Button5.Text = "Actualizar"
            Dim comando As New SqlCommand("", conn)
            'comando.CommandType = CommandType.Text
            'comando.CommandText = "update Editorial set Nombre = '" & TextBox2.Text & "' where Id_Editorial = " & Integer.Parse(TextBox1.Text)

            Try

                Dim validacion As Integer
                validacion = TextBox1.Text

                If validacion >= 1 Then

                    comando.CommandType = CommandType.Text
                    comando.CommandText = "UPDATE Editorial SET Nombre = '" & TextBox2.Text & "' where Id_Editorial = " & Integer.Parse(TextBox1.Text)

                    comando.ExecuteNonQuery()
                    MsgBox("Se actualizó la Editorial con código: " & TextBox1.Text & " a: " & TextBox2.Text & " ", MsgBoxStyle.Information, "Datos Ingresados con Éxito")

                    TextBox1.Text = ""
                    TextBox2.Text = ""
                    TextBox1.Focus()

                Else
                    MsgBox("El código de Editorial no puede ser negativo. ", MsgBoxStyle.Critical, "Error, no se han ingresado el dato")
                End If

            Catch ex As Exception
                'MsgBox(ex.Message)
                MsgBox("Por favor ingrese los datos que correspondan. Recurde que el Codigo debe de existir en Editorial y este no puede quedar vacio.", MsgBoxStyle.Critical, "Error, no se han podido actualizar datos")
            End Try
            conn.Close()
        End If
    End Sub

    '---Botón Regresar

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Me.Close()
    End Sub

End Class