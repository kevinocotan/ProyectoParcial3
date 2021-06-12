Imports System.Data.SqlClient

Public Class Usuarios
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

    Private Sub TextBox3_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox3.KeyPress
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

    Private Sub TextBox4_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox4.KeyPress
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

    Private Sub TextBox5_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox5.KeyPress
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

    Private Sub TextBox6_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox6.KeyPress
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

    'Botón 1 = Cargar

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        conn.Open()

        Dim comando As New SqlCommand("", conn)

        Try
            comando.CommandType = CommandType.Text
            comando.CommandText = "SELECT * FROM Usuarios"

            lector = comando.ExecuteReader
            lector.Read()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        If (lector.HasRows()) Then
            TextBox1.Text = lector(0)
            TextBox2.Text = lector(1)
            TextBox3.Text = lector(2)
            TextBox4.Text = lector(3)
            TextBox5.Text = lector(4)
            TextBox6.Text = lector(5)
            ListBox1.Items.Clear()

            Do
                ListBox1.Items.Add(lector("Id_Usuario") & "      |   " & lector("Nombre") & "    |    " & lector("Apellido") & "    |    " & lector("Materia") & "    |    " & lector("Direccion") & "    |    " & lector("Telefono"))
                TextBox1.Text = ""
                TextBox2.Text = ""
                TextBox3.Text = ""
                TextBox4.Text = ""
                TextBox5.Text = ""
                TextBox6.Text = ""
                TextBox1.Focus()

            Loop While (lector.Read())
        Else
            MsgBox("Para cargar los datos, primero ingreselos en el botón de ""Nuevo""  ", MsgBoxStyle.Critical, "Error, no se pudo cargar los datos")
        End If

        conn.Close()
    End Sub

    'Botón 2 = Buscar

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        conn.Open()
        Dim comando As New SqlCommand("", conn)

        Try
            comando.CommandType = CommandType.Text
            comando.CommandText = "SELECT * FROM Usuarios WHERE id_Usuario=" & Integer.Parse(TextBox1.Text)

            lector = comando.ExecuteReader
            lector.Read()
            If (lector.HasRows()) Then
                TextBox1.Text = lector(0)
                TextBox2.Text = lector(1)
                TextBox3.Text = lector(2)
                TextBox4.Text = lector(3)
                TextBox5.Text = lector(4)
                TextBox6.Text = lector(5)

            Else
                MsgBox("El dato que se ha ingresado, no existe. ", MsgBoxStyle.Critical, "Error")
            End If
        Catch ex As Exception
            'MsgBox(ex.Message)
            MsgBox("Para buscar un dato, por favor ingrese el Codigo primero. ", MsgBoxStyle.Critical, "Error, no se han encontrado datos")
        End Try
        conn.Close()
    End Sub

    'Botón 3 = Nuevo

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If (Button3.Text = "Nuevo") Then
            Button3.Text = "Guardar"
            TextBox1.Text = ""
            TextBox2.Text = ""
            TextBox3.Text = ""
            TextBox4.Text = ""
            TextBox5.Text = ""
            TextBox6.Text = ""
            TextBox1.Focus()

        Else
            If (TextBox2.Text <> "") And (TextBox3.Text <> "") And (TextBox4.Text <> "") And (TextBox5.Text <> "") And (TextBox6.Text <> "") Then
                conn.Open()
                Button3.Text = "Nuevo"
                Dim comando As New SqlCommand("", conn)

                Try
                    Dim validacion As Integer
                    validacion = TextBox1.Text

                    If validacion >= 1 Then
                        comando.CommandType = CommandType.Text
                        comando.CommandText = "INSERT INTO Usuarios VALUES (" & Integer.Parse(TextBox1.Text) & " , '" & TextBox2.Text & "','" & TextBox3.Text & "','" & TextBox4.Text & "','" & TextBox5.Text & "','" & TextBox6.Text & "')"

                        comando.ExecuteNonQuery()
                        MsgBox("Se ingreso el nuevo Usuario: " & TextBox2.Text & " " & TextBox3.Text & " ", MsgBoxStyle.Information, "Datos Ingresados con Éxito")

                        TextBox1.Text = ""
                        TextBox2.Text = ""
                        TextBox3.Text = ""
                        TextBox4.Text = ""
                        TextBox5.Text = ""
                        TextBox6.Text = ""
                        TextBox1.Focus()
                    Else
                        MsgBox("El código de Usuario no puede ser negativo. ", MsgBoxStyle.Critical, "Error, no se han ingresado el dato")
                    End If

                Catch ex As Exception
                    MsgBox("Por favor ingrese los datos que correspondan. O revise si el Código de Usuario ya existe. ", MsgBoxStyle.Critical, "Error al Ingresar el Usuario")
                End Try
            Else
                MsgBox("Hay campos obligatorios que se encuentran vacios. Revise e intente de nuevo por favor. ", MsgBoxStyle.Critical, "Error al ingresar los datos")

            End If
        End If
        conn.Close()
    End Sub

    'Botón 4 = Eliminar

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        conn.Open()
        Dim eliminar As New SqlCommand("", conn)


        Try
            Dim validacion As Integer
            validacion = TextBox1.Text

            If validacion < 1 Then
                MsgBox("El código de Usuario no puede ser negativo. ", MsgBoxStyle.Critical, "Error, no se han ingresado el dato")
            Else
                Try
                    eliminar.CommandType = CommandType.Text
                    eliminar.CommandText = "DELETE FROM Usuarios WHERE Id_Usuario = " & Integer.Parse(TextBox1.Text)

                    eliminar.ExecuteNonQuery()
                    MsgBox("Se eliminó el Usuario con código:  " & TextBox1.Text, MsgBoxStyle.Information, "Datos Eliminados con Éxito")
                Catch ex As Exception
                    MsgBox("Asegurese que este campo no está referenciado en otro Formulario.", MsgBoxStyle.Critical, "Error, al intentar borrar los datos")
                End Try
            End If
        Catch ex As Exception
            MsgBox("Este campo es obligatorio y debe de ser un número entero. ", MsgBoxStyle.Critical, "Error, no se pudo eliminar el Usuario")
        End Try

        conn.Close()
    End Sub

    'Botón 5 = Actualizar

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        If (Button5.Text = "Actualizar") Then
            Button5.Text = "Guardar"
            TextBox1.Focus()
        Else
            conn.Open()
            Button5.Text = "Actualizar"
            Dim comando As New SqlCommand("", conn)


            Try
                Dim validacion As Integer
                validacion = TextBox1.Text

                If validacion >= 1 Then
                    comando.CommandType = CommandType.Text
                    comando.CommandText = "UPDATE Usuarios SET Nombre = '" & TextBox2.Text & "', Apellido = '" & TextBox3.Text & "', Materia = '" & TextBox4.Text & "', Direccion = '" & TextBox5.Text & "', Telefono = '" & TextBox6.Text & "' WHERE Id_Usuario = " & Integer.Parse(TextBox1.Text)

                    comando.ExecuteNonQuery()
                    MsgBox("Se actualizaron los datos del Usuario " & " " & TextBox2.Text & " con código: " & TextBox1.Text, MsgBoxStyle.Information, "Datos Ingresados con Éxito")

                    TextBox1.Text = ""
                    TextBox2.Text = ""
                    TextBox3.Text = ""
                    TextBox4.Text = ""
                    TextBox5.Text = ""
                    TextBox6.Text = ""
                    TextBox1.Focus()

                Else
                    MsgBox("El código de Usuario no puede ser negativo. ", MsgBoxStyle.Critical, "Error, no se han ingresado el dato")
                End If

            Catch ex As Exception
                MsgBox("Por favor ingrese los datos que correspondan. Recurde que el código no puede quedar vacio y debe de existir.", MsgBoxStyle.Critical, "Error, no se han ingresado datos")
            End Try
            conn.Close()
        End If
    End Sub

    'Botón 6 = Regresar

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Me.Close()
    End Sub

End Class
