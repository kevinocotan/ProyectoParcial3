Imports System.Data.SqlClient

Public Class Libros
    Public conn As New SqlConnection(My.Settings.conexion)
    Public lector As SqlDataReader
    Public ds As New DataSet

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

    Private Sub TextBox4_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox4.KeyPress

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

    '--el textbox 6 es de campo fecha y las plecas da un tipo de error. Para no complicarse mucho mejor se deja sin validaciones

    Private Sub TextBox7_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox7.KeyPress

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

    Private Sub TextBox8_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox8.KeyPress

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

    '--Boton Cargar

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        conn.Open()

        Dim comando As New SqlCommand("", conn)

        Try
            comando.CommandType = CommandType.Text
            comando.CommandText = "SELECT a.Id_Libro, a.Titulo, a.Id_Autor, b.Nombre, a.Id_Editorial, c.Nombre AS NomEditorial, a.Materia, a.Fecha_Lanzamiento, a.Edicion, a.Existencias
	                               FROM Libros a INNER JOIN Autor b ON a.Id_Autor = b.Id_Autor 
					                INNER JOIN Editorial c ON a.Id_Editorial = c.Id_Editorial"

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
            TextBox7.Text = lector(6)
            TextBox8.Text = lector(7)

            ListBox1.Items.Clear()

            Do
                ListBox1.Items.Add(lector("Id_Libro") & "   |   " & lector("Titulo") & "      |    " & lector("Id_Autor") & "        |    " & lector("Nombre") & "    |        " & lector("Id_Editorial") & "             |          " & lector("NomEditorial") & "       |         " & lector("Materia") & "    |   " & lector("Fecha_Lanzamiento") & "    |   " & lector("Edicion") & "  |  " & lector("Existencias"))

                TextBox1.Text = ""
                TextBox2.Text = ""
                TextBox3.Text = ""
                TextBox4.Text = ""
                TextBox5.Text = ""
                TextBox6.Text = ""
                TextBox7.Text = ""
                TextBox8.Text = ""
                TextBox1.Focus()

            Loop While (lector.Read())

        Else
            MsgBox("Para cargar los datos, primero ingreselos en el botón de ""Nuevo""  ", MsgBoxStyle.Critical, "Error, no se pudo cargar los datos")

        End If

        conn.Close()
    End Sub

    '---Boton Buscar

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        conn.Open()
        Dim comando As New SqlCommand("", conn)

        Try
            comando.CommandType = CommandType.Text
            comando.CommandText = "select * from Libros where id_libro=" & Integer.Parse(TextBox1.Text)

            lector = comando.ExecuteReader
            lector.Read()
            If (lector.HasRows()) Then
                TextBox1.Text = lector(0)
                TextBox2.Text = lector(1)
                TextBox3.Text = lector(2)
                TextBox4.Text = lector(3)
                TextBox5.Text = lector(4)
                TextBox6.Text = lector(5)
                TextBox7.Text = lector(6)
                TextBox8.Text = lector(7)

            Else
                MsgBox("El dato que se ha ingresado, no existe.", MsgBoxStyle.Critical, "Error")

            End If
        Catch ex As Exception
            MsgBox("Para buscar un dato, por favor ingrese el Codigo primero. ", MsgBoxStyle.Critical, "Error, no se han ingresado datos")
        End Try
        conn.Close()
    End Sub

    '---Boton Nuevo

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        If (Button3.Text = "Nuevo") Then
            Button3.Text = "Guardar"
            TextBox1.Text = ""
            TextBox2.Text = ""
            TextBox3.Text = ""
            TextBox4.Text = ""
            TextBox5.Text = ""
            TextBox6.Text = ""
            TextBox7.Text = ""
            TextBox8.Text = ""
            TextBox1.Focus()

        Else
            If (TextBox2.Text <> "") And (TextBox3.Text <> "") And (TextBox4.Text <> "") And (TextBox5.Text <> "") And (TextBox6.Text <> "") And (TextBox7.Text <> "") And (TextBox8.Text <> "") Then
                conn.Open()
                Button3.Text = "Nuevo"
                Dim comando As New SqlCommand("", conn)

                Try
                    Dim validacion As Integer
                    validacion = TextBox1.Text

                    If validacion >= 1 Then
                        comando.CommandType = CommandType.Text
                        comando.CommandText = "insert into Libros Values (" & Integer.Parse(TextBox1.Text) & " , '" & TextBox2.Text & "' ,  '" & TextBox3.Text & "' , '" & TextBox4.Text & "' ,  '" & TextBox5.Text & "' , '" & TextBox6.Text & "', '" & TextBox7.Text & "' , '" & TextBox8.Text & "')"

                        comando.ExecuteNonQuery()
                        MsgBox("Se ingreso el nuevo Libro: " & TextBox2.Text & " ", MsgBoxStyle.Information, "Datos Ingresados con Éxito")

                        TextBox1.Text = ""
                        TextBox2.Text = ""
                        TextBox3.Text = ""
                        TextBox4.Text = ""
                        TextBox5.Text = ""
                        TextBox6.Text = ""
                        TextBox7.Text = ""
                        TextBox8.Text = ""
                        TextBox1.Focus()

                    Else
                        MsgBox("El código de Libros no puede ser negativo. ", MsgBoxStyle.Critical, "Error, no se han ingresado el dato")
                    End If

                Catch ex As Exception
                    MsgBox("Por favor ingrese los datos que correspondan o revise si el Código del Libro, Autor o de Editorial existe ", MsgBoxStyle.Critical, "Error al ingresar el Libro")
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

                MsgBox("El código de Libro no puede ser negativo. ", MsgBoxStyle.Critical, "Error, no se han ingresado el dato")
            Else
                Try

                    eliminar.CommandType = CommandType.Text
                    eliminar.CommandText = "DELETE from Libros where Id_Libro = " & Integer.Parse(TextBox1.Text)

                    eliminar.ExecuteNonQuery()
                    MsgBox("Se eliminó el Libro con código: " & TextBox1.Text, MsgBoxStyle.Information, "Datos Eliminados con Éxito")
                Catch ex As Exception
                    MsgBox("Asegurese que este campo no está referenciado en otro Formulario.", MsgBoxStyle.Critical, "Error, al intentar borrar los datos")
                End Try
            End If
        Catch ex As Exception
            MsgBox("Este campo es obligatorio y debe de ser un número entero. ", MsgBoxStyle.Critical, "Error, no se pudo eliminar el Libro")
        End Try

        conn.Close()
    End Sub

    '--Boton Actualizar

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
                    comando.CommandText = "update Libros SET Titulo = '" & TextBox2.Text & "' , Id_Autor = '" & TextBox3.Text & "' , Id_Editorial = '" & TextBox4.Text & "', Materia = '" & TextBox5.Text & "' , Fecha_Lanzamiento = '" & TextBox6.Text & "' , Edicion = '" & TextBox7.Text & "'  , Existencias = '" & TextBox8.Text & "' where Id_libro = " & Integer.Parse(TextBox1.Text)

                    comando.ExecuteNonQuery()
                    MsgBox("Se actualizó el Libro con el titulo: " & TextBox2.Text & " ", MsgBoxStyle.Information, "Datos Ingresados con Éxito")

                    TextBox1.Text = ""
                    TextBox2.Text = ""
                    TextBox3.Text = ""
                    TextBox4.Text = ""
                    TextBox5.Text = ""
                    TextBox6.Text = ""
                    TextBox7.Text = ""
                    TextBox8.Text = ""
                    TextBox1.Focus()
                Else
                    MsgBox("El código de Libro no puede ser negativo. ", MsgBoxStyle.Critical, "Error, no se han ingresado el dato")
                End If

            Catch ex As Exception
                MsgBox("Por favor ingrese los datos que correspondan. Recurde que el código no puede quedar vacio y debe de existir.", MsgBoxStyle.Critical, "Error, no se han ingresado datos")
            End Try
            conn.Close()
        End If
    End Sub

    '---Boton Regresar

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Me.Close()
    End Sub

    ' Sub llenarComboBox1(ByVal ComboBox1 As ComboBox)
    'Dim comando As New SqlCommand("", conn)
    'Try
    '    comando.CommandType = CommandType.Text
    '     comando.CommandText = "select Nombre from Autor"
    '    lector = comando.ExecuteReader
    'While lector.Read()
    ' ComboBox1.Items.Add(lector.Item("Nombre"))
    '   End While
    ' Catch ex As Exception
    '   MsgBox(ex.Message)
    ' End Try
    '  conn.Close()
    ' End Sub

End Class