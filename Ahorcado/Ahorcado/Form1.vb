Public Class Form1
    Dim palabras(14) As String
    Dim palabra As String
    Dim cantidad As Integer
    Dim i As Integer
    Dim intento As Integer = 0
    Dim verificar As Boolean = True
    Dim intentosExitosos As Integer = 0

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        MsgBox("Bienvenido!!!")
        palabras(0) = "hola"
        palabras(1) = "temblor"
        palabras(2) = "juniorexpress"
        palabras(3) = "murcielago"
        palabras(4) = "escombros"
        palabras(5) = "kinalquinto"
        palabras(6) = "informatica"
        palabras(7) = "planta"
        palabras(8) = "mario"
        palabras(9) = "perro"
        palabras(10) = "java"
        palabras(11) = "agenda"
        palabras(12) = "panes"
        palabras(13) = "marco"
        palabras(14) = "esperanza"


        Randomize() 'Para que no agarre un patron
        palabra = palabras(15 * Rnd())
        'MsgBox(palabra)
        Enabled = True 'True para activar y False para desactivar
        If verificar = True Then
            Button2.Enabled = True
        End If

        Button2.PerformClick()

        Button2.Visible = False

    End Sub


    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim etiqueta As Label
        Dim posetiqueta As Integer = 328 'Posicion en x

        cantidad = palabra.Length

        For i = 1 To cantidad Step 1
            etiqueta = New Label()
            etiqueta.Text = "_____"
            etiqueta.Width = 50
            etiqueta.Location = New Point(posetiqueta, 250)
            Me.Controls.Add(etiqueta)
            posetiqueta += 55
        Next

        If verificar = True Then
            txtLetra.Enabled = True
            Button3.Enabled = True

        End If

    End Sub

    Private Sub NewToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NewToolStripMenuItem.Click
        Application.Restart()

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim dibujar As Graphics = PictureBox2.CreateGraphics
        Dim tletra As New Font("Harlow Solid Italic", 16)
        Dim inicio As Integer = 328
        Dim espacio As Integer = 0
        Dim letra As Char
        letra = LCase(txtLetra.Text)
        Dim letraExistente As Boolean = False
        Dim comprobar As Boolean = False
        Dim prueba As Boolean = False
        Dim confirmar As Integer



        For i = 1 To Len(palabra) Step 1
            If letra.Equals(GetChar(palabra, i)) Then
                Me.CreateGraphics.DrawString(letra, tletra, Brushes.Red, (inicio + espacio), 225)
                comprobar = True
                letraExistente = True
            End If
            espacio += 55
        Next


        If (ListBox1.Items.Contains(LCase(txtLetra.Text))) Then
            MsgBox("La tecla ya se utilizó")
            comprobar = True
            prueba = True
        Else
            ListBox1.Items.Add(LCase(txtLetra.Text))
        End If


        If (comprobar = False) Then
            MsgBox("La letra no se encuentra en la palabra")
            intento += 1
            letraExistente = False
        End If


        If letraExistente = True Then            '

            If prueba = True Then
                confirmar = True
            End If

            If prueba = False Then



                For i = 1 To Len(palabra) Step 1
                    If letra.Equals(GetChar(palabra, i)) Then
                        intentosExitosos += 1
                        '  MsgBox(intentosExitosos.ToString)

                    End If
                Next


                If (intentosExitosos.Equals(Len(palabra))) Then
                    MsgBox("Ganó", MsgBoxStyle.Information, "Felicidades")
                    confirmar = MsgBox("Desea volver a jugar ", vbYesNo + vbExclamation + vbDefaultButton2, "Gano")
                    If confirmar = MsgBoxResult.Yes Then
                        Application.Restart()
                    Else
                        Me.Close()

                    End If
                End If







            End If

        End If





        If letraExistente = False Then
            Select Case intento
                Case 1
                    dibujar.DrawEllipse(Pens.Black, 10, 5, 50, 50)
                Case 2
                    dibujar.DrawLine(Pens.Black, 35, 55, 35, 125)
                Case 3
                    dibujar.DrawLine(Pens.Black, 35, 80, 10, 60)
                Case 4
                    dibujar.DrawLine(Pens.Black, 35, 80, 55, 60)
                Case 5
                    dibujar.DrawLine(Pens.Black, 35, 125, 10, 145)
                Case 6
                    dibujar.DrawLine(Pens.Black, 35, 125, 55, 145)
                Case 7
                    dibujar.DrawLine(Pens.Black, 20, 60, 45, 60)
                    MsgBox("Perdiste :(")

                    MsgBox("La palabra es " + palabra)

                    confirmar = MsgBox("Desea volver a  jugar ", vbYesNo + vbExclamation + vbDefaultButton2, "Perdió")
                    If confirmar = MsgBoxResult.Yes Then
                        Application.Restart()
                    Else
                        Me.Close()
                    End If
            End Select
        End If
    End Sub

    Private Sub ArchivoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ArchivoToolStripMenuItem.Click
        Dim archivo, archivo2 As String
        OpenFileDialog1.ShowDialog()
        archivo = OpenFileDialog1.FileName
        archivo2 = My.Computer.FileSystem.ReadAllText(archivo)
    End Sub

    Private Sub txtLetra_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtLetra.KeyPress

        Dim tecla As Char
        Dim letra As Boolean

        If Char.IsDigit(e.KeyChar) Then
            e.Handled = True
        ElseIf Char.IsControl(e.KeyChar) Then
            e.Handled = False
        ElseIf Char.IsSymbol(e.KeyChar) Then
            e.Handled = True
        ElseIf Char.IsSeparator(e.KeyChar) Then
            e.Handled = True
        ElseIf Char.IsWhiteSpace(e.KeyChar) Then
            e.Handled = True
        Else
            e.Handled = False
            tecla = e.KeyChar.ToString

            If txtLetra.Text.Contains(tecla) Then
                MsgBox("la tecla ya se ingreso ")
                letra = True
                e.Handled = True
            End If
        End If



    End Sub

    Private Sub SalirToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SalirToolStripMenuItem.Click
        Me.Close()

    End Sub
End Class
