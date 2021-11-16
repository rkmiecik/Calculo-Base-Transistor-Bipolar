Public Class Form1
    Dim ganho As Decimal
    Dim v1 As Decimal
    Dim v2 As Decimal
    Dim resl As Decimal
    Dim corrl As Decimal
    Dim corrt As Decimal
    Dim reb As Decimal
    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        PictureBox1.Image = My.Resources.NPN
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        PictureBox1.Image = My.Resources.PNP
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim alturamax As Integer = My.Computer.Screen.Bounds.Height.ToString
        Dim larguramax As Integer = My.Computer.Screen.Bounds.Width.ToString
        Me.Size = New System.Drawing.Size((larguramax / 1.2), (alturamax / 1.2))
        TableLayoutPanel1.Font = New Font(TableLayoutPanel1.Font.Name, larguramax / 200)
        Me.Location = New Point((Screen.PrimaryScreen.WorkingArea.Width - Me.Width) / 2, (alturamax / 20))
        '''''
        TextBox1.Font = New Font(TextBox1.Font.Name, larguramax / 120)
        TextBox2.Font = New Font(TextBox2.Font.Name, larguramax / 120)
        TextBox3.Font = New Font(TextBox3.Font.Name, larguramax / 120)
        TextBox4.Font = New Font(TextBox4.Font.Name, larguramax / 120)
        Button1.Font = New Font(Button1.Font.Name, larguramax / 100)
    End Sub
    Private Sub textbox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress
        Dim KeyAscii As Short = CShort(Asc(e.KeyChar))

        KeyAscii = CShort(numerosspermitidos(KeyAscii))

        If KeyAscii = 0 Then

            e.Handled = True
        End If

    End Sub
    Private Sub textbox2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox2.KeyPress
        Dim KeyAscii As Short = CShort(Asc(e.KeyChar))

        KeyAscii = CShort(numerosspermitidos(KeyAscii))

        If KeyAscii = 0 Then

            e.Handled = True
        End If

    End Sub
    Private Sub textbox3_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox3.KeyPress
        Dim KeyAscii As Short = CShort(Asc(e.KeyChar))

        KeyAscii = CShort(numerosspermitidos(KeyAscii))

        If KeyAscii = 0 Then

            e.Handled = True
        End If

    End Sub
    Private Sub textbox4_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox4.KeyPress
        Dim KeyAscii As Short = CShort(Asc(e.KeyChar))

        KeyAscii = CShort(numerosspermitidos(KeyAscii))

        If KeyAscii = 0 Then

            e.Handled = True
        End If

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Then
                MessageBox.Show("Preencha todos os campos!", "Campos vazios!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Else
                TextBox1.Text = TextBox1.Text.Replace(".", ",")
                Dim txtbx1 As String = TextBox1.Text
                Dim numtxt1ponto As Object = Split(txtbx1, ".")
                TextBox2.Text = TextBox2.Text.Replace(".", ",")
                Dim txtbx2 As String = TextBox2.Text
                Dim numtxt2ponto As Object = Split(txtbx2, ".")
                TextBox3.Text = TextBox3.Text.Replace(".", ",")
                Dim txtbx3 As String = TextBox3.Text
                Dim numtxt3ponto As Object = Split(txtbx3, ".")
                TextBox4.Text = TextBox4.Text.Replace(".", ",")
                Dim txtbx4 As String = TextBox4.Text
                Dim numtxt4ponto As Object = Split(txtbx4, ".")

                If Convert.ToDecimal(TextBox1.Text) = 0 Or (UBound(numtxt1ponto) > 1) Or TextBox1.Text.StartsWith(",") Or TextBox1.Text.StartsWith(".") Or TextBox1.Text.EndsWith(".") Or TextBox1.Text.EndsWith(",") Then
                    MessageBox.Show("hFE inválido!", "hFE inválido!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                ElseIf Convert.ToDecimal(TextBox2.Text) = 0 Or (UBound(numtxt2ponto) > 1) Or TextBox2.Text.StartsWith(",") Or TextBox2.Text.StartsWith(".") Or TextBox2.Text.EndsWith(".") Or TextBox2.Text.EndsWith(",") Then
                    MessageBox.Show("Tensão V1 inválida!", "Tensão V1 inválida!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                ElseIf Convert.ToDecimal(TextBox3.Text) = 0 Or (UBound(numtxt3ponto) > 1) Or TextBox3.Text.StartsWith(",") Or TextBox3.Text.StartsWith(".") Or TextBox3.Text.EndsWith(".") Or TextBox3.Text.EndsWith(",") Then
                    MessageBox.Show("Tensão V2 inválida!", "Tensão V2 inválida!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                ElseIf Convert.ToDecimal(TextBox4.Text) = 0 Or (UBound(numtxt4ponto) > 1) Or TextBox4.Text.StartsWith(",") Or TextBox4.Text.StartsWith(".") Or TextBox4.Text.EndsWith(".") Or TextBox4.Text.EndsWith(",") Then
                    MessageBox.Show("Resistência da carga inválida!", "Resistência da carga inválida!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Else
                    ganho = Convert.ToDecimal(TextBox1.Text)
                    v1 = Convert.ToDecimal(TextBox2.Text)
                    v2 = Convert.ToDecimal(TextBox3.Text)
                    resl = Convert.ToDecimal(TextBox4.Text)
                    'cálculo da corrente da carga
                    corrl = v1 / resl
                    'cálculo da corrente na base do transistor
                    corrt = corrl / ganho
                    'cálculo do resistor de base do transistor
                    reb = v2 / corrt


                    MessageBox.Show("Valores obtidos:" & Environment.NewLine & Environment.NewLine & "CALCULANDO A CORRENTE DA CARGA: " & Environment.NewLine & "Corrente na carga= V1 / Resistência da Carga" & Environment.NewLine & "Corrente na carga= " & v1 & " / " & resl & Environment.NewLine & "Corrente na carga= " & corrl & "A" & Environment.NewLine & Environment.NewLine & "CALCULANDO A CORRENTE NA BASE DO TRANSISTOR:" & Environment.NewLine & "Corrente de base= Corrente na carga / Ganho" & Environment.NewLine & "Corrente de base= " & corrl & " / " & ganho & Environment.NewLine & "Corrente de base= " & corrt & "A" & Environment.NewLine & Environment.NewLine & "CALCULANDO O RESISTOR DE BASE DO TRANSISTOR:" & Environment.NewLine & "Resistor de base= V2 / Corrente de base" & Environment.NewLine & "Resistor de base= " & v2 & " / " & corrt & Environment.NewLine & "RESISTOR DE BASE= " & reb & " OHMS", "Valores obtidos", MessageBoxButtons.OK, MessageBoxIcon.Information)

                End If
            End If


        Catch ex As Exception

        End Try


    End Sub
End Class
