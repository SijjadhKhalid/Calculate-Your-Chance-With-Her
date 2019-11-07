Public Class Form1
    Private errorMessage As String = "Incorrect Input for one or more Values"
    Private aY, aH, R, W, Ey, Eh, G, Cplus, Cminus, Sy, Sh, K As Integer

    Private Sub txtCplus_TextChanged(sender As Object, e As EventArgs) Handles txtCplus.TextChanged
        '  MessageBox.Show(txtCplus.Text)
        If Not txtCplus.Text() = "" Then
            Cplus = Integer.Parse(txtCplus.Text())
        End If
    End Sub

    Private Sub txtCminus_TextChanged(sender As Object, e As EventArgs) Handles txtCminus.TextChanged
        'MessageBox.Show(txtCminus.Text)
        If Not txtCminus.Text() = "" Then
            Cminus = Integer.Parse(txtCminus.Text())
        End If
    End Sub

    Private Sub txtG_TextChanged(sender As Object, e As EventArgs) Handles txtG.TextChanged
        'MessageBox.Show(txtG.Text)
        If Not txtG.Text() = "" Then
            G = Integer.Parse(txtG.Text())
        End If
    End Sub

    Private Sub txtSh_TextChanged(sender As Object, e As EventArgs) Handles txtSh.TextChanged
        'MessageBox.Show(txtSh.Text)
        If Not txtSh.Text() = "" Then
            Sh = Integer.Parse(txtSh.Text())
        End If
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles txtK.TextChanged
        ' MessageBox.Show(txtK.Text)
        If Not txtK.Text() = "" Then
            K = Integer.Parse(txtK.Text())
        End If
    End Sub

    Private Sub txtSy_TextChanged(sender As Object, e As EventArgs) Handles txtSy.TextChanged
        ' MessageBox.Show(txtSy.Text)
        If Not txtSy.Text() = "" Then
            Sy = Integer.Parse(txtSy.Text())
        End If
    End Sub



    Private Sub txtEh_TextChanged(sender As Object, e As EventArgs) Handles txtEh.TextChanged
        'MessageBox.Show(txtEh.Text)
        If Not txtEh.Text() = "" Then
            Eh = Integer.Parse(txtEh.Text())
        End If
    End Sub


    Private Sub txtEy_TextChanged(sender As Object, e As EventArgs) Handles txtEy.TextChanged
        ' MessageBox.Show(txtEy.Text)
        If Not txtEy.Text() = "" Then
            Ey = Integer.Parse(txtEy.Text())
        End If
    End Sub

    Private Sub txtW_TextChanged(sender As Object, e As EventArgs) Handles txtW.TextChanged
        ' MessageBox.Show(txtW.Text)
        If Not txtW.Text() = "" Then
            W = Integer.Parse(txtW.Text())
        End If
    End Sub

    Private Sub txtR_TextChanged(sender As Object, e As EventArgs) Handles txtR.TextChanged
        ' MessageBox.Show(txtR.Text)
        If Not txtR.Text() = "" Then
            R = Integer.Parse(txtR.Text())
        End If
    End Sub

    Private Sub txtAh_TextChanged(sender As Object, e As EventArgs) Handles txtAh.TextChanged
        ' MessageBox.Show(txtAh.Text)
        If Not txtAh.Text() = "" Then
            aH = Integer.Parse(txtAh.Text())
        End If
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub txtAy_TextChanged(sender As Object, e As EventArgs) Handles txtAy.TextChanged
        ' MessageBox.Show(txtAy.Text)
        If Not txtAy.Text() = "" Then
            aY = Integer.Parse(txtAy.Text())
        End If
    End Sub

    Private Sub submitBtn_Click(sender As Object, e As EventArgs) Handles btnSubmit.Click
        'Determines if true or false
        ' Checks from 1-10
        Dim aYCheck As Boolean = checkInputOneToTen(aY)
        Dim aHCheck As Boolean = checkInputOneToTen(aH)
        Dim WCheck As Boolean = checkInputOneToTen(W)
        Dim GCheck As Boolean = checkInputOneToTen(G)
        Dim CplusCheck As Boolean = checkInputOneToTen(Cplus)
        Dim CminusCheck As Boolean = checkInputOneToTen(Cplus)
        Dim KCheck As Boolean = checkInputOneToTen(K)

        'Checks for 0-10
        Dim RCheck As Boolean = checkInputOneToTen(R)

        'Checks for 12 & 16
        Dim EyCheck As Boolean = checkInput12And16(Ey)
        Dim EhCheck As Boolean = checkInput12And16(Eh)

        ' Checks for greater than 0
        Dim SyCheck As Boolean = checksForPositiveNum(Sy)
        Dim ShCheck As Boolean = checksForPositiveNum(Sh)


        ' Checks if numbers are between 1-10
        If aYCheck And aHCheck And RCheck And WCheck And EyCheck And EhCheck And GCheck And
            CplusCheck And CminusCheck And SyCheck And ShCheck And KCheck Then
            'N2 formats the numeric numbers with 2 decimal places
            MsgBox(executeFormula(aH, aY, W, R, Ey, Eh, G, Cplus, Cminus, Sy, Sh, K).ToString("n2"))
            'If numbers are not in the realm of 1-10 then "Else" is executed
        Else
            MsgBox(errorMessage)
        End If

        ' vbAbort

    End Sub

    Public Function executeFormula(ahInput As Integer, ayInput As Integer,
                                   rInput As Integer, wInput As Integer,
                                   eyInput As Integer, ehInput As Integer,
                                   GInput As Integer, CplusInput As Integer,
                                   CminusInput As Integer, SyInput As Integer,
                                   ShInput As Integer, KInput As Integer) As Double
        'Formula = WG(Ay)^2/(Ah)^3 + [(Ct-C) + (Sy-Sh/100) + (Ey-Eh)/K]
        Dim Formula As Double = ((wInput * GInput * (ayInput) ^ 2) / ((ahInput) ^ 3)) +
           (((CplusInput - CminusInput) + ((SyInput - ShInput) / 100) + (eyInput - ehInput)) / KInput) - rInput

        Return Formula
    End Function

    ' Function checks if number is between 1-10 and returns True or False
    Public Function checkInputOneToTen(input As Integer) As Boolean
        If input < 1 Or input > 10 Then
            System.Console.WriteLine(errorMessage)
            Return False
        End If
        Return True
    End Function

    Public Function checkInput12And16(input As Integer) As Boolean
        If input < 0 Or input > 16 Then
            System.Console.WriteLine(errorMessage)
            Return False
        End If
        Return True
    End Function

    Public Function checksForPositiveNum(input As Integer) As Boolean
        If input < 0 Then
            System.Console.WriteLine(errorMessage)
            Return False
        End If
        Return True
    End Function


    'Private Sub txtAttractive_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtAttractive.KeyPress
    '    'Stack Overflow
    '    If txtAttractive.Text.Length > 11 Then
    '        e.Handled = True
    '        Return
    '    End If
    '    If Asc(e.KeyChar) <> 8 Then
    '        If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
    '            e.Handled = True
    '        End If
    '    End If
    'End Sub
End Class

