Imports ggcAppDriver
Imports MySql.Data.MySqlClient

Public Class frmPresident
    Private p_sEmployID As String
    Private p_sTownIDxx As String
    Private p_sCandIDxx As String

    WriteOnly Property EmployeeID As String
        Set(ByVal Value As String)
            p_sEmployID = Value
        End Set
    End Property

    WriteOnly Property TownIDxx As String
        Set(ByVal Value As String)
            p_sTownIDxx = Value
        End Set
    End Property

    Private Sub Form1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        Select Case e.KeyCode
            Case Keys.Return
                If p_sCandIDxx = "" Then
                    MsgBox("Invalid Entry for President!!!" & vbCrLf & _
                           "Please select using KeyPad!!!", vbCritical, "WARNING")
                Else
                    Dim lsSQL As String

                    lsSQL = "INSERT INTO xxxBallot SET" & _
                                "  sBallotID = " & strParm(GetNextCode("xxxBallot", "sBallotID", True, p_oApp.Connection, True, p_oApp.BranchCode)) & _
                                ", sEmployID = " & strParm(p_sEmployID) & _
                                ", sCandIDxx = " & strParm(p_sCandIDxx) & _
                                ", sTownIDxx = " & strParm(p_sTownIDxx)

                    p_oApp.Execute(lsSQL, "xxxBallot", p_oApp.BranchCode)
                    Me.Close()
                End If
            Case Keys.NumPad1
                optCand1.Checked = True
                optCand2.ForeColor = Color.Black
                optCand3.ForeColor = Color.Black
                optCand4.ForeColor = Color.Black
                optCand5.ForeColor = Color.Black
                optCand6.ForeColor = Color.Black
                Call optCand1_CheckedChanged(sender, e)
            Case Keys.NumPad2
                optCand2.Checked = True
                optCand1.ForeColor = Color.Black
                optCand3.ForeColor = Color.Black
                optCand4.ForeColor = Color.Black
                optCand5.ForeColor = Color.Black
                optCand6.ForeColor = Color.Black
                Call optCand2_CheckedChanged(sender, e)
            Case Keys.NumPad3
                optCand3.Checked = True
                optCand1.ForeColor = Color.Black
                optCand2.ForeColor = Color.Black
                optCand4.ForeColor = Color.Black
                optCand5.ForeColor = Color.Black
                optCand6.ForeColor = Color.Black
                Call optCand3_CheckedChanged(sender, e)
            Case Keys.NumPad4
                optCand4.Checked = True
                optCand1.ForeColor = Color.Black
                optCand2.ForeColor = Color.Black
                optCand3.ForeColor = Color.Black
                optCand5.ForeColor = Color.Black
                optCand6.ForeColor = Color.Black
                Call optCand4_CheckedChanged(sender, e)
            Case Keys.NumPad5
                optCand5.Checked = True
                optCand1.ForeColor = Color.Black
                optCand2.ForeColor = Color.Black
                optCand3.ForeColor = Color.Black
                optCand4.ForeColor = Color.Black
                optCand6.ForeColor = Color.Black
                Call optCand5_CheckedChanged(sender, e)
            Case Keys.NumPad6
                optCand6.Checked = True
                optCand1.ForeColor = Color.Black
                optCand2.ForeColor = Color.Black
                optCand3.ForeColor = Color.Black
                optCand4.ForeColor = Color.Black
                optCand5.ForeColor = Color.Black
                Call optCand6_CheckedChanged(sender, e)
        End Select
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim loDta As DataTable
        Dim lnCtr As Integer
        Dim lsSQL As String

        Windows.Forms.Cursor.Hide()

        loDta = New DataTable
        lsSQL = "SELECT" & _
                    "*" & _
                    " FROM xxxCandidates " & _
                    " WHERE sPositIDx = '03' " & _
                    " ORDER BY sCandIDxx"

        loDta = p_oApp.ExecuteQuery(lsSQL)
        Label10.Text = "PANGULO NG PILIPINAS" & "?"

        lnCtr = 0
        Do While lnCtr < loDta.Rows.Count
            Select Case lnCtr
                Case 0
                    optCand1.Text = UCase(loDta(lnCtr)("sCandName"))
                    optCand1.Tag = UCase(loDta(lnCtr)("sCandIDxx"))
                Case 1
                    optCand2.Text = UCase(loDta(lnCtr)("sCandName"))
                    optCand2.Tag = UCase(loDta(lnCtr)("sCandIDxx"))
                Case 2
                    optCand3.Text = UCase(loDta(lnCtr)("sCandName"))
                    optCand3.Tag = UCase(loDta(lnCtr)("sCandIDxx"))
                Case 3
                    optCand4.Text = UCase(loDta(lnCtr)("sCandName"))
                    optCand4.Tag = UCase(loDta(lnCtr)("sCandIDxx"))
                Case 4
                    optCand5.Text = UCase(loDta(lnCtr)("sCandName"))
                    optCand5.Tag = UCase(loDta(lnCtr)("sCandIDxx"))
                Case 5
                    optCand6.Text = UCase(loDta(lnCtr)("sCandName"))
                    optCand6.Tag = UCase(loDta(lnCtr)("sCandIDxx"))
            End Select
            lnCtr = lnCtr + 1
        Loop

        optCand7.Checked = True
    End Sub

    Private Sub optCand1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optCand1.CheckedChanged
        If optCand1.Checked = True Then
            optCand1.ForeColor = Color.Red
            p_sCandIDxx = optCand1.Tag
        End If
    End Sub

    Private Sub optCand1_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles optCand1.Validated
        optCand1.ForeColor = Color.Black
    End Sub

    Private Sub optCand2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optCand1.CheckedChanged
        If optCand2.Checked = True Then
            optCand2.ForeColor = Color.Red
            p_sCandIDxx = optCand2.Tag
        End If
    End Sub

    Private Sub optCand2_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles optCand1.Validated
        optCand2.ForeColor = Color.Black
    End Sub
    Private Sub optCand3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optCand1.CheckedChanged
        If optCand3.Checked = True Then
            optCand3.ForeColor = Color.Red
            p_sCandIDxx = optCand3.Tag
        End If
    End Sub

    Private Sub optCand3_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles optCand1.Validated
        optCand3.ForeColor = Color.Black
    End Sub
    Private Sub optCand4_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optCand1.CheckedChanged
        If optCand4.Checked = True Then
            optCand4.ForeColor = Color.Red
            p_sCandIDxx = optCand4.Tag
        End If
    End Sub

    Private Sub optCand4_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles optCand1.Validated
        optCand4.ForeColor = Color.Black
    End Sub
    Private Sub optCand5_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optCand1.CheckedChanged
        If optCand5.Checked = True Then
            optCand5.ForeColor = Color.Red
            p_sCandIDxx = optCand5.Tag
        End If
    End Sub

    Private Sub optCand5_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles optCand1.Validated
        optCand5.ForeColor = Color.Black
    End Sub

    Private Sub optCand6_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optCand1.CheckedChanged
        If optCand6.Checked = True Then
            optCand6.ForeColor = Color.Red
            p_sCandIDxx = optCand6.Tag
        End If
    End Sub

    Private Sub optCand6_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles optCand1.Validated
        optCand6.ForeColor = Color.Black
    End Sub
End Class