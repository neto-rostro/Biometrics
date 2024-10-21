Imports ggcAppDriver
Imports MySql.Data.MySqlClient

Public Class frmVotes
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
                    MsgBox("Invalid Entry for MAYOR!!!" & vbCrLf & _
                           "Please select using KeyPad!!!!!!", vbCritical, "WARNING")
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
                Call optCand1_CheckedChanged(sender, e)
            Case Keys.NumPad2
                optCand2.Checked = True
                optCand1.ForeColor = Color.Black
                optCand3.ForeColor = Color.Black
                Call optCand2_CheckedChanged(sender, e)
            Case Keys.NumPad3
                optCand3.Checked = True
                optCand1.ForeColor = Color.Black
                optCand2.ForeColor = Color.Black
                Call optCand3_CheckedChanged(sender, e)
        End Select
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim loDta As DataTable
        Dim lnCtr As Integer
        Dim lsSQL As String

        Windows.Forms.Cursor.Hide()

        loDta = New DataTable
        lsSQL = "SELECT" & _
                       "  a.sCandIDxx" & _
                       ", a.sCandName" & _
                       ", b.sTownName" & _
                    " FROM xxxCandidates a" & _
                       ", TownCity b" & _
                    " WHERE a.sTownIDxx = b.sTownIDxx" & _
                       " AND b.sTownIDxx = " & strParm(p_sTownIDxx) & _
                       " AND a.sPositIDx = '01'" & _
                    " ORDER BY a.sCandIDxx"

        loDta = p_oApp.ExecuteQuery(lsSQL)
        Label6.Text = "ALKALDE NG " & loDta(0).Item("sTownName") & "?"

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
            End Select
            lnCtr = lnCtr + 1
        Loop

        optCand4.Checked = True
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

    Private Sub Label5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label5.Click

    End Sub
End Class