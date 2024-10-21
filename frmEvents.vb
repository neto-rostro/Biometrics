Option Explicit On

Imports ggcTimeKeeper
Imports System.Globalization
Imports System.IO
Imports ggcAppDriver
Imports System.Linq

Public Class frmEvents
    Private Const pxeClearDisp As Integer = 20
    Private Const pxeTotalLog As Integer = 4
    Private Const pxeNonVote As String = "M001"
    Private Const pxeImageLocx As String = "D:\GGC_Systems\Images\Timekeeper\"
    Private Const pxeImageExtn As String = ".jpg"

    Private WithEvents oDevice As Device
    Private oTrans As TimeKeeper
    Private oTardy As TardinessMonitor

    Private p_nLapseTme As Integer = 0
    Private p_sBarrCode As String = ""
    Private p_dTimeLogx As Date
    Private p_bNewTimex As Boolean = True
    Private p_bConnedtd As Boolean = False

    Enum xeStatus
        xeWELCOME = 0
        xeVERIFY = 1
        xeMISMATCH = 2
        xeSUCCESS = 3
        xeDUPLICATE = 4
        xeINVALID = 5
        xeFOUND = 6
        xeMAXENTRY = 7
        xeNOEMPLOY = 8
        xeREMOVEFINGER = 9
        xeNOTMET = 10
    End Enum

    Private Sub frmEvents_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
        oDevice.Disconnect()
        oDevice = Nothing
    End Sub

    Private Sub frmEvents_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Select Case e.KeyCode
            Case Keys.Escape
                If MsgBox("TimeKeeper is closing. Do you want to continue this action?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                    Me.Close()
                End If
            Case Keys.F8
                Shell("D:\GGC_Systems\vb.net\Timekeeper\VehicleLog.exe", AppWinStyle.NormalFocus)
        End Select
    End Sub

    Private Sub frmEvents_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        oDevice = New Device

        oTrans = New TimeKeeper
        oTrans.AppDriver = p_oApp
        oTrans.Branch = p_oApp.BranchCode
        oTrans.ByBranch = True
        oTrans.DisplayConfirmation = False
        oTrans.TotalDetail = pxeTotalLog
        Call oTrans.InitTransaction()

        p_bConnedtd = True ' Connect()
        If Not p_bConnedtd Then
            MsgBox("Unable to connect device. Please check connections and reload the program.", MsgBoxStyle.Critical, "Warning")
            Me.Close()
        End If

        Call ClearFields()
        Call ClearRemarks()
        Call InitDataGrid()

        Timer1.Start()
        Call showRemarks(xeStatus.xeWELCOME)

        lblDate.Text = p_oApp.SysDate.ToString("dddd, MMMM dd, yyyy")
    End Sub

    Private Function Connect() As Boolean
        Connect = oDevice.Connect
    End Function

    Public Sub IsValidForVoting()
        'MOCK ELECTION FOR PRESIDENT
        Dim loDt As DataTable
        Dim lsSQL As String
        Dim lsCondition As String

        lsSQL = "SELECT" & _
                     "   a.sBallotID" & _
                     " , a.sEmployID" & _
                     " , a.sCandIDxx" & _
                     " , a.sTownIDxx" & _
                     " , c.sPositIDx" & _
                     " , b.sTownIDxx xTownIDxx" & _
                " FROM xxxBallot a" & _
                    " LEFT JOIN Client_Master b ON a.sEmployID = b.sClientID" & _
                    " LEFT JOIN xxxCandidates c ON a.sCandIDxx = c.sCandIDxx"

        Timer2.Enabled = False 'stop timer

        If IFNull(p_oApp.getConfiguration("NtnlEltn"), "0") = "1" Then 'national election survey is activated
            'president
            lsCondition = AddCondition(lsSQL, "a.sEmployID = " & strParm(oTrans.EmployeeID) & _
                                                " AND c.sPositIDx = '03'")

            loDt = p_oApp.ExecuteQuery(lsCondition)

            If loDt.Rows.Count = 0 Then
                Me.Hide()

                Dim loForm As frmPresident
                loForm = New frmPresident
                loForm.EmployeeID = oTrans.EmployeeID
                loForm.TownIDxx = oTrans.TownID
                loForm.ShowDialog()
                loForm = Nothing
                Me.Show()
            End If

            'vice-president
            lsCondition = AddCondition(lsSQL, "a.sEmployID = " & strParm(oTrans.EmployeeID) & _
                                                " AND c.sPositIDx = '04'")

            loDt = p_oApp.ExecuteQuery(lsCondition)

            If loDt.Rows.Count = 0 Then
                Me.Hide()
                Dim loForm As frmVicePresident
                loForm = New frmVicePresident
                loForm.EmployeeID = oTrans.EmployeeID
                loForm.TownIDxx = oTrans.TownID
                loForm.ShowDialog()
                loForm = Nothing
                Me.Show()
            End If
        End If

        If IFNull(p_oApp.getConfiguration("LGUElctn"), "0") = "1" Then 'local election survey is activated
            If oTrans.TownID = "0314" Then 'dagupan only
                'mayor
                lsCondition = AddCondition(lsSQL, "a.sEmployID = " & strParm(oTrans.EmployeeID) & _
                                                    " AND c.sPositIDx = '01'")

                loDt = p_oApp.ExecuteQuery(lsCondition)

                If loDt.Rows.Count = 0 Then
                    Me.Hide()
                    Dim loForm As frmVotes
                    loForm = New frmVotes
                    loForm.EmployeeID = oTrans.EmployeeID
                    loForm.TownIDxx = oTrans.TownID
                    loForm.ShowDialog()
                    loForm = Nothing
                    Me.Show()
                End If

                'vice-mayor
                lsCondition = AddCondition(lsSQL, "a.sEmployID = " & strParm(oTrans.EmployeeID) & _
                                                    " AND c.sPositIDx = '02'")

                loDt = p_oApp.ExecuteQuery(lsCondition)

                If loDt.Rows.Count = 0 Then
                    Me.Hide()
                    Dim loForm As frmViceVotes
                    loForm = New frmViceVotes
                    loForm.EmployeeID = oTrans.EmployeeID
                    loForm.TownIDxx = oTrans.TownID
                    loForm.ShowDialog()
                    loForm = Nothing
                    Me.Show()
                End If
            End If
        End If

        p_nLapseTme = 0
        Timer2.Enabled = True 'start timer
    End Sub

    Private Sub oDevice_LogUser(ByVal sBarrCode As String) Handles oDevice.LogUser
        saveLog(sBarrCode)
    End Sub

    Private Sub saveLog(ByVal sBarrCode As String)
        With oTrans
            If .EmployeeID = "" Then
                showRemarks(xeStatus.xeNOEMPLOY)
                GoTo endProc
            End If

            If p_sBarrCode = "" Then GoTo endProc
            If sBarrCode = "" Then GoTo endProc

            Dim lsBarrCode As String = sBarrCode.PadLeft(4, "0")

            If p_sBarrCode = lsBarrCode Then

                .LogTime = p_dTimeLogx

                If .SaveTransaction Then
                    Call IsValidForVoting()
                    LoadEmpLog()
                    showRemarks(xeStatus.xeSUCCESS)
                End If

                p_sBarrCode = ""
            Else
                showRemarks(xeStatus.xeMISMATCH)
            End If
        End With
endProc:
        p_nLapseTme = 0
    End Sub

    Private Sub txtBarCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtBarCode.GotFocus
        txtBarCode.BackColor = Color.LightYellow
    End Sub

    Private Sub TextBox1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtBarCode.KeyDown
        Select Case e.KeyCode
            Case Keys.Enter
                'kalyptus - 2017.02.22 04:15pm
                'Set the branch for each entry of barcode...
                If Len(txtBarCode.Text) = 8 Then
                    Dim lsBranch As String = Strings.Left(txtBarCode.Text, 4)
                    If validBranch(lsBranch) Then
                        oTrans.Branch = lsBranch
                    Else
                        oTrans.Branch = p_oApp.BranchCode
                    End If
                Else
                    oTrans.Branch = p_oApp.BranchCode
                End If

                Timer2.Enabled = False
                Call getEmployee(Strings.Right(txtBarCode.Text, 4))
                Timer2.Enabled = True

                txtBarCode.Text = ""
        End Select
    End Sub

    Private Function getEmployee(ByVal lsBarrCode As String) As Boolean
        If Not IsNumeric(lsBarrCode) Then Return False

        If Not p_bConnedtd Then
            MsgBox("Fingerprint Device is not Connected. Reload the Program", MsgBoxStyle.Exclamation, "Warning")
            Return False
        End If

        If lsBarrCode = "" Then
            Return False
        End If

        If Replace(lsBarrCode, "0", "") = "" Then
            Return False
        End If

        With oTrans
            Dim ldLog = p_oApp.SysDate

            If p_bNewTimex Then
                p_dTimeLogx = ldLog
                p_bNewTimex = False
            End If

            If .getEmployee(lsBarrCode, p_dTimeLogx) Then
                Beep()

                oDevice.axCZKEM1.Beep(0)

                p_sBarrCode = lsBarrCode

                lblName.Text = .FirstName & " " & .LastName & IIf(.SuffixName <> "", " " & .SuffixName, "")
                lblDept.Text = .Department
                lblLogType.Text = IIf(.LogType = TimeKeeper.xeLogType.xeLogout, "LOG OUT", "LOG IN")
                lblLogTime.Text = Format(ldLog, "hh:mm:ss tt")

                If oTrans.SecType = "1" Then

                    If isLogTimeIntervalOk() = True Then
                        If oTrans.SaveTransaction = True Then
                            Call IsValidForVoting()
                            LoadEmpLog()
                            showRemarks(xeStatus.xeSUCCESS)
                        End If
                    Else
                        LoadEmpLog()
                        Call IsValidForVoting()
                        Call showRemarks(xeStatus.xeNOTMET)
                    End If
                    p_nLapseTme = 0
                    Timer2.Start()
                    GoTo endProc
                End If


                Call LoadEmpLog()
                Call LoadPicture()
                Call showRemarks(xeStatus.xeFOUND)

                p_nLapseTme = 0
                Timer2.Start()
            Else
                oTrans.initMaster()
                p_sBarrCode = ""

                Call showRemarks(xeStatus.xeINVALID)
                Return False
            End If
        End With
Endproc:
        Return True
    End Function

    Public Function isLogTimeIntervalOk() As Boolean
        For lnCtr = 0 To oTrans.TotalLog
            If oTrans.DetailTranDate(lnCtr) = "1900-01-01" Then GoTo NextLoop
            Select Case oTrans.DetailTranLine(lnCtr)
                Case 2
                    If DateDiff(DateInterval.Minute, _
                                CDate(Format(oTrans.DetailTranDate(lnCtr), "MMM dd, yyyy") & " " & oTrans.DetailAMIn(lnCtr)),
                                p_dTimeLogx) <= 10 Then Return False
                Case 3
                    If DateDiff(DateInterval.Minute, _
                                CDate(Format(oTrans.DetailTranDate(lnCtr), "MMM dd, yyyy") & " " & oTrans.DetailAMOut(lnCtr)),
                                p_dTimeLogx) <= 10 Then Return False

                Case 4
                    If DateDiff(DateInterval.Minute, _
                                CDate(Format(oTrans.DetailTranDate(lnCtr), "MMM dd, yyyy") & " " & oTrans.DetailPMIn(lnCtr)),
                                p_dTimeLogx) <= 10 Then Return False

                Case 5
                    If DateDiff(DateInterval.Minute, _
                                CDate(Format(oTrans.DetailTranDate(lnCtr), "MMM dd, yyyy") & " " & oTrans.DetailPMOut(lnCtr)),
                                p_dTimeLogx) <= 10 Then Return False

                Case 6
                    If DateDiff(DateInterval.Minute, _
                                CDate(Format(oTrans.DetailTranDate(lnCtr), "MMM dd, yyyy") & " " & oTrans.DetailOTimeIn(lnCtr)),
                                p_dTimeLogx) <= 10 Then Return False

            End Select
NextLoop:
        Next

        Return True
    End Function

    Private Sub ClearFields()
        oTrans.initMaster()

        lblName.Text = ""
        lblDept.Text = ""
        lblLogTime.Text = ""
        lblLogType.Text = ""
        txtBarCode.Text = ""

        InitDataGrid()

        picStat.BackgroundImage = My.Resources.check1
        PictureBox1.BackgroundImage = Nothing
        showRemarks(xeStatus.xeWELCOME)

        p_bNewTimex = True
        p_dTimeLogx = p_oApp.SysDate
        Timer2.Stop()

        p_nLapseTme = 0

        'kalyptus - 2017.02.22 02:04pm
        'set initial value of tardiness to 0
        lblTardiness.Text = "TARDINESS: 0"
        lblWarning.Visible = False

    End Sub

    Private Sub ClearRemarks()
        lblRemarks.Text = ""
    End Sub

    Private Sub Timer1_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        lblTime.Text = Format(p_oApp.SysDate, "hh:mm:ss")
        If Not txtBarCode.Focused Then txtBarCode.Focus()
    End Sub

    Sub showRemarks(ByVal nStat As xeStatus)

        Dim lsRemark As String

        Select Case nStat
            Case 0 'welcome
                lsRemark = "Magandang Buhay! Welcome to Guanzon."
                picStat.BackgroundImage = Nothing
            Case 1 'verify
                lsRemark = "Verifying . . ."
            Case 2 'not matched
                lsRemark = "Authentication Failed. Please try again."
                picStat.BackgroundImage = My.Resources.wrong
            Case 3 'success
                lsRemark = "Log Successfully Saved."
                picStat.BackgroundImage = My.Resources.check1
            Case 4 'duplicate
                lsRemark = "Duplicate Entry."
                picStat.BackgroundImage = Nothing
            Case 5 'invalid
                lsRemark = "Employee Finger Print Not Matched."
                picStat.BackgroundImage = Nothing
            Case 6 'found
                lsRemark = "Place you finger on the sensor."
                picStat.BackgroundImage = Nothing
            Case 7
                lsRemark = "Warning. Maximum Log Entry Reached."
                picStat.BackgroundImage = Nothing
            Case 8
                lsRemark = "Fingerprint Detected! No Employee Loaded!"
            Case 9
                lsRemark = "Please Remove Finger from the Sensor."
            Case 10 ' not met
                lsRemark = "Invalid. Log Time does not met the criteria."
            Case Else
                lsRemark = ""
                picStat.BackgroundImage = Nothing
        End Select

        lblRemarks.Text = lsRemark
    End Sub

    Sub InitDataGrid()
        With DataGridView1
            .Rows.Clear()
            .Rows.Add(pxeTotalLog)
        End With
    End Sub

    Sub LoadEmpLog()
        Dim lbCancel As Boolean = False

        With DataGridView1
            Dim lnCtr As Integer

            For lnCtr = 0 To oTrans.TotalLog
                If oTrans.DetailTranDate(lnCtr) = "1900-01-01" Then GoTo nextLoop
                .Rows(lnCtr).Cells(0).Value = Format(oTrans.DetailTranDate(lnCtr), "MMM dd, yyyy")
                .Rows(lnCtr).Cells(1).Value = oTrans.DetailAMIn(lnCtr)
                .Rows(lnCtr).Cells(2).Value = oTrans.DetailAMOut(lnCtr)
                .Rows(lnCtr).Cells(3).Value = oTrans.DetailPMIn(lnCtr)
                .Rows(lnCtr).Cells(4).Value = oTrans.DetailPMOut(lnCtr)
                '.Rows(lnCtr).Cells(5).Value = oTrans.DetailOTimeIn(lnCtr)
                '.Rows(lnCtr).Cells(6).Value = oTrans.DetailOTimeOut(lnCtr)

                If CDate(Format(oTrans.DetailTranDate(lnCtr), "MMM dd, yyyy")) = CDate(p_oApp.SysDate.ToString("MMM dd, yyyy")) Then
                    Select Case oTrans.DetailTranLine(lnCtr)
                        Case 2
                            If DateDiff(DateInterval.Minute, _
                                        CDate(Format(oTrans.DetailTranDate(lnCtr), "MMM dd, yyyy") & " " & oTrans.DetailAMIn(lnCtr)),
                                        p_dTimeLogx) <= 10 Then lbCancel = True
                        Case 3
                            If DateDiff(DateInterval.Minute, _
                                        CDate(Format(oTrans.DetailTranDate(lnCtr), "MMM dd, yyyy") & " " & oTrans.DetailAMOut(lnCtr)),
                                        p_dTimeLogx) <= 10 Then lbCancel = True
                        Case 4
                            If DateDiff(DateInterval.Minute, _
                                        CDate(Format(oTrans.DetailTranDate(lnCtr), "MMM dd, yyyy") & " " & oTrans.DetailPMIn(lnCtr)),
                                        p_dTimeLogx) <= 10 Then lbCancel = True
                        Case 5
                            If DateDiff(DateInterval.Minute, _
                                        CDate(Format(oTrans.DetailTranDate(lnCtr), "MMM dd, yyyy") & " " & oTrans.DetailPMOut(lnCtr)),
                                        p_dTimeLogx) <= 10 Then lbCancel = True
                            'Case 6
                            '    If DateDiff(DateInterval.Minute, _
                            '                CDate(Format(oTrans.DetailTranDate(lnCtr), "MMM dd, yyyy") & " " & oTrans.DetailOTimeIn(lnCtr)),
                            '                p_dTimeLogx) <= 10 Then lbCancel = True
                        Case Else
                            showRemarks(xeStatus.xeMAXENTRY)
                    End Select
                End If
nextLoop:
            Next

            If lbCancel Then
                showRemarks(xeStatus.xeDUPLICATE)
                p_sBarrCode = ""
            End If
        End With
    End Sub

    Sub LoadPicture()
        Dim lsImgName As String

        With oTrans
            lsImgName = Trim(pxeImageLocx & .LastName & Strings.Left(.FirstName, 3) & pxeImageExtn)

            If File.Exists(lsImgName) Then
                PictureBox1.BackgroundImage = Image.FromFile(lsImgName)
            Else
                PictureBox1.BackgroundImage = Nothing
            End If

            'If .EmployeeID = "M00113001232" Then
            '    PictureBox1.BackgroundImage = My.Resources.MacarlingFed
            'Else
            'End If
        End With
    End Sub

    Private Sub Timer2_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Timer2.Tick
        p_nLapseTme = p_nLapseTme + 1

        If p_nLapseTme >= pxeClearDisp Then
            oTrans.initMaster()
            ClearFields()
            ClearRemarks()
            p_nLapseTme = 0
        End If
    End Sub

    Private Sub oDevice_OnFinger() Handles oDevice.OnFinger
        showRemarks(xeStatus.xeREMOVEFINGER)
        picStat.BackgroundImage = Nothing
    End Sub

    Private Sub txtBarCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtBarCode.LostFocus
        txtBarCode.BackColor = Color.Snow
    End Sub

    Private Function validBranch(ByVal fsBranchCD As String) As Boolean
        Dim lsSQL As String
        Dim loDta As DataTable

        lsSQL = "SELECT sBranchCD" & _
               " FROM Branch" & _
               " WHERE sBranchCD = " & strParm(fsBranchCD)
        loDta = p_oApp.ExecuteQuery(lsSQL)

        validBranch = loDta.Rows.Count > 0
    End Function
End Class
