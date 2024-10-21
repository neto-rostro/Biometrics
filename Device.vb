Option Explicit On

Public Class Device
    Private Const pxeMODULENAME As String = "Device"
    'ip orig sample 
    Private Const pxeDEFAULTIP As String = "192.168.15.101"
    '"192.168.15.101"'GK,GCO,Pedritos,Monarch
    '"192.168.1.201"'ROOSEVELT
    '"192.160.29.155" 

    Private Const pxeDEFAULTPORT As String = "4370"

    'Create Standalone SDK class dynamicly.
    Public axCZKEM1 As New zkemkeeper.CZKEM

    Private p_bConnected As Boolean = False
    Private p_nMachineNo As Integer

    Public Event LogUser(ByVal sBarrCode As String)
    Public Event OnFinger()

    ReadOnly Property IsConnected As Boolean
        Get
            Return p_bConnected
        End Get
    End Property
    ReadOnly Property MachineNo As Integer
        Get
            Return p_nMachineNo
        End Get
    End Property

#Region "Connection"
    Function Connect() As Boolean
        Dim lsProcName As String = pxeMODULENAME & ".Connect"
        Dim lnErrCode As Integer

        Debug.Print(lsProcName)

        p_bConnected = axCZKEM1.Connect_Net(pxeDEFAULTIP, Convert.ToInt32(pxeDEFAULTPORT))

        If p_bConnected Then
            p_nMachineNo = 1

            If axCZKEM1.RegEvent(p_nMachineNo, 65535) = True Then 'Here you can register the realtime events that you want to be triggered(the parameters 65535 means registering all)

                AddHandler axCZKEM1.OnFinger, AddressOf AxCZKEM1_OnFinger
                AddHandler axCZKEM1.OnFingerFeature, AddressOf AxCZKEM1_OnFingerFeature
                AddHandler axCZKEM1.OnVerify, AddressOf AxCZKEM1_OnVerify
                AddHandler axCZKEM1.OnAttTransactionEx, AddressOf AxCZKEM1_OnAttTransactionEx
                AddHandler axCZKEM1.OnEnrollFingerEx, AddressOf AxCZKEM1_OnEnrollFingerEx
                AddHandler axCZKEM1.OnDeleteTemplate, AddressOf AxCZKEM1_OnDeleteTemplate
                AddHandler axCZKEM1.OnNewUser, AddressOf AxCZKEM1_OnNewUser
                AddHandler axCZKEM1.OnAlarm, AddressOf AxCZKEM1_OnAlarm
                AddHandler axCZKEM1.OnDoor, AddressOf AxCZKEM1_OnDoor
                AddHandler axCZKEM1.OnWriteCard, AddressOf AxCZKEM1_OnWriteCard
                AddHandler axCZKEM1.OnEmptyCard, AddressOf AxCZKEM1_OnEmptyCard
                AddHandler axCZKEM1.OnHIDNum, AddressOf AxCZKEM1_OnHIDNum

            End If
        Else
            axCZKEM1.GetLastError(lnErrCode)
            axCZKEM1.MsgBox("Unable to connect the device,ErrorCode=" & lnErrCode, MsgBoxStyle.Exclamation, "Error")
            Return False
        End If

        Return True
    End Function

    Function Disconnect() As Boolean
        Dim lsProcName As String = pxeMODULENAME & ".Disconnect"
        Debug.Print(lsProcName)

        axCZKEM1.Disconnect()

        RemoveHandler axCZKEM1.OnFinger, AddressOf AxCZKEM1_OnFinger
        RemoveHandler axCZKEM1.OnFingerFeature, AddressOf AxCZKEM1_OnFingerFeature
        RemoveHandler axCZKEM1.OnVerify, AddressOf AxCZKEM1_OnVerify
        RemoveHandler axCZKEM1.OnAttTransactionEx, AddressOf AxCZKEM1_OnAttTransactionEx
        RemoveHandler axCZKEM1.OnEnrollFingerEx, AddressOf AxCZKEM1_OnEnrollFingerEx
        RemoveHandler axCZKEM1.OnDeleteTemplate, AddressOf AxCZKEM1_OnDeleteTemplate
        RemoveHandler axCZKEM1.OnNewUser, AddressOf AxCZKEM1_OnNewUser
        RemoveHandler axCZKEM1.OnAlarm, AddressOf AxCZKEM1_OnAlarm
        RemoveHandler axCZKEM1.OnDoor, AddressOf AxCZKEM1_OnDoor
        RemoveHandler axCZKEM1.OnWriteCard, AddressOf AxCZKEM1_OnWriteCard
        RemoveHandler axCZKEM1.OnEmptyCard, AddressOf AxCZKEM1_OnEmptyCard
        RemoveHandler axCZKEM1.OnHIDNum, AddressOf AxCZKEM1_OnHIDNum

        p_bConnected = False

        Return True
    End Function
#End Region

#Region "RealTime Events"
    'When you place your finger on sensor of the device,this event will be triggered
    Private Sub AxCZKEM1_OnFinger()
        RaiseEvent OnFinger()
        Debug.Print("RTEvent OnFinger Has been Triggered")
    End Sub
    'After you have placed your finger on the sensor(or swipe your card to the device),this event will be triggered.
    'If you passes the verification,the returned value userid will be the user enrollnumber,or else the value will be -1;
    Private Sub AxCZKEM1_OnVerify(ByVal iUserID As Integer)
        Debug.Print("RTEvent OnVerify Has been Triggered,Verifying...")
        If iUserID <> -1 Then
            Debug.Print("Verified OK,the UserID is " & iUserID.ToString())
            RaiseEvent LogUser(Right(iUserID.ToString(), 8))
        Else
            Debug.Print("Verified Failed... ")
            RaiseEvent LogUser("XXXX")
        End If
    End Sub
    'If your fingerprint(or your card) passes the verification,this event will be triggered
    Private Sub AxCZKEM1_OnAttTransactionEx(ByVal sEnrollNumber As String, ByVal iIsInValid As Integer, ByVal iAttState As Integer, ByVal iVerifyMethod As Integer, _
                      ByVal iYear As Integer, ByVal iMonth As Integer, ByVal iDay As Integer, ByVal iHour As Integer, ByVal iMinute As Integer, ByVal iSecond As Integer, ByVal iWorkCode As Integer)
        Debug.Print("RTEvent OnAttTrasactionEx Has been Triggered,Verified OK")
        Debug.Print("...UserID:" & sEnrollNumber)
        Debug.Print("...isInvalid:" & iIsInValid.ToString())
        Debug.Print("...attState:" & iAttState.ToString())
        Debug.Print("...VerifyMethod:" & iVerifyMethod.ToString())
        Debug.Print("...Workcode:" & iWorkCode.ToString()) 'the difference between the event OnAttTransaction and OnAttTransactionEx
        Debug.Print("...Time:" & iYear.ToString() & "-" & iMonth.ToString() & "-" & iDay.ToString() & " " & iHour.ToString() & ":" & iMinute.ToString() & ":" & iSecond.ToString())

        'RaiseEvent LogUser(sEnrollNumber)
    End Sub
    'When you have enrolled your finger,this event will be triggered and return the quality of the fingerprint you have enrolled
    Private Sub AxCZKEM1_OnFingerFeature(ByVal iScore As Integer)
        If iScore < 0 Then
            Debug.Print("The quality of your fingerprint is poor")
        Else
            Debug.Print("RTEvent OnFingerFeature Has been Triggered...Score:" & iScore.ToString())
        End If
    End Sub
    'When you have deleted one one fingerprint template,this event will be triggered.
    Private Sub AxCZKEM1_OnDeleteTemplate(ByVal iEnrollNumber As Integer, ByVal iFingerIndex As Integer)
        Debug.Print("RTEvent OnDeleteTemplate Has been Triggered...")
        Debug.Print("...UserID=" & iEnrollNumber.ToString() & " FingerIndex=" & iFingerIndex.ToString())
    End Sub
    'When you have enrolled a new user,this event will be triggered.
    Private Sub AxCZKEM1_OnNewUser(ByVal iEnrollNumber As Integer)
        Debug.Print("RTEvent OnNewUser Has been Triggered...")
        Debug.Print("...NewUserID=" & iEnrollNumber.ToString())
    End Sub
    'When you swipe a card to the device, this event will be triggered to show you the card number.
    Private Sub AxCZKEM1_OnHIDNum(ByVal iCardNumber As Integer)
        Debug.Print("RTEvent OnHIDNum Has been Triggered...")
        Debug.Print("...Cardnumber=" & iCardNumber.ToString())
    End Sub
    'When you are enrolling your finger,this event will be triggered.
    Private Sub AxCZKEM1_OnEnrollFingerEx(ByVal sEnrollNumber As String, ByVal iFingerIndex As Integer, ByVal iActionResult As Integer, ByVal iTemplateLength As Integer)
        If iActionResult = 0 Then
            Debug.Print("RTEvent OnEnrollFigerEx Has been Triggered....")
            Debug.Print(".....UserID: " & sEnrollNumber & " Index: " & iFingerIndex.ToString() & " tmpLen: " & iTemplateLength.ToString())
        Else
            Debug.Print("RTEvent OnEnrollFigerEx Has been Triggered Error,actionResult=" + iActionResult.ToString())
        End If
    End Sub
    '/When the dismantling machine or duress alarm occurs, trigger this event.
    Private Sub AxCZKEM1_OnAlarm(ByVal iAlarmType As Integer, ByVal iEnrollNumber As Integer, ByVal iVerified As Integer)
        Debug.Print("RTEvnet OnAlarm Has been Triggered...")
        Debug.Print("...alarmType=" & iAlarmType.ToString())
        Debug.Print("...enrollNumber=" & iEnrollNumber.ToString())
        Debug.Print("...verified=" & iVerified.ToString())
    End Sub
    'Door sensor event
    Private Sub AxCZKEM1_OnDoor(ByVal iEventType As Integer)
        Debug.Print("RTEvent Ondoor Has been Triggered...")
        Debug.Print("...EventType=" & iEventType.ToString())
    End Sub
    'When you have emptyed the Mifare card,this event will be triggered.
    Private Sub AxCZKEM1_OnEmptyCard(ByVal iActionResult As Integer)
        Debug.Print("RTEvent OnEmptyCard Has been Triggered...")
        If iActionResult = 0 Then
            Debug.Print("...Empty Mifare Card OK")
        Else
            Debug.Print("...Empty Failed")
        End If
    End Sub
    'When you have written into the Mifare card ,this event will be triggered.
    Private Sub AxCZKEM1_OnWriteCard(ByVal iEnrollNumber As Integer, ByVal iActionResult As Integer, ByVal iLength As Integer)
        Debug.Print("RTEvent OnWriteCard Has been Triggered...")
        If iActionResult = 0 Then
            Debug.Print("...Write Mifare Card OK")
            Debug.Print("...EnrollNumber=" & iEnrollNumber.ToString())
            Debug.Print("...TmpLength=" & iLength.ToString())
        Else
            Debug.Print("...Write Failed")
        End If
    End Sub
#End Region

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub
End Class
