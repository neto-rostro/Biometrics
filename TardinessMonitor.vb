'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Guanzon Software Engineering Group
' Guanzon Group of Companies
' Perez Blvd., Dagupan City
'
'    Main Office Integrated Systems Interface
'
' Copyright 2013 and Beyond
' All Rights Reserved
' ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº
' €  All  rights reserved. No part of this  software  €€  This Software is Owned by        €
' €  may be reproduced or transmitted in any form or  €€                                   €
' €  by   any   means,  electronic   or  mechanical,  €€    GUANZON MERCHANDISING CORP.    €
' €  including recording, or by information  storage  €€     Guanzon Bldg. Perez Blvd.     €
' €  and  retrieval  systems, without  prior written  €€           Dagupan City            €
' €  from the author.                                 €€  Tel No. 522-1085 ; 522-9275      €
' ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº
'
' ==========================================================================================
'  Mac [ 03/07/2015 11:30 am ]
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€

Option Explicit On

Imports ggcAppDriver

Public Class TardinessMonitor
    Private p_oApp As ggcAppDriver.GRider
    Private p_sEmployID As String
    Private p_sBranchCD As String
    Private p_sShiftIDx As String
    Private p_cFlexiTym As String
    Private p_sDeptIDxx As String
    Private p_nBreakTme As Integer
    Private p_dAMInxxxx As Date      'Supposed AM In of Employee
    Private p_dPMInxxxx As Date      'Supposed PM In of Employee
    Private p_nTardines As Integer

    Private p_dPeriodFr As Date
    Private p_dPeriodTo As Date
    Private p_dTransact As Date

    Private p_nGracePrd As Integer

    WriteOnly Property AppDriver() As GRider
        Set(ByVal oAppDriver As GRider)
            p_oApp = oAppDriver
        End Set
    End Property

    WriteOnly Property DeptIDxx As String
        Set(value As String)
            p_sDeptIDxx = value
        End Set
    End Property

    WriteOnly Property TranDate As Date
        Set(ByVal Value As Date)
            p_dTransact = Value
        End Set
    End Property

    WriteOnly Property EmployID() As String
        Set(ByVal Value As String)
            p_sEmployID = Value
        End Set
    End Property

    WriteOnly Property BranchCode() As String
        Set(ByVal Value As String)
            p_sBranchCD = Value
        End Set
    End Property

    ReadOnly Property Tardiness() As Integer
        Get
            Return p_nTardines
        End Get
    End Property

    Function ComputeTardiness() As Boolean
        Dim lsSQL As String
        Dim loDta As DataTable
        Dim lnTardyxx1 As Integer
        Dim lnTardyxx2 As Integer

        'Get Period
        lsSQL = "SELECT dPeriodFr, dPeriodTo, cPostedxx" & _
               " FROM Payroll_Period a" & _
                    " LEFT JOIN Payroll_Summary b ON a.sPayPerID = b.sPayPerID" & _
               " ORDER BY a.sPayPerID DESC LIMIT 1"
        loDta = p_oApp.ExecuteQuery(lsSQL)

        If CDate(loDta(0).Item("dPeriodTo")) >= CDate(p_dTransact.ToString("MM/dd/yyyy")) Then
            p_dPeriodFr = loDta(0).Item("dPeriodFr")
        Else
            p_dPeriodFr = DateAdd(DateInterval.Day, 1, loDta(0).Item("dPeriodTo"))
        End If

        p_dPeriodTo = DateAdd(DateInterval.Day, -1, p_dTransact)

        'Get the tardiness for the previous days of this period
        'if the current day is not the start of the period...
        If p_dPeriodTo >= p_dPeriodFr Then
            lsSQL = "SELECT SUM(nTardyxxx) nTardyxxx" & _
                   " FROM Employee_Timesheet" & _
                   " WHERE sEmployID = " & strParm(p_sEmployID) & _
                     " AND dTransact BETWEEN " & dateParm(p_dPeriodFr) & " AND " & dateParm(p_dPeriodTo)
            loDta = p_oApp.ExecuteQuery(lsSQL)
            If loDta.Rows.Count > 0 Then
                p_nTardines = IFNull(loDta(0).Item("nTardyxxx"), 0)
            End If
        End If

        lnTardyxx1 = 0
        lnTardyxx2 = 0

        'Get the log for the current date...
        lsSQL = "SELECT *" & _
               " FROM Employee_Log" & _
               " WHERE sEmployID = " & strParm(p_sEmployID) & _
                 " AND dTransact = " & dateParm(p_dTransact)
        loDta = p_oApp.ExecuteQuery(lsSQL)
        If loDta.Rows.Count > 0 Then
            'Get login of employees shift...
            Call getOtherInfo(loDta(0).Item("nTranLine"))

            If p_cFlexiTym = xeLogical.NO Then
                If IsDate(loDta(0).Item("dAMInxxxx").ToString) Then
                    lnTardyxx1 = DateDiff("n", Format(p_dAMInxxxx, "HH:mm"), Format(CDate(loDta(0).Item("dAMInxxxx").ToString), "HH:mm"))
                    lnTardyxx1 = IIf(lnTardyxx1 > p_nGracePrd, lnTardyxx1, 0)
                End If

                'Does the employee has pm login
                If IsDate(loDta(0).Item("dPMInxxxx").ToString) Then
                    'Is the employee from the main office
                    If InStr(1, "M001»M0W1»M029", p_sBranchCD) > 0 Then
                        If p_sDeptIDxx = "015" And p_sDeptIDxx = "024" Then
                            lnTardyxx2 = DateDiff("n", Format(p_dPMInxxxx, "HH:mm"), Format(CDate(loDta(0).Item("dPMInxxxx").ToString), "HH:mm"))
                        End If
                    Else
                        'Does the employee has am logout
                        If IsDate(loDta(0).Item("dAMOutxxx").ToString) Then
                            lnTardyxx2 = DateDiff("n", Format(CDate(loDta(0).Item("dAMOutxxx").ToString), "HH:mm"), Format(CDate(loDta(0).Item("dPMInxxxx").ToString), "HH:mm")) - p_nBreakTme
                        End If
                    End If
                    lnTardyxx2 = IIf(lnTardyxx2 > 0, lnTardyxx2, 0)
                End If
            End If

            p_nTardines = p_nTardines + lnTardyxx1 + lnTardyxx2
        End If

        Return True
    End Function

    Private Sub getOtherInfo(ByVal fnTranLine As Integer)
        Dim lsSQL As String
        Dim loDta As DataTable

        lsSQL = "SELECT b.sShiftIDx, b.dTimeInAM, b.dTimeInPM, b.cFlexiTym, b.nBreakTme" & _
               " FROM Employee_Shift_Change a" & _
                  " LEFT JOIN Shift b ON a.sShiftIDx = b.sShiftIDx" & _
               " WHERE a.sEmployID = " & strParm(p_sEmployID) & _
                 " AND a.dSchdShft = " & dateParm(p_dTransact) & _
                 " AND a.cTranStat IN ('1', '2', '4')"
        loDta = p_oApp.ExecuteQuery(lsSQL)

        If loDta.Rows.Count = 0 Then
            'get shift id
            lsSQL = "SELECT" & _
                           " IFNULL(a.sShiftIDx, c.sShiftIDx) sShiftIDx" & _
                   " FROM Employee_Master001 c" & _
                     " LEFT JOIN Employee_Shift a" & _
                        " ON c.sEmployID = a.sEmployID" & _
                       " AND a.nDayOfWik = DAYOFWEEK(" & dateParm(p_dTransact) & ")" & _
                   " WHERE c.sEmployID = " & strParm(p_sEmployID)
            loDta = p_oApp.ExecuteQuery(lsSQL)
            p_sShiftIDx = loDta(0).Item("sShiftIDx")

            'get schedule of shift
            lsSQL = "SELECT dTimeInAM, dTimeInPM, cFlexiTym, nBreakTme" & _
                   " FROM Shift b" & _
                   " WHERE  sShiftIDx = " & strParm(p_sShiftIDx)
            loDta = p_oApp.ExecuteQuery(lsSQL)

            p_dAMInxxxx = CDate(loDta(0).Item("dTimeInAM").ToString)
            p_dPMInxxxx = CDate(loDta(0).Item("dTimeInPM").ToString)
            p_cFlexiTym = loDta(0).Item("cFlexiTym")
            p_nBreakTme = loDta(0).Item("nBreakTme")
        Else
            p_sShiftIDx = loDta(0).Item("sShiftIDx")
            p_dAMInxxxx = CDate(loDta(0).Item("dTimeInAM").ToString)
            p_dPMInxxxx = CDate(loDta(0).Item("dTimeInPM").ToString)
            p_cFlexiTym = loDta(0).Item("cFlexiTym")
            p_nBreakTme = loDta(0).Item("nBreakTme")
        End If

        p_nGracePrd = CInt(p_oApp.getConfiguration("GRCPRD4T"))
    End Sub
End Class
