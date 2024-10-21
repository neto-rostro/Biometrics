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
'       User IDs
'           GCO - M0W1180026
'           GK - M001180040
'           GCC - M001180041
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€

Option Explicit On

Imports ggcAppDriver
Imports Microsoft.VisualBasic.DateAndTime
Imports System.STAThreadAttribute

Module modMain
    Public p_oApp As New GRider("PETMgr")
    Private p_oForm As frmEvents

    Sub Main()
        p_oApp = New GRider("PETMgr")
        If Not p_oApp.LoadEnv Then
            MsgBox("Unable to Load GRider Application Driver.", MsgBoxStyle.Exclamation, "Exit")
            Exit Sub
        End If
        p_oApp.LogUser("M001111122") 'mac
        'p_oApp.LogUser("M001180040") 'gk
        'p_oApp.LogUser("M0W1180026") 'perez

        p_oForm = New frmEvents
        Application.EnableVisualStyles()
        Application.Run(p_oForm)
        p_oForm.Refresh()
    End Sub

    Function LongDate(ByVal sDate As String, Optional ByVal bMonthAbbv As Boolean = True) As String
        Dim ldDate As Date = CDate(sDate)
        Dim lnDay As Integer
        Dim lnMonth As Integer
        Dim lnYear As Integer
        Dim lsMonthName As String

        lnDay = Microsoft.VisualBasic.DateAndTime.Day(ldDate)
        lnMonth = Microsoft.VisualBasic.DateAndTime.Month(ldDate)
        lnYear = Microsoft.VisualBasic.DateAndTime.Year(ldDate)
        lsMonthName = Microsoft.VisualBasic.DateAndTime.MonthName(lnMonth)

        Return MonthName(lnMonth, bMonthAbbv) & " " & lnDay & ", " & lnYear
    End Function

    Function FormatString(ByVal value As VariantType, ByVal format As String) As String
        Dim lnLen As Integer = value.ToString.Length
        Dim lnFLen As Integer = format.ToString.Length

        If InStr(format, "-", CompareMethod.Text) > 0 Then
            Dim lasFormat() As String = Split(format, "-")
            Dim lanFormat(UBound(lasFormat)) As Integer
            Dim lnCtr As Integer = 0

            For lnCtr = 0 To UBound(lasFormat)
                lanFormat(lnCtr) = lasFormat(lnCtr).ToString.Length
            Next
        End If

        Return ""
    End Function
End Module
