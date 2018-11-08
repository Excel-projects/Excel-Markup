Option Strict On
Option Explicit On

Imports System.Environment
Imports System.Drawing
Imports System.Windows.Forms
Imports Markup.Scripts

Public Class frmSettings

    Private Sub frmSettings_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        Try
            My.Settings.Save()

        Catch ex As Exception
            ErrorHandler.DisplayMessage(ex)
            Exit Try

        End Try

    End Sub

    Private Sub frmSettings_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Call LoadSettings()
        'Call SetFormIcon(Me, My.Resources.Settings)
        Dim strVersion As String = My.Application.Info.Version.Major & "." & My.Application.Info.Version.Minor & "." & My.Application.Info.Version.Build & "." & My.Application.Info.Version.Revision
        Me.Text = "Settings for " & My.Application.Info.Title.ToString.Replace("&", "&&") & Space(1) & strVersion
    End Sub

    Public Sub LoadSettings()
        Try
            Me.pgdSettings.SelectedObject = My.Settings
            ''Only show "user" settings
            'Dim userAttr As New System.Configuration.UserScopedSettingAttribute
            'Dim attrs As New System.ComponentModel.AttributeCollection(userAttr)
            'pgdSettings.BrowsableAttributes = attrs

        Catch ex As Exception
            ErrorHandler.DisplayMessage(ex)

        End Try

    End Sub

    Public Sub SetFormIcon(ByRef frmCurrent As Form, ByRef bmp As Bitmap)
        Try
            frmCurrent.Icon = Icon.FromHandle(bmp.GetHicon)

        Catch ex As Exception
            ErrorHandler.DisplayMessage(ex)

        End Try

    End Sub

    Public Sub ErrorMsg(ByRef ex As Exception)
        Dim Msg As String
        Dim sf As New System.Diagnostics.StackFrame(1)
        Dim caller As System.Reflection.MethodBase = sf.GetMethod()
        Dim Proc As String = (caller.Name).Trim

        Msg = "Contact your system administrator." & vbCrLf
        Msg += "Procedure: " & Proc & vbCrLf
        Msg += "Description: " & ex.ToString & vbCrLf   '
        MsgBox(Msg, vbCritical, "Unexpected Error")

    End Sub

End Class