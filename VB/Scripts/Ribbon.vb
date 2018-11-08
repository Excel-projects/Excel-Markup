Option Strict Off
Option Explicit On

Imports System.Diagnostics
Imports System.Windows.Forms
Imports Markup.Scripts
'Imports System.IO.Path
'Imports System.Runtime.InteropServices
'Imports Microsoft.Office.Tools.Ribbon
'Imports Excel = Microsoft.Office.Interop.Excel
'Imports System.Windows

Namespace Scripts

    <Runtime.InteropServices.ComVisible(True)>
    Public Class Ribbon
        Implements Office.IRibbonExtensibility
        Private ribbon As Office.IRibbonUI
        Public Shared ribbonref As Ribbon

#Region "  Ribbon Events  "

        Public Sub New()
        End Sub

        Public Function GetCustomUI(ByVal ribbonID As String) As String Implements Office.IRibbonExtensibility.GetCustomUI
            Return GetResourceText("Markup.Ribbon.xml")
        End Function

        Private Shared Function GetResourceText(ByVal resourceName As String) As String
            Dim asm As Reflection.Assembly = Reflection.Assembly.GetExecutingAssembly()
            Dim resourceNames() As String = asm.GetManifestResourceNames()
            For i As Integer = 0 To resourceNames.Length - 1
                If String.Compare(resourceName, resourceNames(i), StringComparison.OrdinalIgnoreCase) = 0 Then
                    Using resourceReader As IO.StreamReader = New IO.StreamReader(asm.GetManifestResourceStream(resourceNames(i)))
                        If resourceReader IsNot Nothing Then
                            Return resourceReader.ReadToEnd()
                        End If
                    End Using
                End If
            Next
            Return Nothing
        End Function

        Public Sub Ribbon_Load(ByVal ribbonUI As Office.IRibbonUI)
            Me.ribbon = ribbonUI
        End Sub

        Public Function GetButtonImage(ByVal control As Office.IRibbonControl) As System.Drawing.Bitmap
            Try
                Select Case control.Id.ToString
                    Case "grpRevision", "btnRev" : Return My.Resources.Resources.RevTri
                    Case "grpMarkups", "btnCloudAll" : Return My.Resources.Resources.Cloud
                    Case "btnCloudHold" : Return My.Resources.Resources.CloudHold
                    Case "btnCloudHatch" : Return My.Resources.Resources.CloudHatch
                    Case "btnAreaHatch" : Return My.Resources.Resources.Hatch
                    Case "btnCloudPartLeft" : Return My.Resources.Resources.CloudPartLeft
                    Case "btnCloudPartRight" : Return My.Resources.Resources.CloudPartRight
                    Case "btnCloudPartTop" : Return My.Resources.Resources.CloudPartTop
                    Case "btnCloudPartBottom" : Return My.Resources.Resources.CloudPartBottom
                    Case Else : Return Nothing
                End Select

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)
                Return Nothing

            End Try

        End Function

        Public Function GetEnabled(ByVal control As Office.IRibbonControl) As Boolean
            Try

                Select Case control.Id
                    Case "btnRev", "btnCloudAll", "btnCloudHold", "btnCloudHatch", "btnAreaHatch", "btnCloudPartLeft", "btnCloudPartRight", "btnCloudPartTop", "btnCloudPartBottom"
                        Return ErrorHandler.IsEnabled(False)
                    Case Else
                        Return False
                End Select

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)
                Return False
            End Try
        End Function

        Public Function GetLabelText(ByVal control As Office.IRibbonControl) As String
            Try
                Select Case control.Id
                    Case "tabMarkup"
                        If Application.ProductVersion.Substring(0, 2) = "15" Then
                            Return My.Application.Info.Title.ToUpper()
                        Else
                            Return My.Application.Info.Title
                        End If

                    Case "txtCopyright"
                        Return "© " + My.Application.Info.Copyright.ToString
                    Case "txtDescription"
                        Dim AppVersion As String = My.Application.Info.Version.Major & "." & My.Application.Info.Version.Minor & "." & My.Application.Info.Version.Build & "." & My.Application.Info.Version.Revision
                        Return My.Application.Info.Title.Replace("&", "&&") + " " + AppVersion
                    Case "txtReleaseDate"
                        Dim dteCreateDate As DateTime = My.Settings.App_ReleaseDate
                        Return dteCreateDate.ToString("dd-MMM-yyyy hh:mm tt")
                    Case "txtRevisionCharacter"
                        Return My.Settings.Markup_TriangleRevisionCharacter
                    Case Else
                        Return String.Empty
                End Select

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)
                Return String.Empty

            End Try

        End Function

        Public Function GetItemCount(ByVal Control As Office.IRibbonControl) As Integer
            Try
                Select Case Control.Id.ToString
                    Case "drpColorType"
                        Return 4
                    Case Else
                        Return 0
                End Select

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)
                Return 0

            End Try

        End Function

        Public Function GetItemLabel(ByVal Control As Office.IRibbonControl, Index As Integer) As String
            Try
                Select Case Control.Id.ToString
                    Case "drpColorType"
                        Select Case Index
                            Case 0
                                Return "BLACK"
                            Case 1
                                Return "BLUE"
                            Case 2
                                Return "RED"
                            Case 3
                                Return "GREEN"
                            Case Else
                                Return String.Empty
                        End Select
                    Case Else
                        Return String.Empty
                End Select

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)
                Return String.Empty

            End Try

        End Function

        Public Sub MyAction(ByVal Control As Office.IRibbonControl, ItemId As String, Index As Integer)
            Try
                Select Case Control.Id.ToString
                    Case "drpColorType"
                        Select Case Index
                            Case 0
                                My.Settings.Markup_ShapeLineColor = System.Drawing.Color.Black
                            Case 1
                                My.Settings.Markup_ShapeLineColor = System.Drawing.Color.Green
                            Case 2
                                My.Settings.Markup_ShapeLineColor = System.Drawing.Color.Red
                            Case 3
                                My.Settings.Markup_ShapeLineColor = System.Drawing.Color.Blue
                            Case Else
                                My.Settings.Markup_ShapeLineColor = System.Drawing.Color.Black
                        End Select
                    Case Else
                        'nothing
                End Select

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)
                My.Settings.Markup_ShapeLineColor = System.Drawing.Color.Black

            End Try

        End Sub

        Public Sub GetSelectedItemID(ByVal Control As Office.IRibbonControl, ByRef itemID As Object)
            Try
                Select Case Control.Id.ToString
                    Case "drpColorType"
                        'itemID = My.Settings.ShapeLineColor
                    Case Else
                        itemID = String.Empty
                End Select

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)
                itemID = String.Empty

            End Try

        End Sub

        Public Sub OnAction(ByVal control As Office.IRibbonControl)
            Try

                Select Case control.Id
                    Case "btnSelectColor"
                        SelectLineColor()
                    Case "btnUpdateColor"
                        'UpdateLineColor()
                    Case "btnRev"
                        CreateRevisionTriangle()
                    Case "btnCloudAll"
                        CreateCloudPart("ALL")
                    Case "btnCloudHold"
                        CreateCloudHold()
                    Case "btnCloudHatch"
                        CreateCloudHatch()
                    Case "btnAreaHatch"
                        CreateAreaHatching()
                    Case "btnCloudPartLeft"
                        CreateCloudPart("L")
                    Case "btnCloudPartRight"
                        CreateCloudPart("R")
                    Case "btnCloudPartTop"
                        CreateCloudPart("T")
                    Case "btnCloudPartBottom"
                        CreateCloudPart("B")
                    Case "btnRemoveLastShape"
                        RemoveLastShape()
                    Case "btnRemoveAllShapes"
                        RemoveAllShapes()
                    Case "btnSettings"
                        OpenSettings()
                    Case "btnOpenReadMe"
                        OpenReadMe()
                    Case "btnOpenNewIssue"
                        'OpenNewIssue()
                End Select

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)
            End Try
        End Sub

#End Region

#Region "  Ribbon Buttons  "

        Public Sub CreateRevisionTriangle()
            Try
                If ErrorHandler.IsEnabled() Then
                    Dim ShapeFirstCnt As Integer = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Count
                    Dim triArray(0 To 3, 0 To 1) As Single
                    Dim x As Integer = 0
                    Dim y As Integer = Globals.ThisAddIn.Application.Selection.Top
                    Dim h As Integer = Globals.ThisAddIn.Application.Selection.RowHeight
                    Dim w As Integer = Convert.ToInt32(h * 2.2 / Math.Sqrt(3))
                    Dim f As Integer = Globals.ThisAddIn.Application.Selection.Font.Size
                    If Globals.ThisAddIn.Application.Selection.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter And Globals.ThisAddIn.Application.Selection.Width > w Then
                        x = Globals.ThisAddIn.Application.Selection.Left + (Globals.ThisAddIn.Application.Selection.Width - w) / 2
                    Else
                        x = Globals.ThisAddIn.Application.Selection.Left
                    End If
                    triArray(0, 0) = CSng(x + w / 2)
                    triArray(0, 1) = y
                    triArray(1, 0) = x
                    triArray(1, 1) = y + h
                    triArray(2, 0) = x + w
                    triArray(2, 1) = y + h
                    triArray(3, 0) = CSng(x + w / 2)    ' Last point has same coordinates as first
                    triArray(3, 1) = y
                    Dim tri As String = Globals.ThisAddIn.Application.ActiveSheet.Shapes.AddPolyline(triArray).Name
                    Globals.ThisAddIn.Application.ActiveSheet.Shapes(tri).Line.Weight = 1.5 'reset the line weight of the triangle
                    Globals.ThisAddIn.Application.ActiveSheet.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, CSng(x), CSng(y + h * 0.2), CSng(w), CSng(h * 0.8)).Select() 'make the textbox wider
                    'Dim tb As String = Globals.ThisAddIn.Application.Selection.Name
                    My.Settings.Markup_TriangleRevisionCharacter = InputBox("Please enter a revision.", "Revision Triangle", My.Settings.Markup_TriangleRevisionCharacter)
                    Globals.ThisAddIn.Application.Selection.Characters.Text = My.Settings.Markup_TriangleRevisionCharacter
                    Globals.ThisAddIn.Application.Selection.Font.Color = My.Settings.Markup_ShapeLineColor
                    Globals.ThisAddIn.Application.Selection.Font.Size = f '* 0.8 'may want to make this smaller
                    Globals.ThisAddIn.Application.Selection.Border.LineStyle = Excel.Constants.xlNone
                    Globals.ThisAddIn.Application.Selection.Interior.ColorIndex = Excel.Constants.xlNone
                    Globals.ThisAddIn.Application.Selection.Shadow = False
                    Globals.ThisAddIn.Application.Selection.RoundedCorners = False
                    Globals.ThisAddIn.Application.Selection.HorizontalAlignment = Excel.Constants.xlCenter
                    Globals.ThisAddIn.Application.Selection.VerticalAlignment = Excel.Constants.xlCenter
                    Globals.ThisAddIn.Application.Selection.AutoSize = True 'autosize the text in the triangle
                    Dim ShapeLastCnt As Integer = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Count
                    CreateShapeGroup(ShapeLastCnt, ShapeFirstCnt)
                    'Globals.ThisAddIn.Application.ActiveSheet.Shapes.Range(Range(tri, tb)).Group.Select()
                    Globals.ThisAddIn.Application.ActiveSheet.Shapes(tri).select()
                    Globals.ThisAddIn.Application.Selection.Interior.Pattern = Excel.XlPattern.xlPatternNone 'added for Excel 2007
                    SetLineColor()
                    Globals.ThisAddIn.Application.ActiveCell.Select()
                End If

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)

            End Try

        End Sub

        Public Sub CreateCloudHold()
            Try
                If ErrorHandler.IsEnabled() Then
                    Dim x As Double = Globals.ThisAddIn.Application.Selection.Left
                    Dim y As Double = Globals.ThisAddIn.Application.Selection.Top
                    Dim h As Double = Globals.ThisAddIn.Application.Selection.Height
                    Dim w As Double = Globals.ThisAddIn.Application.Selection.Width
                    Dim off As Double = 7.5
                    x = x - off / 2
                    w = w + off
                    y = y - off / 2
                    h = h + off
                    Dim ShapeFirstCnt As Integer = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Count
                    CreateCloudLine(x, y + h, x, y)          ' Left
                    CreateCloudLine(x + w, y + h, x, y + h)  ' Bottom
                    CreateCloudLine(x + w, y, x + w, y + h)  ' Right
                    CreateCloudLine(x, y, x + w, y)          ' Top
                    Dim ShapeLastCnt As Integer = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Count
                    CreateShapeGroup(ShapeLastCnt, ShapeFirstCnt)
                End If

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)

            End Try

        End Sub

        Public Sub CreateCloudHatch()
            Try
                If ErrorHandler.IsEnabled() Then
                    Dim x As Double = Globals.ThisAddIn.Application.Selection.Left
                    Dim y As Double = Globals.ThisAddIn.Application.Selection.Top
                    Dim h As Double = Globals.ThisAddIn.Application.Selection.Height
                    Dim w As Double = Globals.ThisAddIn.Application.Selection.Width
                    Dim ShapeFirstCnt As Integer = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Count
                    CreateCloudPart("ALL")
                    CreateHatchArea(x, y, h, w)
                    Dim ShapeLastCnt As Integer = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Count
                    CreateShapeGroup(ShapeLastCnt, ShapeFirstCnt)
                End If

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)

            End Try

        End Sub

        Public Sub CreateAreaHatching()
            Try
                If ErrorHandler.IsEnabled() Then
                    Dim SingleArea As Excel.Range 'Object Excel.Areas
                    Dim SelectAreaCnt As Integer = Globals.ThisAddIn.Application.Selection.Areas.Count
                    Dim ShapeFirstCnt As Integer = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Count
                    If SelectAreaCnt > 1 Then
                        For Each SingleArea In Globals.ThisAddIn.Application.Selection.Areas
                            CreateHatchArea(SingleArea.Left, SingleArea.Top, SingleArea.Height, SingleArea.Width)
                        Next
                    Else
                        CreateHatchArea(Globals.ThisAddIn.Application.Selection.Left, Globals.ThisAddIn.Application.Selection.Top, Globals.ThisAddIn.Application.Selection.Height, Globals.ThisAddIn.Application.Selection.Width)
                    End If
                    Dim CntShapeLast As Integer = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Count
                    CreateShapeGroup(CntShapeLast, ShapeFirstCnt)
                End If

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)

            End Try

        End Sub

        Public Sub SelectLineColor()
            My.Settings.Markup_ShapeLineColor = SelectColor()
        End Sub

        Public Sub RemoveLastShape()
            Try
                Dim result As Integer = MessageBox.Show("Are you sure you would like to delete the last shape that has been created?", "Delete Last Shape?", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                If result = DialogResult.No Then
                    Exit Try
                ElseIf result = DialogResult.Yes Then
                    Dim xlWorkSheet As Excel.Worksheet = Globals.ThisAddIn.Application.ActiveSheet
                    Dim strShapeName As String = My.Settings.Markup_LastShapeName
                    For Each shp In xlWorkSheet.Shapes
                        ' To avoid datavalidation shape getting deleted
                        If shp.name = strShapeName Then
                            shp.Delete()
                        End If
                    Next
                End If

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)

            End Try

        End Sub

        Public Sub RemoveAllShapes()
            Try
                Dim shp As Excel.Shape
                Dim xlWorkSheet As Excel.Worksheet = Globals.ThisAddIn.Application.ActiveSheet

                Dim result As Integer = MessageBox.Show("Are you sure you would like to delete all the shapes in the active worksheet?", "Delete All Shapes?", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                If result = DialogResult.No Then
                    Exit Try
                ElseIf result = DialogResult.Yes Then
                    ' Loop through all the shapes in the worksheet and delete them
                    For Each shp In xlWorkSheet.Shapes
                        ' To avoid datavalidation shape getting deleted
                        If shp.Type <> 8 Then shp.Delete()
                    Next
                End If

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)

            End Try

        End Sub

        Public Sub OpenSettings()
            Try
                Dim FormSettings As New frmSettings
                FormSettings.ShowDialog()
                ribbon.Invalidate()

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)

            End Try

        End Sub

        Public Sub CreateEmailMessage()
            Try
                'Dim Msg As String = "mailto:" + My.Settings.HelpEmail
                'Dim Product As String = My.Application.Info.Title.ToString.Replace("&", "&&")
                'Msg += "?subject=" + Product & " " & My.Application.Info.Version.ToString
                'Msg += "&body=Please create a ticket for user " + Environment.UserName + " "
                'Msg += "on machine " + Environment.MachineName + " "
                'Msg += "and assign it to group: " + My.Settings.HelpGroup + "."
                'System.Diagnostics.Process.Start(Msg)

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)

            End Try

        End Sub

        Public Sub OpenReadMe()
            OpenFile(My.Settings.App_PathReadMe)
        End Sub

#End Region

#Region "  Subroutines  "

        Public Sub CreateArc(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double, ByVal Length As Double)
            Try
                If ErrorHandler.IsEnabled() Then
                    Dim shapeName As String = "CloudArc"
                    Dim i As Integer = 0
                    Dim angle As Double = 60
                    Dim segments As Double = angle / 10
                    Dim arcArray(0 To CInt(segments), 0 To 1) As Single
                    Dim theta As Double = Angle * Math.PI / 180
                    Dim xm As Double = (x1 + x2) / 2
                    Dim ym As Double = (y1 + y2) / 2
                    Dim xd As Double = (x2 - x1)
                    Dim yd As Double = (y2 - y1)
                    Dim d As Double = Math.Sqrt(xd * xd + yd * yd)
                    Dim r As Double = d / 2 / Math.Sin(theta / 2)
                    Dim h As Double = r * Math.Cos(theta / 2)
                    Dim xc As Double = xm + yd / (2 * Math.Tan(theta / 2))
                    Dim yc As Double = ym - xd / (2 * Math.Tan(theta / 2))
                    Dim dtheta As Double = theta / segments
                    arcArray(0, 0) = CSng(x1)
                    arcArray(0, 1) = CSng(y1)
                    Dim a As Double = Math.Atan2(y1 - yc, x1 - xc) - dtheta
                    'Debug.WriteLine("theta, xm, ym, xd, yd, d, r, h, xc, yc, dtheta: " & theta & ", " & xm & ", " & ym & ", " & xd & ", " & yd & ", " & d & ", " & r & ", " & h & ", " & xc & ", " & yc & ", " & dtheta)
                    For i = 1 To CInt(segments) - 1
                        arcArray(i, 0) = CSng(xc + r * Math.Cos(a))
                        arcArray(i, 1) = CSng(yc + r * Math.Sin(a))
                        a = a - dtheta
                    Next
                    arcArray(i, 0) = CSng(x2)
                    arcArray(i, 1) = CSng(y2)
                    Dim cloudArc As Excel.Shape = Globals.ThisAddIn.Application.ActiveSheet.Shapes.AddPolyline(arcArray)
                    cloudArc.Select()
                    cloudArc.Name = shapeName + " " + DateTime.Now.ToString(My.Settings.Markup_ShapeDateFormat)
                    Globals.ThisAddIn.Application.Selection.Interior.Pattern = Excel.Constants.xlNone
                    SetLineColor()
                End If

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)

            End Try

        End Sub

        Public Sub CreateCloudLine(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double)
            Try
                If ErrorHandler.IsEnabled() Then
                    Dim shapeName As String = "CloudLine"
                    Dim length As Double = My.Settings.Markup_ShapeLineSpacing
                    Dim i As Integer = 0
                    Dim x As Double = 0
                    Dim y As Double = 0
                    Dim dx As Double = x2 - x1
                    Dim dy As Double = y2 - y1
                    Dim d As Double = Math.Sqrt(dx * dx + dy * dy)
                    Dim segments As Double = Fix(d / length)
                    If segments < 2 Then segments = 2
                    Dim deltax As Double = (dx / segments)
                    Dim deltay As Double = (dy / segments) ' Convert.ToInt32
                    Dim xp As Double = x1
                    Dim yp As Double = y1
                    Debug.WriteLine("xp, yp, x, y: " & xp & ", " & yp & ", " & x & ", " & y)
                    For i = 1 To CInt(segments)
                        x = xp + deltax
                        y = yp + deltay
                        CreateArc(xp, yp, x, y, length)
                        xp = x : yp = y
                    Next
                End If

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)

            End Try

        End Sub

        Public Sub CreateCloudPart(ByVal CloudPart As String)
            Try
                If ErrorHandler.IsEnabled() Then
                    Dim i As Integer = 0
                    Dim x As Integer = Globals.ThisAddIn.Application.Selection.Left
                    Dim y As Integer = Globals.ThisAddIn.Application.Selection.Top
                    Dim h As Integer = Globals.ThisAddIn.Application.Selection.Height
                    Dim w As Integer = Globals.ThisAddIn.Application.Selection.Width
                    Dim ShapeFirstCnt As Integer = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Count
                    If CloudPart = "B" Or CloudPart = "ALL" Then
                        CreateCloudLine(x, y + h, x + w, y + h)
                    End If
                    If CloudPart = "T" Or CloudPart = "ALL" Then
                        CreateCloudLine(x + w, y, x, y)
                    End If
                    If CloudPart = "L" Or CloudPart = "ALL" Then
                        CreateCloudLine(x, y, x, y + h)
                    End If
                    If CloudPart = "R" Or CloudPart = "ALL" Then
                        CreateCloudLine(x + w, y + h, x + w, y)
                    End If
                    Dim ShapeLastCnt As Integer = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Count
                    CreateShapeGroup(ShapeLastCnt, ShapeFirstCnt)
                End If

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)

            End Try

        End Sub

        Public Sub CreateHatchArea(ByVal x As Double, ByVal y As Double, ByVal h As Double, ByVal w As Double)
            Try
                If ErrorHandler.IsEnabled() Then
                    Dim length As Double = My.Settings.Markup_ShapeLineSpacing
                    Dim xx1, yy1 As Double
                    Dim xx2, yy2 As Double
                    Dim x1 As Double = x + y
                    Dim x2 As Double = Fix(x1 / length)
                    Dim x3 As Double = (x2 + 1) * length
                    Dim xDiff As Double = x3 - x1
                    Dim xl As Double = x
                    Dim xr As Double = x + w
                    Dim xw As Double = w
                    Dim yt As Double = y
                    Dim yb As Double = y + h
                    Dim yw As Double = h
                    Dim xsp As Double = x + xDiff
                    Dim ysp As Double = yt + xDiff
                    If xw >= yw Then
                        While xsp < xr + yw
                            If xsp - xl < yb - yt Then
                                xx1 = xsp
                                yy1 = yt
                                xx2 = xl
                                yy2 = yt + (xx1 - xl)
                            Else
                                If xsp <= xr Then
                                    xx1 = xsp
                                    yy1 = yt
                                    xx2 = xx1 - yw
                                    yy2 = yb
                                Else
                                    xx2 = xsp - yw
                                    yy2 = yb
                                    xx1 = xr
                                    yy1 = yb - (xr - xx2)
                                End If
                            End If
                            Globals.ThisAddIn.Application.ActiveSheet.Shapes.AddLine(CSng(xx1), CSng(yy1), CSng(xx2), CSng(yy2)).Select()
                            SetLineColor()
                            xsp = xsp + length
                        End While
                    Else
                        While ysp < yb + xw
                            If ysp - yt < xw Then
                                xx2 = xl
                                yy2 = ysp
                                xx1 = xl + (yy2 - yt)
                                yy1 = yt
                            Else
                                If ysp <= yb Then
                                    xx2 = xl
                                    yy2 = ysp
                                    xx1 = xr
                                    yy1 = yy2 - xw
                                Else
                                    xx1 = xr
                                    xx2 = xl + (ysp - yb)
                                    yy2 = yb
                                    yy1 = yb - (xr - xx2)
                                End If
                            End If
                            Globals.ThisAddIn.Application.ActiveSheet.Shapes.AddLine(CSng(xx1), CSng(yy1), CSng(xx2), CSng(yy2)).Select()
                            SetLineColor()
                            ysp = ysp + length
                        End While
                    End If
                End If

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)

            End Try

        End Sub

        Public Sub CreateShapeGroup(ByVal ShapeLast As Integer, ByVal ShapeFirst As Integer)
            Try
                Dim arrShape() As String
                Dim ShapeCnt As Integer = 0
                Dim i As Integer
                If ShapeLast > ShapeFirst Then
                    ShapeCnt = 1
                    ReDim arrShape(0 To ShapeLast - ShapeFirst) 'changed to 0
                    For i = ShapeFirst + 1 To ShapeLast
                        arrShape(ShapeCnt) = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Item(i).Name
                        ShapeCnt += 1
                    Next
                    Dim shpGroup As Excel.ShapeRange = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Range(arrShape)  'grouping error in 2007 has to do with images  24-NOV-2010 ALD
                    shpGroup.Group()
                    shpGroup.Select()
                    My.Settings.Markup_LastShapeName = shpGroup.Name

                End If

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)

            End Try

        End Sub

        Public Sub SetLineColor()
            Try
                Globals.ThisAddIn.Application.Selection.ShapeRange.Line.ForeColor.RGB = My.Settings.Markup_ShapeLineColor 'RGB(0, 0, 0)

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)

            End Try

        End Sub

        Public Function SelectColor() As System.Drawing.Color
            Try
                Dim colorDlg As New ColorDialog()
                If colorDlg.ShowDialog() = DialogResult.OK Then
                    My.Settings.Markup_ShapeLineColor = colorDlg.Color
                End If
                Return My.Settings.Markup_ShapeLineColor

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)

                Return System.Drawing.Color.Black
            End Try
        End Function

        Public Sub OpenFile(ByVal FilePath As String)
            Try
                Dim pStart As New System.Diagnostics.Process
                If FilePath = String.Empty Then Exit Try
                pStart.StartInfo.FileName = FilePath
                pStart.Start()

            Catch ex As System.ComponentModel.Win32Exception
                'MessageBox.Show("No application is assicated to this file type." & vbCrLf & vbCrLf & FilePath, "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Try

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)

            End Try

        End Sub

        Public Shared Sub InvalidateRibbon()
            Try
                ribbonref.ribbon.Invalidate()

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)

            End Try

        End Sub

        Public Sub DebugPoint(x As Double, y As Double)
            '--------------------------------------------------------------------------------------------------------------------
            ' debugging routine which simply draws a rectangle centred at the specified coordinates.
            '--------------------------------------------------------------------------------------------------------------------
            Try
                Globals.ThisAddIn.Application.ActiveSheet.Rectangles.Add(x - 3, y - 3, 6, 6).Select()
                Globals.ThisAddIn.Application.Selection.Border.LineStyle = Excel.Constants.xlAutomatic
                Globals.ThisAddIn.Application.Selection.Border.Weight = Excel.XlBorderWeight.xlHairline

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)

            End Try

        End Sub

#End Region

    End Class

End Namespace