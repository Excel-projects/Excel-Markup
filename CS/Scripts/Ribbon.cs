using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace Markup.Scripts
{
    /// <summary> 
    /// Class for the ribbon procedures
    /// </summary>
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        /// <summary>
        /// Used to reference the ribbon object
        /// </summary>
        public static Ribbon ribbonref;

        /// <summary>
        /// Settings TaskPane
        /// </summary>
        public TaskPane.Settings mySettings;

        /// <summary>
        /// Settings Custom Task Pane
        /// </summary>
        public Microsoft.Office.Tools.CustomTaskPane myTaskPaneSettings;

        /// <summary> 
        /// The Clouds ribbon
        /// </summary>
        public Ribbon()
        {
        }

        #region | IRibbonExtensibility Members |

        /// <summary> 
        /// Loads the XML markup, either from an XML customization file or from XML markup embedded in the procedure, that customizes the Ribbon user interface.
        /// </summary>
        /// <param name="ribbonID">Represents the XML customization file </param>
        /// <returns>A method that returns a bitmap image for the control id. </returns> 
        /// <remarks></remarks>
        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("Markup.Ribbon.xml");
        }

        #endregion

        #region | Helpers |

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion

        #region | Ribbon Events |

        /// <summary> 
        /// loads the ribbon UI and creates a log record
        /// </summary>
        /// <param name="ribbonUI">Represents the IRibbonUI instance that is provided by the Microsoft Office application to the Ribbon extensibility code. </param>
        /// <remarks></remarks>
        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            try
            {
                this.ribbon = ribbonUI;
                ThisAddIn.e_ribbon = ribbonUI;
                Properties.Settings.Default.Markup_LastShapeName = "";
                ErrorHandler.SetLogPath();
                ErrorHandler.CreateLogRecord();
                AssemblyInfo.SetAddRemoveProgramsIcon("ExcelAddin.ico");
                System.Globalization.CultureInfo enUS = new System.Globalization.CultureInfo("en-US");
                System.Threading.Thread.CurrentThread.CurrentCulture = enUS;
            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);
            }
        }

        /// <summary> 
        /// Assigns an image to a button on the ribbon in the xml file
        /// </summary>
        /// <param name="control">Represents the object passed into the callback procedure of a control in a ribbon or another user interface that can be customized by using Office Fluent ribbon extensibility. </param>
        /// <returns>A method that returns a bitmap image for the control id. </returns> 
        public System.Drawing.Bitmap GetButtonImage(Office.IRibbonControl control)
        {
            try
            {
                switch (control.Id)
                {
                    case "grpMarkups":
                    case "btnRev":
                        return Properties.Resources.RevTri;
                    case "btnCloudAll":
                        return Properties.Resources.Cloud;
                    case "btnCloudHold":
                        return Properties.Resources.CloudHold;
                    case "btnCloudHatch":
                        return Properties.Resources.CloudHatch;
                    case "btnAreaHatch":
                        return Properties.Resources.Hatch;
                    case  "btnCloudPartLeft":
                        return Properties.Resources.CloudPartLeft;
                    case  "btnCloudPartRight":
                        return Properties.Resources.CloudPartRight;
                    case "btnCloudPartTop":
                        return Properties.Resources.CloudPartTop;
                    case "btnCloudPartBottom":
                        return Properties.Resources.CloudPartBottom;
                    default:
                        return null;
                }

            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);
                return null;

            }

        }

        /// <summary> 
        /// Assigns text to a label on the ribbon from the xml file
        /// </summary>
        /// <param name="control">Represents the object passed into the callback procedure of a control in a ribbon or another user interface that can be customized by using Office Fluent ribbon extensibility. </param>
        /// <returns>A method that returns a string for a label. </returns> 
        public string GetLabelText(Office.IRibbonControl control)
        {
            try
            {
                switch (control.Id)
                {
                    case "tabMarkup":
                        if (Application.ProductVersion.Substring(0, 2) == "15") //for Excel 2013
                        {
                            return AssemblyInfo.Title.ToUpper();
                        }
                        else
                        {
                            return AssemblyInfo.Title;
                        }
                    case "txtCopyright":
                        return "© " + AssemblyInfo.Copyright;
                    case "txtDescription":
                        return AssemblyInfo.Title.Replace("&", "&&") + " " + AssemblyInfo.AssemblyVersion;
                    case "txtReleaseDate":
                        DateTime dteCreateDate = Properties.Settings.Default.App_ReleaseDate;
                        return dteCreateDate.ToString("dd-MMM-yyyy hh:mm tt");
                    default:
                        return string.Empty;
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);
                return string.Empty;
            }
        }

        /// <summary> 
        /// Assigns the number of items for a combobox or dropdown
        /// </summary>
        /// <param name="control">Represents the object passed into the callback procedure of a control in a ribbon or another user interface that can be customized by using Office Fluent ribbon extensibility. </param>
        /// <returns>A method that returns an integer of total count of items used for a combobox or dropdown </returns> 
        public int GetItemCount(Office.IRibbonControl control)
        {
            try
            {
                switch (control.Id)
                {
                    case "drpColorType":
                        return 4;
                    default:
                        return 0;
                }

            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);
                return 0;

            }

        }

        /// <summary> 
        /// Assigns the values to a combobox or dropdown based on an index
        /// </summary>
        /// <param name="control">Represents the object passed into the callback procedure of a control in a ribbon or another user interface that can be customized by using Office Fluent ribbon extensibility. </param>
        /// <param name="index">Represents the index of the combobox or dropdown value </param>
        /// <returns>A method that returns a string per index of a combobox or dropdown </returns> 
        public string GetItemLabel(Office.IRibbonControl control, int index)
        {
            try
            {
                switch (control.Id)
                {
                    case "drpColorType":
                        switch (index)
                        {
                            case 0:
                                return "BLACK";
                            case 1:
                                return "BLUE";
                            case 2:
                                return "RED";
                            case 3:
                                return "GREEN";
                            default:
                                return string.Empty;
                        }
                    default:
                        return string.Empty;
                }

            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);
                return string.Empty;

            }

        }
                
        /// <summary> 
        /// Assigns default values to dropdowns
        /// </summary>
        /// <param name="control">Represents the object passed into the callback procedure of a control in a ribbon or another user interface that can be customized by using Office Fluent ribbon extensibility. </param>
        /// <returns>A method that returns a string for the default value of a dropdown </returns> 
        public string GetSelectedItemID(Office.IRibbonControl control)
        {
            try
            {
                Properties.Settings.Default.Markup_ShapeLineColor = System.Drawing.Color.Black;
                return control.Id; 

            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);
                return control.Id;

            }

        }

        /// <summary> 
        /// Assigns the enabled to controls
        /// </summary>
        /// <param name="control">Represents the object passed into the callback procedure of a control in a ribbon or another user interface that can be customized by using Office Fluent ribbon extensibility. </param>
        /// <returns>A method that returns true or false if the control is enabled </returns> 
        public bool GetEnabled(Office.IRibbonControl control)
        {
            try
            {
                switch (control.Id)
                {
                    case "btnRev":
                    case "btnCloudAll":
                    case "btnCloudHold":
                    case "btnCloudHatch":
                    case "btnAreaHatch":
                    case "btnCloudPartLeft":
                    case "btnCloudPartRight":
                    case "btnCloudPartTop":
                    case "btnCloudPartBottom":
                        return ErrorHandler.IsEnabled(false);
                    default:
                        return false;
                }

            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);
                return false;

            }

        }

        /// <summary>
        /// Assigns the value to an application setting
        /// </summary>
        /// <param name="control">Represents the object passed into the callback procedure of a control in a ribbon or another user interface that can be customized by using Office Fluent ribbon extensibility. </param>
        /// <returns>A method that returns true or false if the control is enabled </returns> 
        public void OnAction(Office.IRibbonControl control)
        {
            try
            {
                switch (control.Id)
                {
                    case "btnSelectColor":
                        SelectLineColor();
                        break;
                    case "btnRev":
                        CreateRevisionTriangle();
                        break;
                    case "btnCloudAll":
                        CreateCloudPart("ALL");
                        break;
                    case "btnCloudHold":
                        CreateCloudHold();
                        break;
                    case "btnCloudHatch":
                        CreateCloudHatch();
                        break;
                    case "btnAreaHatch":
                        CreateAreaHatching();
                        break;
                    case "btnCloudPartLeft":
                        CreateCloudPart("L");
                        break;
                    case "btnCloudPartRight":
                        CreateCloudPart("R");
                        break;
                    case "btnCloudPartTop":
                        CreateCloudPart("T");
                        break;
                    case "btnCloudPartBottom":
                        CreateCloudPart("B");
                        break;
                    case "btnRemoveLastShape":
                        RemoveLastShape();
                        break;
                    case "btnRemoveAllShapes":
                        RemoveAllShapes();
                        break;
                    case "btnSettings":
                        OpenSettings();
                        break;
                    case "btnOpenReadMe":
                        OpenReadMe();
                        break;
                    case "btnOpenNewIssue":
                        OpenNewIssue();
                        break;
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);
            }

        }

        /// <summary> 
        /// Preforms an action based on the index of the control
        /// </summary>
        /// <param name="control">Represents the object passed into the callback procedure of a control in a ribbon or another user interface that can be customized by using Office Fluent ribbon extensibility. </param>
        /// <param name="itemId">Represents the item identifier of the combobox </param>
        /// <param name="index">Represents the index value of the combobox </param>
        public void OnAction_Dropdown(Office.IRibbonControl control, string itemId, int index)
        {
            try
            {
                switch (control.Id)
                {
                    case "drpColorType":
                        switch (index)
                        {
                            case 0:
                                Properties.Settings.Default.Markup_ShapeLineColor = System.Drawing.Color.Black;
                                break;
                            case 1:
                                Properties.Settings.Default.Markup_ShapeLineColor = System.Drawing.Color.Blue;
                                break;
                            case 2:
                                Properties.Settings.Default.Markup_ShapeLineColor = System.Drawing.Color.Red;
                                break;
                            case 3:
                                Properties.Settings.Default.Markup_ShapeLineColor = System.Drawing.Color.Green;
                                break;
                            default:
                                Properties.Settings.Default.Markup_ShapeLineColor = System.Drawing.Color.Black;
                                break;
                        }
                        break;
                    default:
                        break;
                }

            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);
                Properties.Settings.Default.Markup_ShapeLineColor = System.Drawing.Color.Black;

            }

        }

        #endregion

        #region | Ribbon Buttons |

        /// <summary> 
        /// Creates a revision triangle and prompts the user for text value(s)
        /// </summary>
        /// <remarks></remarks>
        public void CreateRevisionTriangle()
        {
            Excel.Shape shpTriangle = null;
            Excel.Shape txtTriangle = null;
            Excel.ShapeRange shapeRange = null;
            try
            {
                if (ErrorHandler.IsEnabled(true) == false)
                {
                    return;
                }
                ErrorHandler.CreateLogRecord();
                string shapeName = "RevTri";
                Single[,] triArray = new Single[4, 2];
                double x = 0;
                double y = Globals.ThisAddIn.Application.Selection.Top;
                double h = Globals.ThisAddIn.Application.Selection.RowHeight;
                double w = Convert.ToInt32(h * 2.2 / Math.Sqrt(3));
                double f = Globals.ThisAddIn.Application.Selection.Font.Size;
                double selWidth = Globals.ThisAddIn.Application.Selection.Width;
                double selLeft = Globals.ThisAddIn.Application.Selection.Left;
                double selHorAli = Globals.ThisAddIn.Application.Selection.HorizontalAlignment;
                double xlAliCntr = Convert.ToDouble(Excel.XlHAlign.xlHAlignCenter);

                if (selHorAli == xlAliCntr & selWidth > w)
                {
                    x = selLeft + (selWidth - w) / 2;
                }
                else
                {
                    x = selLeft;
                }
                triArray[0, 0] = Convert.ToSingle(x + w / 2);
                triArray[0, 1] = Convert.ToSingle(y);
                triArray[1, 0] = Convert.ToSingle(x);
                triArray[1, 1] = Convert.ToSingle(y + h);
                triArray[2, 0] = Convert.ToSingle(x + w);
                triArray[2, 1] = Convert.ToSingle(y + h);
                triArray[3, 0] = Convert.ToSingle(x + w / 2);
                triArray[3, 1] = Convert.ToSingle(y);
                shpTriangle = Globals.ThisAddIn.Application.ActiveSheet.Shapes.AddPolyline(triArray);
                shpTriangle.Name = shapeName + AddSpaces(1) + DateTime.Now.ToString(Properties.Settings.Default.Markup_ShapeDateFormat);
                shpTriangle.Line.Weight = Convert.ToSingle(1.5);
                txtTriangle = Globals.ThisAddIn.Application.ActiveSheet.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Convert.ToSingle(x), Convert.ToSingle(y + h * 0.2), Convert.ToSingle(w), Convert.ToSingle(h * 0.8));
                txtTriangle.Select();
                string revChar = Properties.Settings.Default.Markup_RevisionTriangleCharacter;
                if (DialogBox.InputBox("Please enter a revision character", "Revision Triangle", ref revChar) == DialogResult.OK)
                {
                    Properties.Settings.Default.Markup_RevisionTriangleCharacter = revChar;
                }
                txtTriangle.TextEffect.Text = revChar;
                Globals.ThisAddIn.Application.Selection.Font.Color = Properties.Settings.Default.Markup_ShapeLineColor;
                Globals.ThisAddIn.Application.Selection.Font.Size = f;
                Globals.ThisAddIn.Application.Selection.Border.LineStyle = Excel.Constants.xlNone;
                Globals.ThisAddIn.Application.Selection.Interior.ColorIndex = Excel.Constants.xlNone;
                Globals.ThisAddIn.Application.Selection.Shadow = false;
                Globals.ThisAddIn.Application.Selection.RoundedCorners = false;
                txtTriangle.TextFrame.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                txtTriangle.TextFrame.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                txtTriangle.TextFrame.AutoSize = true;
                txtTriangle.Name = shapeName + AddSpaces(1) + DateTime.Now.ToString(Properties.Settings.Default.Markup_ShapeDateFormat);
                object[] shapes = { shpTriangle.Name, txtTriangle.Name };
                shapeRange = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Range(shapes);
                shapeRange.Group();
                shapeRange.Name = shapeName + AddSpaces(1) + DateTime.Now.ToString(Properties.Settings.Default.Markup_ShapeDateFormat);
                shpTriangle.Select();
                Globals.ThisAddIn.Application.Selection.Interior.Pattern = Excel.XlPattern.xlPatternNone;
                SetLineColor();
                Globals.ThisAddIn.Application.ActiveCell.Select();
                Properties.Settings.Default.Markup_LastShapeName = shapeRange.Name;
            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);

            }
            finally
            {
                if (shapeRange != null) Marshal.ReleaseComObject(shapeRange);
                if (txtTriangle != null) Marshal.ReleaseComObject(txtTriangle);
                if (shpTriangle != null) Marshal.ReleaseComObject(shpTriangle);
            }
        }

        /// <summary> 
        /// Creates an inverse/hold cloud object for the selected cells
        /// </summary>
        /// <remarks></remarks>
        public void CreateCloudHold()
        {
            Excel.Shape cloudLineTop = null;
            Excel.Shape cloudLineRight = null;
            Excel.Shape cloudLineBottom = null;
            Excel.Shape cloudLineLeft = null;
            Excel.ShapeRange shapeRange = null;
            try
            {
                if (ErrorHandler.IsEnabled(true) == false)
                {
                    return;
                }
                ErrorHandler.CreateLogRecord();
                string shapeName = "CloudHold";
                double x = Globals.ThisAddIn.Application.Selection.Left;
                double y = Globals.ThisAddIn.Application.Selection.Top;
                double h = Globals.ThisAddIn.Application.Selection.Height;
                double w = Globals.ThisAddIn.Application.Selection.Width;
                double off = 7.5;
                x = x - off / 2;
                w = w + off;
                y = y - off / 2;
                h = h + off;
                cloudLineTop = CreateCloudLine(x, y, x + w, y);
                cloudLineRight = CreateCloudLine(x + w, y, x + w, y + h);
                cloudLineBottom = CreateCloudLine(x + w, y + h, x, y + h);
                cloudLineLeft = CreateCloudLine(x, y + h, x, y);
                if (cloudLineBottom != null && cloudLineTop != null && cloudLineLeft != null && cloudLineRight != null) // only if there are no errors in returning an Excel shape
                {
                    object[] shapes = { cloudLineBottom.Name, cloudLineTop.Name, cloudLineLeft.Name, cloudLineRight.Name };
                    shapeRange = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Range(shapes);
                    shapeRange.Group();
                    shapeRange.Name = shapeName + AddSpaces(1) + DateTime.Now.ToString(Properties.Settings.Default.Markup_ShapeDateFormat);
                    Properties.Settings.Default.Markup_LastShapeName = shapeRange.Name;
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);

            }
            finally
            {
                if (cloudLineTop != null) Marshal.ReleaseComObject(cloudLineTop);
                if (cloudLineRight != null) Marshal.ReleaseComObject(cloudLineRight);
                if (cloudLineBottom != null) Marshal.ReleaseComObject(cloudLineBottom);
                if (cloudLineLeft != null) Marshal.ReleaseComObject(cloudLineLeft);
                if (shapeRange != null) Marshal.ReleaseComObject(shapeRange);
            }
        }

        /// <summary> 
        /// Creates a cloud object for the selected cells with hatches for demolition
        /// </summary>
        /// <remarks></remarks>
        public void CreateCloudHatch()
        {
            Excel.Shape cloudPart = null;
            Excel.Shape hatchArea = null;
            Excel.ShapeRange shapeRange = null;
            try
            {
                if (ErrorHandler.IsEnabled(true) == false)
                {
                    return;
                }
                ErrorHandler.CreateLogRecord();
                string shapeName = "CloudHatch";
                double x = Globals.ThisAddIn.Application.Selection.Left;
                double y = Globals.ThisAddIn.Application.Selection.Top;
                double h = Globals.ThisAddIn.Application.Selection.Height;
                double w = Globals.ThisAddIn.Application.Selection.Width;
                cloudPart = CreateCloudPart("ALL");
                hatchArea = CreateHatchArea(x, y, h, w);
                if (cloudPart != null && hatchArea != null) 
                {
                    object[] shapes = { cloudPart.Name, hatchArea.Name };
                    shapeRange = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Range(shapes);
                    shapeRange.Group();
                    shapeRange.Name = shapeName + AddSpaces(1) + DateTime.Now.ToString(Properties.Settings.Default.Markup_ShapeDateFormat);
                    Properties.Settings.Default.Markup_LastShapeName = shapeRange.Name;
                    Marshal.FinalReleaseComObject(cloudPart);
                    Marshal.FinalReleaseComObject(hatchArea);
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);

            }
            finally
            {
                if (cloudPart != null) Marshal.ReleaseComObject(cloudPart);
                if (hatchArea != null) Marshal.ReleaseComObject(hatchArea);
                if (shapeRange != null) Marshal.ReleaseComObject(shapeRange);
            }
        }

        /// <summary> 
        /// Creates demolition hatching for selected cells
        /// </summary>
        /// <remarks></remarks>
        public void CreateAreaHatching()
        {
            try
            {
                if (ErrorHandler.IsEnabled(true) == false)
                {
                    return;
                }
                ErrorHandler.CreateLogRecord();
                int selectAreaCnt = Globals.ThisAddIn.Application.Selection.Areas.Count;
                if (selectAreaCnt > 1)
                {
                    foreach (Excel.Range singleArea in Globals.ThisAddIn.Application.Selection.Areas)
                    {
                        CreateHatchArea(singleArea.Left, singleArea.Top, singleArea.Height, singleArea.Width);
                    }
                }
                else
                {
                    CreateHatchArea(Globals.ThisAddIn.Application.Selection.Left, Globals.ThisAddIn.Application.Selection.Top, Globals.ThisAddIn.Application.Selection.Height, Globals.ThisAddIn.Application.Selection.Width);
                }

            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);

            }
        }

        /// <summary> 
        /// Select the color for the line from a dialog box
        /// </summary>
        /// <remarks></remarks>
        public void SelectLineColor()
        {
            ErrorHandler.CreateLogRecord();
            Properties.Settings.Default.Markup_ShapeLineColor = SelectColor();
        }

        /// <summary> 
        /// Remove the last shape created from the current session
        /// </summary>
        /// <remarks></remarks>
        public void RemoveLastShape()
        {
            Excel.Worksheet xlWorkSheet = null;
            try
            {
                ErrorHandler.CreateLogRecord();
                if (ErrorHandler.IsActiveDocument(true) == false)
                {
                    return;
                }
                DialogResult dialogResult = MessageBox.Show("Are you sure you would like to delete the last shape that has been created?", "Delete Last Shape?", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (dialogResult == DialogResult.No)
                {
                    return;
                }
                else if (dialogResult == DialogResult.Yes)
                {
                    xlWorkSheet = Globals.ThisAddIn.Application.ActiveSheet;
                    string strShapeName = Properties.Settings.Default.Markup_LastShapeName;
                    foreach (Excel.Shape shp in xlWorkSheet.Shapes)
                    {
                        if (shp.Name == strShapeName)
                        {
                            shp.Delete();
                        }
                        if (shp != null) Marshal.ReleaseComObject(shp);
                    }
                    if (xlWorkSheet != null) Marshal.ReleaseComObject(xlWorkSheet);
                }
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                MessageBox.Show("You are currently editing a cell." + Environment.NewLine + Environment.NewLine + "Please finish editing and press return.", "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);

            }
            finally
            {
                if (xlWorkSheet != null) Marshal.ReleaseComObject(xlWorkSheet);
            }
        }

        /// <summary> 
        /// Remove all shapes from the current sheet
        /// </summary>
        /// <remarks></remarks>
        public void RemoveAllShapes()
        {
            Excel.Worksheet xlWorkSheet = null;
            try
            {

                ErrorHandler.CreateLogRecord();
                if (ErrorHandler.IsActiveDocument(true) == false)
                {
                    return;
                }
                DialogResult dialogResult = MessageBox.Show("Are you sure you would like to delete all the shapes in the active worksheet?", "Delete All Shapes?", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (dialogResult == DialogResult.No)
                {
                    return;
                }
                else if (dialogResult == DialogResult.Yes)
                {
                    xlWorkSheet = Globals.ThisAddIn.Application.ActiveSheet;
                    foreach (Excel.Shape shp in xlWorkSheet.Shapes)
                    {
                        if (shp.Type == Microsoft.Office.Core.MsoShapeType.msoGroup || shp.Type == Microsoft.Office.Core.MsoShapeType.msoLine || shp.Type == Microsoft.Office.Core.MsoShapeType.msoFreeform)
                        {
                            string s = shp.Name;
                            if (s.Contains("RevTri") || s.Contains("Cloud") || s.Contains("AreaHatch"))
                            {
                                shp.Delete();
                            }
                        }
                        if (shp != null) Marshal.ReleaseComObject(shp);
                    }
                    if (xlWorkSheet != null) Marshal.ReleaseComObject(xlWorkSheet);
                }
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                MessageBox.Show("You are currently editing a cell." + Environment.NewLine + Environment.NewLine + "Please finish editing and press return.", "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);

            }
            finally
            {
                if (xlWorkSheet != null) Marshal.ReleaseComObject(xlWorkSheet);
            }

        }

        /// <summary> 
        /// Opens the settings form
        /// </summary>
        /// <remarks></remarks>
        public void OpenSettings()
        {
            try
            {
                if (myTaskPaneSettings != null)
                {
                    if (myTaskPaneSettings.Visible == true)
                    {
                        myTaskPaneSettings.Visible = false;
                    }
                    else
                    {
                        myTaskPaneSettings.Visible = true;
                    }
                }
                else
                {
                    mySettings = new TaskPane.Settings();
                    myTaskPaneSettings = Globals.ThisAddIn.CustomTaskPanes.Add(mySettings, "Settings for " + Scripts.AssemblyInfo.Title);
                    myTaskPaneSettings.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight;
                    myTaskPaneSettings.DockPositionRestrict = Office.MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoChange;
                    myTaskPaneSettings.Width = 675;
                    myTaskPaneSettings.Visible = true;
                }

            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);
            }
        }

        /// <summary> 
        /// Opens an as built help file
        /// </summary>
        /// <remarks></remarks>
        public void OpenReadMe()
        {
            ErrorHandler.CreateLogRecord();
            System.Diagnostics.Process.Start(Properties.Settings.Default.App_PathReadMe);

        }

        /// <summary> 
        /// Opens an as built help file
        /// </summary>
        /// <remarks></remarks>
        public void OpenNewIssue()
        {
            ErrorHandler.CreateLogRecord();
            System.Diagnostics.Process.Start(Properties.Settings.Default.App_PathNewIssue);

        }

        #endregion

        #region | Subroutines |

        /// <summary> 
        /// Creates an arc based on user selection
        /// </summary>
        /// <param name="x1">X Axis 1 </param>
        /// <param name="y1">Y Axis 1 </param>
        /// <param name="x2">X Axis 2 </param>
        /// <param name="y2">Y Axis 2 </param>
        /// <param name="length">description here... </param>
        /// <returns>A method that creates an arc shape between 2 sets for coordinates </returns> 
        /// <remarks></remarks>
        public Excel.Shape CreateArc(double x1, double y1, double x2, double y2, double length)
        {
            Excel.Shape cloudArc = null;
            try
            {
                string shapeName = "CloudArc";
                int i = 0;
                double angle = 60;
                double segments = angle / 10;
                float[,] arcArray = new float[Convert.ToInt32(segments) + 1, 2];
                double theta = angle * Math.PI / 180;
                double xm = (x1 + x2) / 2;
                double ym = (y1 + y2) / 2;
                double xd = (x2 - x1);
                double yd = (y2 - y1);
                double d = Math.Sqrt(xd * xd + yd * yd);
                double r = d / 2 / Math.Sin(theta / 2);
                double xc = xm + yd / (2 * Math.Tan(theta / 2));
                double yc = ym - xd / (2 * Math.Tan(theta / 2));
                double dtheta = theta / segments;
                arcArray[0, 0] = Convert.ToSingle(x1);
                arcArray[0, 1] = Convert.ToSingle(y1);
                double a = Math.Atan2(y1 - yc, x1 - xc) - dtheta;
                for (i = 1; i <= Convert.ToInt32(segments) - 1; i++)
                {
                    arcArray[i, 0] = Convert.ToSingle(xc + r * Math.Cos(a));
                    arcArray[i, 1] = Convert.ToSingle(yc + r * Math.Sin(a));
                    a = a - dtheta;
                }
                arcArray[i, 0] = Convert.ToSingle(x2);
                arcArray[i, 1] = Convert.ToSingle(y2);
                cloudArc = Globals.ThisAddIn.Application.ActiveSheet.Shapes.AddPolyline(arcArray);
                cloudArc.Select();
                cloudArc.Name = shapeName + AddSpaces(1) + DateTime.Now.ToString(Properties.Settings.Default.Markup_ShapeDateFormat);
                Globals.ThisAddIn.Application.Selection.Interior.Pattern = Excel.Constants.xlNone;
                SetLineColor();
                return cloudArc;

            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);
                return null;

            }
        }

        /// <summary> 
        /// Draws the side of a cloud between the specified coordinates.
        /// </summary>
        /// <param name="x1">X Axis 1 </param>
        /// <param name="y1">Y Axis 1 </param>
        /// <param name="x2">X Axis 2 </param>
        /// <param name="y2">Y Axis 2 </param>
        /// <returns>A method that creates series of arc shapes between 2 sets for coordinates </returns> 
        /// <remarks></remarks>
        public Excel.Shape CreateCloudLine(double x1, double y1, double x2, double y2)
        {
            Excel.Shape cloudArc = null;
            Excel.Shape cloudLine = null;
            Excel.ShapeRange shapeRange = null;
            try
            {
                double length = 25;
                int i = 0;
                double x = 0;
                double y = 0;
                double dx = x2 - x1;
                double dy = y2 - y1;
                double d = Math.Sqrt(dx * dx + dy * dy);
                double segments = Math.Ceiling(d / length);
                if (segments < 2)
                    segments = 2;
                double deltax = (dx / segments);
                double deltay = (dy / segments);
                double xp = x1;
                double yp = y1;
                string shapeName = "CloudLine";
                object[] shapes = new object[Convert.ToInt32(segments)];
                for (i = 1; i <= Convert.ToInt32(segments); i++)
                {
                    x = xp + deltax;
                    y = yp + deltay;
                    cloudArc = CreateArc(xp, yp, x, y, length);
                    shapes[i - 1] = cloudArc.Name; 
                    xp = x;
                    yp = y;
                }
                shapeRange = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Range(shapes);
                shapeRange.Group();
                shapeRange.Name = shapeName + AddSpaces(1) + DateTime.Now.ToString(Properties.Settings.Default.Markup_ShapeDateFormat);
                cloudLine = Globals.ThisAddIn.Application.ActiveSheet.Shapes(shapeRange.Name);
                Properties.Settings.Default.Markup_LastShapeName = shapeRange.Name;
                return cloudLine;
            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex, true);
                return null;

            }
        }

        /// <summary> 
        /// Creates a part of the cloud for the selected cells (B=Bottom, T=Top, L=Left, R=Right, ALL=all sides of cloud)
        /// </summary>
        /// <param name="cloudPart">The name of the side of the cloud part </param>
        /// <returns>the grouped shape </returns> 
        /// <remarks></remarks>
        public Excel.Shape CreateCloudPart(string cloudPart)
        {
            if (ErrorHandler.IsEnabled(true) == false)
            {
                return null;
            }
            Excel.Shape cloudLineBottom = null;
            Excel.Shape cloudLineTop = null;
            Excel.Shape cloudLineLeft = null;
            Excel.Shape cloudLineRight = null;
            Excel.Shape cloudLine = null;
            Excel.ShapeRange shapeRange = null;
            try
            {
                double x = Globals.ThisAddIn.Application.Selection.Left;
                double y = Globals.ThisAddIn.Application.Selection.Top;
                double h = Globals.ThisAddIn.Application.Selection.Height;
                double w = Globals.ThisAddIn.Application.Selection.Width;

                if (cloudPart == "B" | cloudPart == "ALL")
                {
                    cloudLineBottom = CreateCloudLine(x, y + h, x + w, y + h);
                }
                if (cloudPart == "T" | cloudPart == "ALL")
                {
                    cloudLineTop = CreateCloudLine(x + w, y, x, y);
                }
                if (cloudPart == "L" | cloudPart == "ALL")
                {
                    cloudLineLeft = CreateCloudLine(x, y, x, y + h);
                }
                if (cloudPart == "R" | cloudPart == "ALL")
                {
                    cloudLineRight = CreateCloudLine(x + w, y + h, x + w, y);
                }

                if (cloudPart == "ALL" && cloudLineBottom != null && cloudLineTop != null && cloudLineLeft != null && cloudLineRight != null)
                {
                    string shapeName = "Cloud";
                    object[] shapes = { cloudLineBottom.Name, cloudLineTop.Name, cloudLineLeft.Name, cloudLineRight.Name };
                    shapeRange = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Range(shapes);
                    shapeRange.Group();
                    shapeRange.Name = shapeName + AddSpaces(1) + DateTime.Now.ToString(Properties.Settings.Default.Markup_ShapeDateFormat);
                    cloudLine = Globals.ThisAddIn.Application.ActiveSheet.Shapes(shapeRange.Name);
                    Properties.Settings.Default.Markup_LastShapeName = shapeRange.Name;
                    return cloudLine;
                }
                else
                {
                    return null;
                }

            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);
                return null;

            }
            finally
            {
                if (cloudLineBottom != null) Marshal.ReleaseComObject(cloudLineBottom);
                if (cloudLineTop != null) Marshal.ReleaseComObject(cloudLineTop);
                if (cloudLineLeft != null) Marshal.ReleaseComObject(cloudLineLeft);
                if (cloudLineRight != null) Marshal.ReleaseComObject(cloudLineRight);
                if (cloudLine != null) Marshal.ReleaseComObject(cloudLine);
                if (shapeRange != null) Marshal.ReleaseComObject(shapeRange);
            }

        }

        /// <summary> 
        /// Creates a hatch area for the selected cells
        /// </summary>
        /// <param name="x">X Axis </param>
        /// <param name="y">Y Axis </param>
        /// <param name="h">Height </param>
        /// <param name="w">Width </param>
        /// <returns>A method that creates series of line shapes between within an area of 2 sets for coordinates </returns> 
        /// <remarks></remarks>
        public Excel.Shape CreateHatchArea(double x, double y, double h, double w)
        {
            string shapeName = "AreaHatch";
            double xx1 = 0;
            double yy1 = 0;
            double xx2 = 0;
            double yy2 = 0;
            double x1 = x + y;
            double x2 = Math.Floor(x1 / 20);
            double x3 = (x2 + 1) * 20;
            double xDiff = x3 - x1;
            double xl = x;
            double xr = x + w;
            double xw = w;
            double yt = y;
            double yb = y + h;
            double yw = h;
            double xsp = x + xDiff;
            double ysp = yt + xDiff;
            Excel.Shape hatchLine1 = null;
            Excel.Shape hatchLine2 = null;
            Excel.Shape hatchArea = null;
            Excel.ShapeRange shapeRange = null;
            List<object> shapesList = new List<object>();
            try
            {
                if (xw >= yw)
                {
                    while (xsp < xr + yw)
                    {
                        if (xsp - xl < yb - yt)
                        {
                            xx1 = xsp;
                            yy1 = yt;
                            xx2 = xl;
                            yy2 = yt + (xx1 - xl);
                        }
                        else
                        {
                            if (xsp <= xr)
                            {
                                xx1 = xsp;
                                yy1 = yt;
                                xx2 = xx1 - yw;
                                yy2 = yb;
                            }
                            else
                            {
                                xx2 = xsp - yw;
                                yy2 = yb;
                                xx1 = xr;
                                yy1 = yb - (xr - xx2);
                            }
                        }
                        hatchLine1 = Globals.ThisAddIn.Application.ActiveSheet.Shapes.AddLine(Convert.ToSingle(xx1), Convert.ToSingle(yy1), Convert.ToSingle(xx2), Convert.ToSingle(yy2));
                        hatchLine1.Select();
                        hatchLine1.Name = shapeName + AddSpaces(1) + DateTime.Now.ToString(Properties.Settings.Default.Markup_ShapeDateFormat);
                        shapesList.Add(hatchLine1.Name);
                        SetLineColor();
                        xsp = xsp + 20;
                    }
                }
                else
                {
                    while (ysp < yb + xw)
                    {
                        if (ysp - yt < xw)
                        {
                            xx2 = xl;
                            yy2 = ysp;
                            xx1 = xl + (yy2 - yt);
                            yy1 = yt;
                        }
                        else
                        {
                            if (ysp <= yb)
                            {
                                xx2 = xl;
                                yy2 = ysp;
                                xx1 = xr;
                                yy1 = yy2 - xw;
                            }
                            else
                            {
                                xx1 = xr;
                                xx2 = xl + (ysp - yb);
                                yy2 = yb;
                                yy1 = yb - (xr - xx2);
                            }
                        }
                        hatchLine2 = Globals.ThisAddIn.Application.ActiveSheet.Shapes.AddLine(Convert.ToSingle(xx1), Convert.ToSingle(yy1), Convert.ToSingle(xx2), Convert.ToSingle(yy2));
                        hatchLine2.Select();
                        hatchLine2.Name = shapeName + AddSpaces(1) + DateTime.Now.ToString(Properties.Settings.Default.Markup_ShapeDateFormat);
                        shapesList.Add(hatchLine2.Name);
                        SetLineColor();
                        ysp = ysp + 20;
                    }
                }
                object[] shapes = shapesList.ToArray();
                shapeRange = Globals.ThisAddIn.Application.ActiveSheet.Shapes.Range(shapes);
                shapeRange.Group();
                shapeRange.Name = shapeName + AddSpaces(1) + DateTime.Now.ToString(Properties.Settings.Default.Markup_ShapeDateFormat);
                hatchArea = Globals.ThisAddIn.Application.ActiveSheet.Shapes(shapeRange.Name);
                Properties.Settings.Default.Markup_LastShapeName = shapeRange.Name;
                return hatchArea;
            }

            catch (System.Runtime.InteropServices.COMException ex)
            {
                ErrorHandler.DisplayMessage(ex, true);
                return null;

            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex, true);
                return null;

            }
            finally
            {
                if (shapeRange != null) Marshal.FinalReleaseComObject(shapeRange);
                if (hatchLine1 != null) Marshal.FinalReleaseComObject(hatchLine1);
                if (hatchLine2 != null) Marshal.FinalReleaseComObject(hatchLine2);

            }
        }

        /// <summary>
        /// Set the shape line color for the clouds
        /// </summary>
        public void SetLineColor()
        {
            try
            {
                Globals.ThisAddIn.Application.Selection.ShapeRange.Line.Visible = true;
                Globals.ThisAddIn.Application.Selection.ShapeRange.Line.ForeColor.RGB = Properties.Settings.Default.Markup_ShapeLineColor;
            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);

            }

        }

        /// <summary> 
        /// Select a color from a dialog box
        /// </summary>
        /// <returns>The selected color from the dialog box </returns> 
        /// <remarks></remarks>
        public System.Drawing.Color SelectColor()
        {
            try
            {
                ColorDialog colorDlg = new ColorDialog();
                if (colorDlg.ShowDialog() == DialogResult.OK)
                {
                    Properties.Settings.Default.Markup_ShapeLineColor = colorDlg.Color;
                }
                if (ErrorHandler.IsActiveSelection(false) == false)
                {
                    Globals.ThisAddIn.Application.Selection.ShapeRange.Line.ForeColor.RGB = Properties.Settings.Default.Markup_ShapeLineColor;
                }
                return Properties.Settings.Default.Markup_ShapeLineColor;
            }

            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);
                return System.Drawing.Color.Black;

            }
        }

        /// <summary> 
        /// Creates x number of spaces to use in string variables
        /// </summary>
        /// <param name="numberOfSpaces">Represents the number of spaces to add to a string </param>
        /// <returns>A method that returns a string of spaces </returns> 
        /// <remarks></remarks>
        public string AddSpaces(int numberOfSpaces = 1)
        {
            try
            {
                string myString = string.Empty;
                myString = myString.PadRight(numberOfSpaces);
                return myString;
            }

            catch (Exception)
            {
                return " ";

            }
        }
        
        /// <summary>
        /// Used to update/reset the ribbon values
        /// </summary>
        public void InvalidateRibbon()
        {
            ribbon.Invalidate();
        }

        #endregion

    }
}