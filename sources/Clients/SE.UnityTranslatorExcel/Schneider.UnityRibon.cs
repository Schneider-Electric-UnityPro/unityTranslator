using System;
using System.Collections.Generic;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.Excel;
using System.Threading.Tasks;
using log4net;
using System.Diagnostics;

namespace SE.UnityCommentsExcel
{
    public partial class UnityCommentRibbon
    {
        #region field
        private object _locker = new object();
        private ExcelViewModel _viewModel = null;
        private string _wokingSheetName = "Unity comments translator";
        private static ILog _log;
        #endregion

        #region helpers



        /// <summary>
        /// waiting cursor handler
        /// </summary>
        private class ExcelWaitCursor : IDisposable
        {

            Application _excel;
            XlMousePointer _cursor;

            internal ExcelWaitCursor(Application excel)
            {
                //forces the wait cursor
                _excel = excel;
                _cursor = excel.Cursor;
                _excel.Cursor = XlMousePointer.xlWait;
            }

            public void Dispose()
            {
                //restore
                _excel.Cursor = _cursor;
            }
        }


        /// <summary>
        /// lock screen update handler
        /// </summary>
        private class ExcelScreenUpdater : IDisposable
        {

            Application _excel;

            internal ExcelScreenUpdater(Application excel)
            {
                //forces the wait cursor
                _excel = excel;
                excel.ScreenUpdating = false;
            }

            public void Dispose()
            {
                //restore
                _excel.ScreenUpdating = true;
            }
        }



        /// <summary>
        /// lock event handler
        /// </summary>
        private class ExcelEventLocker : IDisposable
        {

            Application _excel;

            internal ExcelEventLocker(Application excel)
            {
                //forces the wait cursor
                _excel = excel;
                excel.EnableEvents = false;
            }

            public void Dispose()
            {
                //restore
                _excel.EnableEvents = true;
            }
        }
        #endregion

        #region properties
        /// <summary>
        /// Gets or sets the logger.
        /// </summary>
        /// <value>
        /// The log.
        /// </value>
        public static ILog Log
        {
            get
            {
                return _log;
            }
            set
            {
                //propagate
                ExcelViewModel.Log = _log = value;
            }
        }


        /// <summary>
        /// Gets or sets the view model.
        /// </summary>
        /// <value>
        /// The view model.
        /// </value>
        private ExcelViewModel ViewModel
        {
            get
            {
                lock (_locker)
                {
                    return _viewModel;
                }
            }

            set
            {
                lock (_locker)
                {
                    if (_viewModel != null)
                    {
                        _viewModel.Dispose();
                    }
                    _viewModel = value;
                }
            }
        }
        #endregion

        /// <summary>
        /// Called when [load].
        /// </summary>
        /// <param name="sender">The sender.</param>
        /// <param name="e">The <see cref="RibbonUIEventArgs"/> instance containing the event data.</param>
        private void OnLoad(object sender, RibbonUIEventArgs e)
        {
            Log?.Info($"Schneider Electric comment Addin loaded in excel");
            ViewModel = new ExcelViewModel();
            
            TargetLang.ItemsLoading += TargetLang_ItemsLoading;
        }

        private void TargetLang_ItemsLoading(object sender, RibbonControlEventArgs e)
        {
           
        }

        /// <summary>
        /// Called when [open] unity
        /// </summary>
        /// <param name="sender">The sender.</param>
        /// <param name="e">The <see cref="RibbonControlEventArgs"/> instance containing the event data.</param>
        private void OnOpen(object sender, RibbonControlEventArgs e)
        {
            var ws = InitialiseWorksheet(ExcelFromArgs(e));
            if (ws != null)
            {
                System.Windows.Forms.OpenFileDialog openFileDialog = new System.Windows.Forms.OpenFileDialog();
                openFileDialog.DefaultExt = ".zef";
                openFileDialog.Multiselect = false;
                openFileDialog.Filter = "Unity export Files (*.zef)|*.zef|Unity projects (.stu)|*.stu|Unity Archives Files (*.sta)|*.sta";
                if (openFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    var file = openFileDialog.FileName;                
                    Task.Run(async () =>
                    {
                        Log?.Info($"Open unity application into worksheet ({file})");
                        using (ExcelWaitCursor cursor = new ExcelWaitCursor(ws.Application))
                        {
                            await ViewModel.Open(file);
                            UpdateExcelSheetFromTable(ws);
                        }
                    });
                    
                }
            }
        }

        /// Excels from arguments.
        /// </summary>
        /// <param name="e">The <see cref="Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs" /> instance containing the event data.</param>
        /// <returns></returns>
        private static Application ExcelFromArgs(RibbonControlEventArgs e)
        {
            return e?.Control?.Context?.Application;
        }

        /// <summary>
        /// Called when [save].
        /// </summary>
        /// <param name="sender">The sender.</param>
        /// <param name="e">The <see cref="RibbonControlEventArgs"/> instance containing the event data.</param>
        private void OnSave(object sender, RibbonControlEventArgs e)
        {
            System.Windows.Forms.SaveFileDialog filedlg = new System.Windows.Forms.SaveFileDialog();
            filedlg.DefaultExt = ".zef";
            filedlg.Filter = "unity export format (.zef)|*.zef";
            if (filedlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                var ws = SelectWorkSheet(ExcelFromArgs(e));
                if (ws != null)
                {
                    Task.Run(async () =>
                    {
                        Log?.Info($"Save worksheet ({filedlg.FileName})");
                        using (ExcelWaitCursor cursor = new ExcelWaitCursor(ws.Application))
                        {
                            UpdateTableFromExcel(ws);
                            await ViewModel.Save(filedlg.FileName);
                        }
                    });
                }
            }
        }

        /// <summary>
        /// Called when [import].
        /// </summary>
        /// <param name="sender">The sender.</param>
        /// <param name="e">The <see cref="RibbonControlEventArgs"/> instance containing the event data.</param>
        private void OnImport(object sender, RibbonControlEventArgs e)
        {
            var ws = InitialiseWorksheet(ExcelFromArgs(e));
            if (ws != null)
            {
                System.Windows.Forms.OpenFileDialog openFileDialog = new System.Windows.Forms.OpenFileDialog();
                openFileDialog.DefaultExt = ".xml";
                openFileDialog.Filter = "Unity Comment translation projects (.xml)|*.xml";
                if (openFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    Task.Run(async () =>
                    {
                        Log?.Info($"Load worksheet ({openFileDialog.FileName})");
                        using (ExcelWaitCursor cursor = new ExcelWaitCursor(ws.Application))
                        {
                            await ViewModel.Import(openFileDialog.FileName);
                            UpdateExcelSheetFromTable(ws);
                        }
                    });
                }
            }
        }

        /// <summary>
        /// Called when [export].
        /// </summary>
        /// <param name="sender">The sender.</param>
        /// <param name="e">The <see cref="RibbonControlEventArgs"/> instance containing the event data.</param>
        private void OnExport(object sender, RibbonControlEventArgs e)
        {
            System.Windows.Forms.SaveFileDialog filedlg = new System.Windows.Forms.SaveFileDialog();
            filedlg.DefaultExt = ".xml";
            filedlg.Filter = "Unity Comment translation projects (.xml)|*.xml";
            if (filedlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                var ws = SelectWorkSheet(ExcelFromArgs(e));
                if (ws != null)
                {
                    Task.Run(async () =>
                    {
                        Log.Info($"Save worksheet ({filedlg.FileName})");
                        using (ExcelWaitCursor cursor = new ExcelWaitCursor(ws.Application))
                        {
                            UpdateTableFromExcel(ws);
                            await ViewModel.Export(filedlg.FileName);
                        }
                    });
                }
            }
        }

        /// <summary>
        /// Called when [translate].
        /// </summary>
        /// <param name="sender">The sender.</param>
        /// <param name="e">The <see cref="RibbonControlEventArgs"/> instance containing the event data.</param>
        private void OnTranslate(object sender, RibbonControlEventArgs e)
        {
            var app = ExcelFromArgs(e);
            var ws = SelectWorkSheet(app);
            Range selection = null;
            if (ExcelViewModel.TranslateSelectionOnly)
            {
                selection = app?.Selection;
            }
            Task.Run(async () =>
            {
                List<string> Ids = null;
                using (ExcelWaitCursor cursor = new ExcelWaitCursor(ws.Application))
                {
                    if (selection != null)
                    {
                        Ids = new List<string>();
                        foreach (Range row in selection.Rows)
                        {
                            var array = RowData(row);
                            Ids.Add((string)array[ColumnsProperties.Index(ColumnsProperties.ID)]);
                        }
                        //var translations= await ViewModel.PartialTranslate(Ids);
                    }
                    UpdateTableFromExcel(ws);
                    await ViewModel.Translate(Ids);
                    UpdateExcelSheetFromTable(ws);
                    
                }
            });
        }


        /// <summary>
        /// Called when [target languagechange].
        /// </summary>
        /// <param name="sender">The sender.</param>
        /// <param name="e">The <see cref="Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs" /> instance containing the event data.</param>
        private void OnTargetLanguagechange(object sender, RibbonControlEventArgs e)
        {
            RibbonComboBox drop = (RibbonComboBox)sender;
            if (ViewModel != null)
            {
                var lang = drop.Text;
                ExcelViewModel.SetTargetLang(lang);
            }
        }


        /// <summary>
        /// Called when [filter selection change].
        /// </summary>
        /// <param name="sender">The sender.</param>
        /// <param name="e">The <see cref="Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs" /> instance containing the event data.</param>
        private void OnFilterSelectionChange(object sender, RibbonControlEventArgs e)
        {
            RibbonCheckBox check = (RibbonCheckBox)sender;
            ExcelViewModel.TranslateSelectionOnly = check.Checked;
        }

        /// <summary>
        /// Called when [erase mode change].
        /// </summary>
        /// <param name="sender">The sender.</param>
        /// <param name="e">The <see cref="Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs" /> instance containing the event data.</param>
        private void OnEraseModeChange(object sender, RibbonControlEventArgs e)
        {
            RibbonCheckBox check = (RibbonCheckBox)sender;
            ExcelViewModel.EraseExistingTranslations = check.Checked;
        }

        #region Table



        /// <summary>
        /// Fills the table from the sheet.
        /// </summary>
        /// <param name="excelSheet">The excel sheet.</param>
        private void UpdateTableFromExcel(Worksheet excelSheet)
        {
            //populate data content row by row         
            CommentTable dataTable = new CommentTable();
            int rows = ViewModel.Table.Rows.Count;
            Range r = excelSheet.Rows.Resize[rows, ColumnsProperties.NbCols].Offset[1];
            foreach (Range row in r)
            {
                AddRow(RowData(row), dataTable);
            }

            Log?.Info($"Table updated : {rows} rows");
            ViewModel.Table = dataTable;
        }

        /// <summary>
        /// add a row to the table.
        /// </summary>
        /// <param name="data">The data.</param>
        /// <param name="dataTable">The data table.</param>
        private void AddRow(object[] data, CommentTable dataTable)
        {
            if (data[0] != null)
            {
                //at least application is not null
                var tableRow = dataTable.CreateNewRow();
                tableRow.ItemArray = data;
                dataTable.Rows.Add(tableRow);
            }
        }

        /// <summary>
        /// Fills the sheet from the table.
        /// </summary>
        /// <param name="excelSheet">The excel sheet.</param>
        private void UpdateExcelSheetFromTable(Worksheet excelSheet)
        {
            Stopwatch s = Stopwatch.StartNew();
            System.Data.DataTable dataTable = ViewModel.Table;

            using (var evtlock = new ExcelEventLocker(excelSheet.Application))
            {
                // loop through each row and add values to our sheet

                if (dataTable != null)
                {
                    using (var displaylock = new ExcelScreenUpdater(excelSheet.Application))
                    {
                        FillExcelSheet(excelSheet, dataTable);
                        //calculate util range
                        Range target = excelSheet.UsedRange;


                        if (dataTable.Rows.Count < 2)
                        {
                            target = excelSheet.Cells.Resize[5, ColumnsProperties.NbCols];
                        }
                        else
                        {
                            target = target.Resize[dataTable.Rows.Count, ColumnsProperties.NbCols];
                        }


                        //column resizes
                        target.EntireColumn.AutoFit();

                        //hide must be redo after resize (autofit)
                        ColumnRange(target.Columns, ColumnsProperties.ID, false).Hidden = true;

                        //update look and feel and set un conditional formating
                        FormatContent(target);
                    }

                }
            }
            s.Stop();
            Log?.Info($"Excel Sheet updated : {dataTable?.Rows.Count} rows in {s.ElapsedMilliseconds} ms");

        }

        /// <summary>
        ///  objects array == values of the row.
        /// </summary>
        /// <param name="row">The row.</param>
        /// <returns></returns>
        private static object[] RowData(Range row)
        {
            List<object> d = new List<object>();
            foreach (Range x in row.Resize[1, ColumnsProperties.NbCols].Cells)
            {
                d.Add(x.Value);
            }
            return d.ToArray();
        }

        /// <summary>
        /// Fills the specified excel sheet with table content
        /// </summary>
        /// <param name="excelSheet">The excel sheet.</param>
        /// <param name="dataTable">The data table.</param>
        private static void FillExcelSheet(Worksheet excelSheet, System.Data.DataTable dataTable)
        {
            Stopwatch s = Stopwatch.StartNew();

            if (dataTable.Rows.Count > 0)
            {
                // populate data content row by row
                Range tableArea = excelSheet.Range["A2"].Resize[dataTable.Rows.Count, dataTable.Columns.Count];
                for (int Idx = 0; Idx < dataTable.Rows.Count; Idx++)
                {
                    CommentRow row = (CommentRow)dataTable.Rows[Idx];
                    var excelrow = tableArea.Rows[Idx + 1];
                    excelrow.Value = row.ItemArray;
                }

                s.Stop();
                Log?.Info($"Excel sheet filling in {s.ElapsedMilliseconds} ms");
            }
            else
            {
                s.Stop();
                Log?.Warn($"No data found . Empty excel sheet {s.ElapsedMilliseconds} ms");
            }

        }

        

        #endregion

        #region Excel Worksheet


        /// <summary>
        /// Formats the colums.
        /// </summary>
        /// <param name="ws">The ws.</param>
        private void FormatColums(Worksheet ws)
        {
            if (ws != null)
            {
                if (ViewModel == null)
                {
                    ViewModel = new ExcelViewModel();
                }
                var dataTable = ViewModel.Table;
                for (int i = 0; i < dataTable.Columns.Count; i++)
                {
                    Range cell = ws.Cells[1, i + 1];
                    cell.Value = dataTable.Columns[i].ColumnName;
                    SetColumnType(ws, cell.EntireColumn, dataTable.Columns[i].DataType);
                }
                //default
                ws.Cells.Font.Color = System.Drawing.Color.Black;

                var header = ws.Range["A1"].Resize[1, ColumnsProperties.NbCols];
                header.Font.Bold = true;
                header.Interior.Color = XlRgbColor.rgbGreen;
                header.Font.Color = XlRgbColor.rgbGhostWhite;

                //freeze
                FreezeHeader(ws);

            }
        }

        /// <summary>
        /// Initialises the worksheet.
        /// </summary>
        /// <param name="e">The <see cref="RibbonControlEventArgs"/> instance containing the event data.</param>
        /// <returns></returns>
        private Worksheet InitialiseWorksheet(Application excel)
        {
            Worksheet ws = SelectWorkSheet(excel);
            ViewModel = new ExcelViewModel();
            if (ws != null)
            {
                using (var evtlock = new ExcelEventLocker(ws.Application))
                {
                    ws.UsedRange.Offset[1].ClearContents();
                    ws.UsedRange.Offset[1].ClearFormats();
                }
            }
            else
            {
                ws = CreateWorkSheet(excel);
                //makes sure event are enabled
                ws.Application.EnableEvents = true;
            }
            ws.Activate();
            return ws;
        }

        /// <summary>
        /// Selects the work sheet.
        /// </summary>
        /// <param name="excel">The excel.</param>
        /// <returns></returns>
        private Worksheet SelectWorkSheet(Application excel)
        {
            Worksheet sheet = null;
            try
            {
                sheet = ((Worksheet)excel?.Worksheets[_wokingSheetName]);
            }
            catch
            {

            }
            return sheet;
        }

        /// <summary>
        /// Actives the sheet.
        /// </summary>
        /// <param name="excel">The excel.</param>
        /// <returns></returns>
        private Worksheet CreateWorkSheet(Application excel)
        {
            Worksheet sheet = null;
            try
            {
                sheet = ((Worksheet)excel.Worksheets.Add());
                if (sheet != null)
                {
                    sheet.Name = _wokingSheetName;
                    FormatColums(sheet);
                }
            }
            catch
            {

            }
            return sheet;
        }

        
        #endregion

        #region Excel Formating
        /// <summary>
        /// Formats the sheet.
        /// </summary>
        /// <param name="target">The excel sheet.</param>
        /// <param name="rows">The rows.</param>
        /// <returns></returns>
        private static void FormatContent(Range target)
        {

            try
            {
                Stopwatch s = Stopwatch.StartNew();
                if (target.Row == 1)
                {
                    //skip headers
                    target = target.Resize[target.Rows.Count - 1, ColumnsProperties.NbCols].Offset[1];
                }


                Alternate(target);

                //font and colors
                FontStyle(ColumnRange(target, ColumnsProperties.Context), false, true);
                FontStyle(ColumnRange(target, ColumnsProperties.Comment), true);
                FontStyle(ColumnRange(target, ColumnsProperties.Translation), true);
                s.Stop();
                Log?.Info($"Format display took {s.ElapsedMilliseconds} ms for {target.Columns.Count * target.Rows.Count} cells");


            }
            finally
            {
            }
        }

        /// <summary>
        /// Alternates the specified target.
        /// </summary>
        /// <param name="target">The target.</param>
        /// <param name="color">The color.</param>
        private static void Alternate(Range target, XlRgbColor color = XlRgbColor.rgbPaleGreen)
        {
            int cols = ColumnsProperties.NbCols;
            var n = target.Rows.Count;
            var k = target.Row;
            for (int i = k; i < k + n; i++)
            {
                var absoluteRange = target.Worksheet.Rows[i].Resize[1, cols];
                absoluteRange.Interior.Color = (i % 2 == 0) ? color : XlRgbColor.rgbGhostWhite;

            }
        }

        /// <summary>
        /// Freezes the header.
        /// </summary>
        /// <param name="excelSheet">The excel sheet.</param>
        private static void FreezeHeader(Worksheet excelSheet)
        {
            //enable filter and freeze header
            excelSheet.EnableAutoFilter = true;
            excelSheet.Cells.AutoFilter(1);
            excelSheet.Application.ActiveWindow.SplitRow = 1;
            excelSheet.Application.ActiveWindow.FreezePanes = true;
        }


        /// <summary>
        /// Formats the specified target.
        /// </summary>
        /// <param name="target">The target.</param>
        /// <param name="condition">The condition.</param>
        /// <param name="color">The color.</param>
        /// <returns></returns>
        private static FormatCondition Format(Range target, string condition, XlRgbColor color)
        {
            FormatCondition format = target.FormatConditions.Add(XlFormatConditionType.xlExpression, XlFormatConditionOperator.xlEqual, condition);
            //action : change color
            format.Font.Color = color;
            return format;
        }

        /// <summary>
        /// Formats the specified target.
        /// </summary>
        /// <param name="target">The target.</param>
        /// <param name="condition">The condition.</param>
        /// <returns></returns>
        private static FormatCondition Format(Range target, string condition)
        {
            return target.FormatConditions.Add(XlFormatConditionType.xlExpression, XlFormatConditionOperator.xlEqual, condition);
        }


        /// <summary>
        /// Column range
        /// </summary>
        /// <param name="target">The excel sheet.</param>
        /// <param name="columnName">Name of the column.</param>
        /// <param name="withoutHeader">if set to <c>true</c> [without header].</param>
        /// <returns></returns>
        private static Range ColumnRange(Range target, string columnName, bool withoutHeader = true)
        {
            int index = ColumnsProperties.Index(columnName);
            Range r = target.Columns[index + 1].EntireColumn;
            return (withoutHeader) ? r.Resize[target.Rows.Count].Offset[1] : r;
        }

        /// <summary>
        /// Fonts the style.
        /// </summary>
        /// <param name="target">The target range.</param>
        /// <param name="bold">if set to <c>true</c> [bold].</param>
        /// <param name="italic">if set to <c>true</c> [italic].</param>
        /// <param name="color">The color.</param>
        /// <returns></returns>
        private static Range FontStyle(Range target, bool bold = false, bool italic = false, XlRgbColor color = XlRgbColor.rgbBlack)
        {
            Font font = target.Font;
            font.Bold = bold;
            font.Italic = italic;
            font.Color = color;
            return target; //for cascading styles
        }

        /// <summary>
        /// Sets the type of the column.
        /// </summary>
        /// <param name="range">The range of cells (usually a column).</param>
        /// <param name="t">The t.</param>
        private static void SetColumnType(Worksheet excelSheet, Range range, Type t)
        {
            try
            {
                bool isEnum = t.IsEnum;
                string typeName = t.FullName.ToString();
                if (isEnum)
                {
                    string values = string.Empty;
                    foreach (string str in Enum.GetNames(t))
                    {
                        if (!string.IsNullOrEmpty(values))
                        {
                            values += ",";
                        }

                        values += str;
                    }
                    range.Validation.Add(XlDVType.xlValidateList
                                                       , XlDVAlertStyle.xlValidAlertStop
                                                       , XlFormatConditionOperator.xlBetween
                                                       , values
                                                       , Type.Missing);
                    range.Validation.InCellDropdown = true;
                    range.NumberFormat = "@";
                }
                else
                {
                    switch (typeName)
                    {
                        case "System.DateTime":
                            {
                                range.NumberFormat = "MM/DD/YY";
                                break;
                            }
                        case "System.Boolean":
                            {

                                range.Validation.Add(XlDVType.xlValidateList
                                                       , XlDVAlertStyle.xlValidAlertStop
                                                       , XlFormatConditionOperator.xlBetween
                                                       , "True,False"
                                                       , Type.Missing);
                                range.Validation.InCellDropdown = true;
                                break;
                            }

                        case "System.Decimal":
                        //{
                        //    range.NumberFormat = "## ###.##";
                        //    break;
                        //}
                        case "System.Double":
                        //{
                        //    range.NumberFormat = "#.#";
                        //    break;
                        //}
                        case "System.Int":
                        case "System.Int8":
                        case "System.Int16":
                        case "System.Int32":
                        case "System.Int64":
                            {
                                // use general format for int
                                range.NumberFormat = "0";
                                break;
                            }
                        default: //text
                            range.NumberFormat = "@";
                            break;
                    }
                }
            }
            catch
            {

            }
        }


        #endregion
    }
}
