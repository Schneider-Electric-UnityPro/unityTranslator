// Copyright (c) 2016 Schneider-Electric

using System;
using System.Data;
using System.Linq;
using log4net;

namespace SE.UnityCommentsExcel
{
    /// <summary>
    /// 
    /// </summary>
    static internal class ColumnsProperties
    {
        #region fields

        /// <summary>
        /// The columns names
        /// </summary>
        private static string[] ColumnsName = { "comments", "translation" ,"id", "context"};

        #endregion

        #region accessors
        internal static int NbCols { get { return ColumnsName.Count(); } }
        internal static string Comment { get; } = ColumnsName[0];
        internal static string Translation { get; } = ColumnsName[1];
        internal static string ID { get; } = ColumnsName[2];
        internal static string Context { get; } = ColumnsName[3];
        #endregion

        /// <summary>
        /// Indexes the specified col.
        /// </summary>
        /// <param name="col">The col.</param>
        /// <returns></returns>
        static internal int Index(string col)
        {
            return Array.IndexOf(ColumnsName, col);
        }

        /// <summary>
        ///return col name from index
        /// </summary>
        /// <param name="col">The col name.</param>
        /// <returns></returns>
        static internal string Name(int col)
        {
            return  (col >0 && col < NbCols)? ColumnsName[col-1] : null;
        }
    }


    /// <summary>
    /// data table used as pivot to the scan model
    /// flatten the structure for excel
    /// </summary>
    public class CommentTable : DataTable
    {
        /// <summary>
        /// Gets or sets the logger.
        /// </summary>
        /// <value>
        /// The log.
        /// </value>
        internal static ILog Log { get; set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="ScanTable" /> class.
        /// </summary>
        public CommentTable()
        {
            TableName = "Unity translate";
            //Comment source
            Columns.Add(new DataColumn(ColumnsProperties.Comment, typeof(string)));
            //translation
            Columns.Add(new DataColumn(ColumnsProperties.Translation, typeof(string)));
            //id
            Columns.Add(new DataColumn(ColumnsProperties.ID, typeof(string)));
            //Comment context
            Columns.Add(new DataColumn(ColumnsProperties.Context, typeof(string)));
        }

        /// <summary>
        /// Creates the new row.
        /// </summary>
        /// <returns></returns>
        public CommentRow CreateNewRow()
        {
            return (CommentRow)NewRow();
        }

        /// <summary>
        /// Gets the type of the row.
        /// </summary>
        /// <returns></returns>
        protected override Type GetRowType()
        {
            return typeof(CommentRow);
        }

        /// <summary>
        /// News the row from builder.
        /// </summary>
        /// <param name="builder">The builder.</param>
        /// <returns></returns>
        protected override DataRow NewRowFromBuilder(DataRowBuilder builder)
        {
            return new CommentRow(builder);
        }

    }
}
