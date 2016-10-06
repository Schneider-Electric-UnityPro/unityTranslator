using log4net;
using System;
using System.Collections.Generic;
using System.Data;

namespace SE.UnityCommentsExcel
{
    /// <summary>
    /// 
    /// </summary>
    public class CommentRow : DataRow
    {
        #region properties

        /// <summary>
        /// Gets or sets the logger.
        /// </summary>
        /// <value>
        /// The log.
        /// </value>
        internal static ILog Log { get; set; }


        /// <summary>
        /// Gets or sets the comment.
        /// </summary>
        /// <value>
        /// The comment.
        /// </value>
        public string Comment
        {
            get { return Get<string>(ColumnsProperties.Comment); }
            set { base[ColumnsProperties.Comment] = value; }
        }

        /// <summary>
        /// Gets or sets the translation.
        /// </summary>
        /// <value>
        /// The translation.
        /// </value>
        public string Translation
        {
            get { return Get<string>(ColumnsProperties.Translation); }
            set { base[ColumnsProperties.Translation] = value; }
        }

        /// <summary>
        /// Gets or sets the identifier.
        /// </summary>
        /// <value>
        /// The identifier.
        /// </value>
        public string ID
        {
            get { return Get<string>(ColumnsProperties.ID); }
            set { base[ColumnsProperties.ID] = value; }
        }

        /// <summary>
        /// Gets or sets the context.
        /// </summary>
        /// <value>
        /// The context.
        /// </value>
        public string Context
        {
            get { return Get<string>(ColumnsProperties.Context); }
            set { base[ColumnsProperties.Context] = value; }
        }

        //
        // Summary:
        //     Gets or sets all the values for this row through an array.
        //
        // Returns:
        //     An array of type System.Object.
        //
        // Exceptions:
        //   T:System.ArgumentException:
        //     The array is larger than the number of columns in the table.
        //
        //   T:System.InvalidCastException:
        //     A value in the array does not match its System.Data.DataColumn.DataType in its
        //     respective System.Data.DataColumn.
        //
        //   T:System.Data.ConstraintException:
        //     An edit broke a constraint.
        //
        //   T:System.Data.ReadOnlyException:
        //     An edit tried to change the value of a read-only column.
        //
        //   T:System.Data.NoNullAllowedException:
        //     An edit tried to put a null value in a column where System.Data.DataColumn.AllowDBNull
        //     of the System.Data.DataColumn object is false.
        //
        //   T:System.Data.DeletedRowInaccessibleException:
        //     The row has been deleted.
        public new object[] ItemArray
        {
            get
            {
                var list = new List<object>();
                list.Add(this.Comment);
                list.Add(this.Translation);
                list.Add(this.ID);
                list.Add(this.Context);

                return list.ToArray();
            }
            set
            {
                Comment = StringFromValue(value, ColumnsProperties.Comment);
                Translation = StringFromValue(value, ColumnsProperties.Translation);
                ID = StringFromValue(value, ColumnsProperties.ID);
                Context = StringFromValue(value, ColumnsProperties.Context);
            }

        }

        


        #endregion


        /// <summary>
        /// Initializes a new instance of the <see cref="ScanRow" /> class.
        /// </summary>
        /// <param name="builder">The builder.</param>
        internal CommentRow(DataRowBuilder builder)
                    : base(builder)
        {
            
        }


        #region private
        /// <summary>
        /// Gets the  value for the specified column
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="colName">Name of the col.</param>
        /// <param name="nullValue">The null value.</param>
        /// <returns></returns>
        private T Get<T>(string colName, T nullValue = default(T))
        {
            if (base[colName] != DBNull.Value)
            {
                try
                {
                    return (T)base[colName];
                }
                catch
                {
                    return default(T);
                }
            }
            return nullValue;
        }


        /// <summary>
        /// Strings from value.
        /// </summary>
        /// <param name="value">The value.</param>
        /// <param name="colname">The colname.</param>
        /// <param name="defValue">The definition value.</param>
        /// <returns></returns>
        private static string StringFromValue(object[] value, string colname, string defValue = "")
        {
            string res = defValue;
            try
            {
                int index = ColumnsProperties.Index(colname);
                if (index != -1)
                {
                    res = value[index]?.ToString();
                }
            }
            catch
            {
                res = defValue;
            }
            return res;
        }


        #endregion
    }
}
