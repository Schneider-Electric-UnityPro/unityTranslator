// Copyright (c) 2016 Schneider-Electric
using System;
using System.Xml.Serialization;

namespace SchneiderElectric.UnityComments
{
    /// <summary>
    /// 
    /// </summary>
    [System.Serializable()]
    [XmlType(TypeName = "Comment")]
    public class Comment
    {

        /// <summary>
        /// Gets the key.
        /// </summary>
        /// <value>
        /// The key.
        /// </value>
        [XmlAttribute()]
        public string Key { get;  set; } = $"{{{Guid.NewGuid().ToString()}}}";

        /// <summary>
        /// Gets the value.
        /// </summary>
        /// <value>
        /// The value.
        /// </value>
        [XmlElement("Source", Form = System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string Source { get;  set; }
        /// <summary>
        /// Gets or sets the context.
        /// </summary>
        /// <value>
        /// The context.
        /// </value>
        [XmlAttribute()]
        public string Context { get; set; }

        /// <summary>
        /// Gets or sets the translation.
        /// </summary>
        /// <value>
        /// The translation.
        /// </value>
        [XmlElement("Translation", Form = System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string Translation { get; set; }


        /// <summary>
        /// Initializes a new instance of the <see cref="Comment" /> class.
        /// </summary>
        /// <param name="value">The value.</param>
        public Comment(string value, string context = null)
        {
            Source = value;
            Context = context;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="Comment" /> class.
        /// default for serialisation
        /// </summary>
        public Comment()
        {

        }
    }
}
