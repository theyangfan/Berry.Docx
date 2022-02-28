using System;
using System.Collections.Generic;
using System.Text;

namespace Berry.Docx
{
    /// <summary>
    /// Represent a custom attribute for the enumeration field.
    /// </summary>
    [AttributeUsage(AttributeTargets.Field)]
    public sealed class DescriptionAttribute : Attribute
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="DescriptionAttribute"/> class with a specified
        ///  workaround description.
        /// </summary>
        /// <param name="description">The description.</param>
        public DescriptionAttribute(string description)
        {
            Val = description;
        }

        /// <summary>
        /// Gets the description of the current attribute.
        /// </summary>
        public string Val
        {
            get;
            private set;
        }
    }

    public static class EnumExtentions
    {
        public static string Description<T>(this T t)
        {
            string fieldName = Enum.GetName(typeof(T), t);
            object[] attrs = typeof(T).GetField(fieldName).GetCustomAttributes(typeof(DescriptionAttribute), false);
            if(attrs.Length > 0)
            {
                DescriptionAttribute attr = attrs[0] as DescriptionAttribute;
                return attr.Val;
            }
            return fieldName;
        }
    }
}
