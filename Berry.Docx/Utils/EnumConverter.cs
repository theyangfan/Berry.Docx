using System;
using System.Collections.Generic;
using System.Text;

namespace Berry.Docx
{
    internal static class EnumConverter
    {
        /// <summary>
        /// 
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="val"></param>
        /// <returns></returns>
        /// <exception cref="InvalidCastException"></exception>
        public static T Convert<T>(this Enum val)
        {
            Type type = val.GetType();
            string fieldname = Enum.GetName(type, val);
            object value = null;
            try
            {
                value = Enum.Parse(typeof(T), fieldname, true);
            }
            catch (Exception) 
            {
                throw new InvalidCastException($"{typeof(T).FullName} does not have a field named \"{fieldname}\" !");
            }
            return value != null ? (T)value : default(T);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="val"></param>
        /// <param name="defaultValue"></param>
        /// <returns></returns>
        public static T Convert<T>(this Enum val, T defaultValue)
        {
            Type type = val.GetType();
            string fieldname = Enum.GetName(type, val);
            object value = null;
            try
            {
                value = Enum.Parse(typeof(T), fieldname, true);
            }
            catch (Exception)
            {}
            return value != null ? (T)value : defaultValue;
        }
    }
}
