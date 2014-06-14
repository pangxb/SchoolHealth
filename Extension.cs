using System;
using System.Collections.Generic;
using System.Text;
using System.Collections;
using System.Reflection;
using System.Diagnostics;
using DevComponents.DotNetBar.Controls;

//////////////////////////////////////////////////////////////////////////
// for using extension methods in Framework 2.0
// If upgrading the Framework, the following code should be commented
//////////////////////////////////////////////////////////////////////////
namespace System.Runtime.CompilerServices
{
    public sealed class ExtensionAttribute : Attribute { }
}

namespace SchoolHealth
{

    public static class Extension
    {
        public static void SelectItem(this ComboBoxEx cmb,  string text)
        {


            int index = cmb.FindStringExact(text);
            if(index>-1)
            {
                cmb.SelectedIndex = index;
                return;
            }
            else 
            {
                if (cmb.Items.Count > 0)
                    cmb.SelectedIndex = 0;
            }
        }

        public static float ToFloat(this Double val)
        {
            return (float)val;
        }

        public static byte ToByte(this Int32 val)
        {
            return (byte)val;
        }

        public static bool IsNullOrEmpty(this string str)
        {
            return string.IsNullOrEmpty(str);
        }

      

        #region 修饰符
        public static BindingFlags nonPublic = BindingFlags.NonPublic | BindingFlags.GetField | BindingFlags.Instance;

        public static void SetValueField(this object obj, string name, object setValue)
        {
            var field = obj.GetType().GetField(name, nonPublic);
            field.SetValue(obj, setValue);
        }

        public static object GetValueField(this object obj, string name)
        {
            var field = obj.GetType().GetField(name, nonPublic);
            var result = field.GetValue(obj);
            return result;
        }

        public static object InvokeMethod(this object obj, string name, object[] parms)
        {
            var method = obj.GetType().GetMethod(name, nonPublic);
            var result = method.Invoke(obj, parms);
            return result;
        }

        #endregion

    }
}
