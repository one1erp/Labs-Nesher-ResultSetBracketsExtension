using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using Common;
using LSEXT;


namespace ResultSetBracketsExtension
{
    [ComVisible(true)]
    [ProgId("ResultSetBracketsExtension.ResultSetBracketsExtensionCLS")]
    public class ResultExtension : IResultFormat
    {


        public ResultFieldChange FieldChange(ref LSExtensionParameters Parameters)
        {
            return LSEXT.ResultFieldChange.rcAllow;

        }

        public ResultEntryFormat Format(ref LSExtensionParameters Parameters, ResultEntryPhase Phase)
        {
            try
            {


                if (Phase == ResultEntryPhase.reValidate)
                {
                    try
                    {
                        var formatted_result = Parameters.Parameter("formatted_result").Value;

                        if (formatted_result == null)
                        {
                            return LSEXT.ResultEntryFormat.rfDoDefault;
                        }
                        else
                        {
                            var newValue = SetBrackest(formatted_result);
                            if (newValue == null)
                            {
                                return LSEXT.ResultEntryFormat.rfDoDefault;

                            }
                            Parameters.Parameter("formatted_result").Value = newValue;

                        }

                        return LSEXT.ResultEntryFormat.rfSkipDefault;
                    }
                    catch (COMException exception)
                    {
                        //    Logger.WriteLogFile(exception);

                    }
                }
                return LSEXT.ResultEntryFormat.rfDoDefault;
            }
            catch (Exception e)
            {

                Logger.WriteLogFile(e);
                return LSEXT.ResultEntryFormat.rfDoDefault;

            }
        }

        private string SetBrackest(string formattedResult)
        {
            var index = IsNumeric(formattedResult);
            if (index == 0)
            {
                return null;
            }
            string a = formattedResult.Substring(0, index);
            string b = formattedResult.Substring(index);
            string o = a + " (" + b + ")";
            return o;

        }

        public int IsNumeric(string s)
        {
            foreach (char c in s)
            {
                if (!char.IsDigit(c) && c != '.')
                {
                    return s.IndexOf(c);
                }
            }

            return 0;
        }

    }
}
