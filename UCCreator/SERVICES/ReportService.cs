using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UCCreator.SERVICES
{
    public static class ReportService
    {
        // FIELDS
        private static string _ReportContent = "";
        private static List<string> _ErrorStack = new List<string>();
        private static string _Title = "";
        private static int ErrorCount = 0;

        // VARIABLES
        public static string ReportContent
        {
            get { return _ReportContent; }
            set { if (value != _ReportContent) { _ReportContent = value; } }
        }

        public static List<string> ErrorStack
        {
            get { return _ErrorStack; }
            set { if(value != _ErrorStack) { _ErrorStack = value; } }
        }

        public static string Title
        {
            get { return _Title; }
            set { if (value != _Title) { _Title = value; } }
        }

        // METHODS
        // -------
        /// <summary>
        /// Write new message to the Report
        /// </summary>
        /// <param name="msg"></param>
        public static void Write(string msg)
        {
            ReportContent += msg + Environment.NewLine;
        }

        /// <summary>
        /// Set title for current error stack
        /// </summary>
        /// <param name="title"></param>
        public static void SetTitle(string title)
        {
            Title = title;
        }

        /// <summary>
        /// Get content of Report
        /// </summary>
        /// <returns></returns>
        public static string GetContent()
        {
            // Get final content
            string output = Environment.NewLine + Environment.NewLine +
            "PROCESS REPORT:        " + ErrorCount.ToString() + " ERRORS OCCURRED" + Environment.NewLine +
            "--------------" + Environment.NewLine +
            Environment.NewLine;

            if (ReportContent.Replace(" ","") != "")
            {
                output += ReportContent;
            }
            else
            {
                output += "NO ERRORS.";
            }

            // Reset Report Service
            ReportContent = "";
            ErrorStack.Clear();
            ErrorCount = 0;

            // Return final content
            return output;
        }


        /// <summary>
        /// Add new error message to current error stack
        /// </summary>
        /// <param name="err_msg"></param>
        public static void AddErrorMsg(string err_msg)
        {
            ErrorStack.Add(err_msg);
        }

        /// <summary>
        /// Processes content of current error stack in Report content and resets error stack after
        /// </summary>
        /// <returns></returns>
        public static void ProcessErrorStack()
        {
            // Add ErrorStack content into Report content
            if (ErrorStack.Count > 0)
            {
                Write(Title);

                foreach (string error in ErrorStack)
                {
                    Write("[ERROR]  " + error + Environment.NewLine);
                };

                ErrorCount += ErrorStack.Count;
            }

            // Reset ErrorStack
            ErrorStack.Clear();
        }
    }
}
