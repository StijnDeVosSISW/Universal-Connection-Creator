using NXOpen;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UCCreator
{
    class Program
    {
        static void Main(string[] args)
        {
            NXOpen.Session theSession = NXOpen.Session.GetSession();
            NXOpen.ListingWindow lw = theSession.ListingWindow;
            lw.Open();

            lw.WriteFullline(
                " ------------------------------ " + Environment.NewLine +
                " ------------------------------ " + Environment.NewLine +
                "| UNIVERSAL CONNECTION CREATOR |" + Environment.NewLine +
                " ------------------------------ " + Environment.NewLine +
                " ------------------------------ " + Environment.NewLine);

            try
            {
                
            }
            catch (Exception e)
            {
                lw.WriteFullline("!ERROR occurred: " + Environment.NewLine +
                    e.ToString());
            }
        }

        /// <summary>
        /// Define Unload Option of NXOpen application
        /// </summary>
        /// <param name="arg"></param>
        /// <returns></returns>
        public static int GetUnloadOption(string arg)
        {
            //return System.Convert.ToInt32(Session.LibraryUnloadOption.Explicitly);
            return System.Convert.ToInt32(Session.LibraryUnloadOption.Immediately);
            // return System.Convert.ToInt32(Session.LibraryUnloadOption.AtTermination);
        }
    }
}
