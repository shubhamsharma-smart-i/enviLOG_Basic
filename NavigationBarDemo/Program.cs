using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Windows.Forms;
using PDDL.Common;
using System.Security.Principal;
using System.Security.AccessControl;
using System.IO;

namespace PDDL
{
   static class Program
   {
      /// <summary>
      /// The main entry point for the application.
      /// </summary>
      [STAThread]
      static void Main()
      {
         Global.Access.GrantAccess(Application.StartupPath);
         Application.EnableVisualStyles();
         Application.SetCompatibleTextRenderingDefault(false);
         Application.Run(new MainView());      
      }
   }

}
    