using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.Office.Interop.OneNote;
using System.Runtime.InteropServices;
using System.Diagnostics;

/*Due to the imperfect nature of Tenjin Sync (as well as the limited methods of syncing OneNote), it's recommended to run it as its own user.*/
namespace Tenjin_Sync
{
    public partial class Form1 : Form
    {
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool SetForegroundWindow(IntPtr hWnd);
        [DllImport("user32.dll", EntryPoint = "FindWindow")]
        private static extern int FindWindow(string sClass, bool sWindow);
        
        //The path to publish the XML feeds to (Usually somewhere in the IIS or Apache folder, ensure the user has write permissions)
        String iisPath = "C:\\inetpub\\wwwroot\\onenote\\";

        //Full names, as registered in their respective Microsoft accounts
        String roommate1 = "";
        String roommate2 = "";

        public Form1()
        {
            InitializeComponent();
            refreshOnenote();
        }

        public void readHomework()
        {
            var onenoteApp = new Microsoft.Office.Interop.OneNote.Application();

            string notebookXml;
            onenoteApp.GetHierarchy(null, HierarchyScope.hsPages, out notebookXml);

            var doc = XDocument.Parse(notebookXml);
            var ns = doc.Root.Name.Namespace;
            foreach (var notebookNode in from node in doc.Descendants(ns + "Notebook") select node)
            {
                if (notebookNode.Attribute("name").Value == "Tenjin")
                {
                    foreach (var sectionNode in from node in notebookNode.Descendants(ns + "Section") select node)
                    {
                        foreach (var pageNode in from node in sectionNode.Descendants(ns + "Page") select node)
                        {
                            if (pageNode.Attribute("name").Value == roommate1.Split(' ')[0])
                            {
                                string hwxml;
                                onenoteApp.GetPageContent(pageNode.Attribute("ID").Value, out hwxml);
                                System.IO.File.WriteAllText(iisPath + roommate1.Split(' ')[0] + ".xml", XDocument.Parse(hwxml).ToString().Replace(roommate1, "").Replace(roommate1.Split(' ')[0], ""));
                            }
                            if (pageNode.Attribute("name").Value == roommate2.Split(' ')[0])
                            {
                                string hwxml;
                                onenoteApp.GetPageContent(pageNode.Attribute("ID").Value, out hwxml);
                                System.IO.File.WriteAllText(iisPath + roommate2.Split(' ')[0] + ".xml", XDocument.Parse(hwxml).ToString().Replace(roommate2, "").Replace(roommate2.Split(' ')[0], ""));
                            }
                        }
                    }
                }
            }

            refreshOnenote();
        }

        public void refreshOnenote()
        {
            Process.Start("C:\\Program Files\\Microsoft Office\\Office15\\ONENOTE.EXE");
            BringWindowToTop("Homework - Tenjin", false);

            System.Windows.Forms.Timer syncTimer = new System.Windows.Forms.Timer();
            syncTimer.Tick += new EventHandler(killOnenote);
            syncTimer.Interval = 1000 * 60 * 2;
            syncTimer.Enabled = true;
            syncTimer.Start();
        }

        public void killOnenote(Object source, EventArgs e)
        {
            foreach (Process proc in Process.GetProcessesByName("ONENOTE"))
            {
                proc.Kill();
            }

            readHomework();
        }

        public static bool BringWindowToTop(string windowName, bool wait)
        {
            int hWnd = FindWindow(windowName, wait);
            if (hWnd != 0)
            {
                return SetForegroundWindow((IntPtr)hWnd);
            }

            return false;
        }
    }
}