using System;
using System.IO;
using IWshRuntimeLibrary;
using System.Windows.Forms;

namespace Fold
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            //Install the Send To shortcut
            Install();
            InitializeComponent();
        }
        private void Install()
        {
            //Get the Send To shortcut folder
            string folder = Environment.GetEnvironmentVariable("appdata") + "\\Microsoft\\Windows\\SendTo\\";
            //Create a WshShell
            WshShell shell = new WshShell();
            //Create the shortcut
            IWshShortcut sc = (IWshShortcut) shell.CreateShortcut(folder + "Fold.lnk");
            sc.Description = "Fold SendTo Handler";
            sc.IconLocation = Environment.GetCommandLineArgs()[0].Replace(".exe", ".ico") ;
            sc.TargetPath = Environment.GetCommandLineArgs()[0];
            //Save it
            sc.Save();
        }
        private void GoClick(object sender, EventArgs e)
        {
            //If the textbox isn't empty
            if (!string.IsNullOrEmpty(textBox1.Text))
            {
                //Get the file names from Send To
                string[] args = Environment.GetCommandLineArgs();
                //Get the path of those files
                string path = args[1].Substring(0, args[1].LastIndexOf('\\'));
                //Get the new folder's name
                string FolderName = textBox1.Text;
                //Get the new folder's path
                string newFolderPath = path + '\\' + FolderName;
                //Create it if it doesn't exist.
                Directory.CreateDirectory(newFolderPath);
                //For each file name
                foreach (string s in args)
                {
                    //Ignore it if it's the executable's file name
                    if (s == args[0]) continue;
                    //Move the selected file into the new folder.
                    System.IO.File.Move(s, s.Substring(0, s.LastIndexOf('\\')) + "\\" + FolderName + s.Substring(s.LastIndexOf('\\')));
                }
                //Exit.
                Application.Exit();
            }
        }
    }
}
