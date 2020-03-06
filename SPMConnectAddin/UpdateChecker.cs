using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Threading;
using System.Windows.Forms;
using System.Xml;

namespace SPMConnectAddin
{
    internal class UpdateChecker
    {
        public static void startAsync(string metaFileURL, string name, string root, bool silent)
        {
            _UpdateChecker uc = new _UpdateChecker();
            uc.rootname = root;
            uc.URL = metaFileURL;
            uc.silent = silent;
            uc.name = name;
            Thread t = new Thread(uc.start);
            t.Start();
        }

        public static void start(string metaFileURL, string name, string root, bool silent)
        {
            _UpdateChecker uc = new _UpdateChecker();
            uc.rootname = root;
            uc.URL = metaFileURL;
            uc.silent = silent;
            uc.name = name;
            uc.start();
        }
    }

    internal class _UpdateChecker
    {
        public string rootname { get; set; }
        public string name { get; set; }
        public string URL { get; set; }
        public bool silent { get; set; }

        public void start()
        {
            XmlDocument x = new XmlDocument();

            try
            {
                x.Load(this.URL);
                XmlNode root = x.DocumentElement;
                if (root.Name != rootname)
                {
                    throw new XmlException();
                }
                string version = x.SelectSingleNode("descendant::version").InnerText;
                if (version == null)
                {
                    throw new XmlException();
                }
                if (String.Compare(version, Assembly.GetExecutingAssembly().GetName().Version.ToString()) == 1)
                {
                    DialogResult r = MessageBox.Show("New version of " + name + " is available to install. Would you like to install it?", "SPM Connect - Update available", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                    if (r == DialogResult.Yes)
                    {
                        string applicationfolder = @"\\spm-adfs\SDBASE\SPM Connect Addin\" + version + "\\" + "SPMConnectAddin" + version + ".msi";
                        File.Copy(applicationfolder, System.IO.Path.GetTempPath() + "SPMConnectAddin" + version + ".msi");
                        Process.Start(System.IO.Path.GetTempPath() + "SPMConnectAddin" + version + ".msi");
                    }
                }
                else if (!silent)
                {
                    MessageBox.Show("No new version version available for this Addin. Please check back later.", "SPM Connect - No update available", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception e)
            {
                if (!silent)
                {
                    MessageBox.Show(e.Message, "SPM Connect Update Checker", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
    }
}