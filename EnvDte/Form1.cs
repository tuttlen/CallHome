using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using EnvDTE;
using EnvDTE80;
using EnvDTE90;
//using EnvDTE100;


namespace EnvDte
{
    public partial class Form1 : Form
    {
        List<string> ItemsWeCallHome;
        public Form1()
        {
            InitializeComponent();
            ItemsWeCallHome = new List<string>();
            var asdf = File.OpenText(@"SupportingFiles\KnownSolutionNames.txt");

            while (!asdf.EndOfStream)
            {
                var lineRead = asdf.ReadLine().Trim();

                ItemsWeCallHome.Add(lineRead);
            }


            //ItemsWeCallHome.Add(@"CompileAllWithUnitTests.sln");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var x = GetInstances();

            var y = x.Last().ActiveDocument.Collection;

            var asp = x.Last().ActiveSolutionProjects;

            var asdf = x.Last().FullName;
            var aSolutionThatWeAreKnowledgeableAbout = GetProperSolution(x);

            var countOfDocs = aSolutionThatWeAreKnowledgeableAbout.Documents.Count;
            for (int i = 1; i < countOfDocs; i++)
            {
                try
                {
                    var name = aSolutionThatWeAreKnowledgeableAbout.Documents.Item(i).Name;
                    Console.WriteLine(@"name " + name);
                }
                catch (Exception)
                {
                    Console.WriteLine(@"For some reason we cant touch this. Is it because we are editing it?");
                }
            }
            
            var theselc = aSolutionThatWeAreKnowledgeableAbout.ActiveDocument.Selection;
            theselc.GotoLine(10, true);           


            //x.Last().ItemOperations.OpenFile(@"C:\src\Command Center\Main\LandisGyr.CommandCenter.AUTD\AutdCollectorProperties.cs");
        }

        private DTE GetProperSolution(IEnumerable<DTE> x)
        {
            string solutionName = string.Empty;
            foreach (var item in x)
            {
                solutionName = item.Solution.FullName;
                var fileName = Path.GetFileName(solutionName);
                if (ItemsWeCallHome.Contains(fileName))
                {
                    return item;
                }                
            }
            return null;
        }

        public IEnumerable<DTE> GetInstances()
        {
            IRunningObjectTable rot;
            IEnumMoniker enumMoniker;
            int retVal = GetRunningObjectTable(0, out rot);

            if (retVal == 0)
            {
                rot.EnumRunning(out enumMoniker);

                IntPtr fetched = IntPtr.Zero;
                IMoniker[] moniker = new IMoniker[1];
                while (enumMoniker.Next(1, moniker, fetched) == 0)
                {
                    IBindCtx bindCtx;
                    CreateBindCtx(0, out bindCtx);
                    string displayName;
                    moniker[0].GetDisplayName(bindCtx, null, out displayName);
                    Console.WriteLine("Display Name: {0}", displayName);
                    bool isVisualStudio = displayName.StartsWith("!VisualStudio");
                    object anObject;
                    if (isVisualStudio)
                    {
                        rot.GetObject(moniker[0], out anObject);
                        yield return anObject as DTE;
                    }
                }
            }
        }

        [DllImport("ole32.dll")]
        private static extern void CreateBindCtx(int reserved, out IBindCtx ppbc);

        [DllImport("ole32.dll")]
        private static extern int GetRunningObjectTable(int reserved, out IRunningObjectTable prot);
    }
}


