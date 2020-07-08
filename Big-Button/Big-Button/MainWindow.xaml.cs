using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Xceed.Words.NET;

namespace Big_Button
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void BigButton_Click(object sender, RoutedEventArgs e)
        {
            var document = DocX.Create("MyReport.docx");
            document.InsertParagraph($"REPORT\n");
            document.Save();

            LabelLiveOutput.Content = "Created document";

            createFoldersAndFiles();
        }

        public void createFoldersAndFiles()
        {
            for (int k=0; k<3; k++)
            {
                #region loop with stopwatch
                var stopwatch = new Stopwatch();
                stopwatch.Start();

                string folderName = @"c:\Big-Button-Folder";
                string pathString = System.IO.Path.Combine(folderName, "AllFiles");

                if (System.IO.Directory.Exists(pathString))
                {
                    string[] files = Directory.GetFiles(pathString);
                    foreach (string file in files)
                    {
                        File.SetAttributes(file, FileAttributes.Normal);
                        File.Delete(file);
                    }

                    System.IO.Directory.Delete(pathString);
                };

                System.IO.Directory.CreateDirectory(pathString);

                for (int i = 0; i <= 100; i++)
                {
                    string fileName = "file" + i + ".txt";
                    string filePath = System.IO.Path.Combine(pathString, fileName);
                    
                    if(!System.IO.File.Exists(filePath))
                    {
                        for(int j=0; j<(k+1)*10; j++)
                        {
                            File.WriteAllText(filePath, $"Line {j+1}");
                        }
                    }
                    else
                    {
                        Console.WriteLine($"{fileName} already exists");
                    }
                }

                stopwatch.Stop();
                var timeElapsed = stopwatch.Elapsed;
                var timeElapsedMilli = stopwatch.ElapsedMilliseconds;
                var timeTicks = stopwatch.ElapsedTicks;

                manipulateReport("Iteration " + k, timeElapsed, timeElapsedMilli, timeTicks, (k+1)*10);

                #endregion
            }

            LabelLiveOutput.Content = "Finished timing, opening file";
            Process.Start("WINWORD.EXE", "MyReport.docx");
        }

        public static void manipulateReport(string fileName, TimeSpan timeElapsed, double timeElapsedMilli, double timeTicks, int amountOfLines)
        {
            var document1 = DocX.Load("MyReport.docx");
            document1.InsertParagraph($"===============================");
            document1.InsertParagraph($"{fileName} stats, with {amountOfLines} lines");
            document1.InsertParagraph($"===============================");
            document1.InsertParagraph($"Time Elapsed: " + timeElapsed);
            document1.InsertParagraph($"Time Elapsed Milli: " + timeElapsedMilli);
            document1.InsertParagraph($"Time Ticks: " + timeTicks);
            document1.InsertParagraph($"===============================\n");
            document1.Save();
        }
    }
}
