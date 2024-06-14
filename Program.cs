using OAS.Util.Configuration;
using OAS.Util.Logging;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Printing;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ZebraSpooler {
    internal class Program {

        private static IniFile config;

        static int Main(string[] args) {
            int ret = -1;
            try {
                config = new IniFile("Zebra.ini");

                Log.LogFileName = config.ReadString("LogDirectory", null, "Log") + "\\Zebra.log";
                Log.Init(true, 200);

                if (args.Length == 0) {
                    ret = 68100;
                    return 68100;
                }

                bool delete = config.ReadBool("DeleteFile");

                if (Directory.Exists(args[0])) {
                    foreach (string file in Directory.EnumerateFiles(args[0])) {
                        try {
                            ret = Print(file);
                            if (ret != 0) {
                                return ret;
                            }
                        } finally {
                            if (delete) {
                                File.Delete(file);
                            }
                        }
                    }
                } else {
                    try { 
                        ret = Print(args[0]);
                    } finally {
                        if (delete) {
                            File.Delete(args[0]);
                        }
                    }
                }

            }catch(Exception ex) {
                Log.WriteFault(ex, "Printing Error");
                ret = 68040;
            } finally {
                Log.Write("Result: " + ret);
                Log.Close();
            }
            return ret;
        }

        private static int Print(String file) {

            Image image = Image.FromFile(file);
            int r = config.ReadInt("Rotate");
            bool fx = config.ReadBool("FlipX");
            bool fy = config.ReadBool("FlipY");

            if (fx) {
                image.RotateFlip(RotateFlipType.RotateNoneFlipX);
            }
            if (fy) {
                image.RotateFlip(RotateFlipType.RotateNoneFlipX);
            }
            if (r == 90) {
                image.RotateFlip(RotateFlipType.Rotate90FlipNone);
            } else if (r == 180) {
                image.RotateFlip(RotateFlipType.Rotate180FlipNone);
            } else if (r == 270) {
                image.RotateFlip(RotateFlipType.Rotate270FlipNone);
            }
            int w;
            int h;
            int bw = config.ReadInt("PrintWidth") / 3;
            int bh = config.ReadInt("PrintHeight") / 3;
            int xoff = config.ReadInt("XOffset");
            int yoff = config.ReadInt("YOffset");
            int stretch = config.ReadInt("StretchMode", null, 3);
            switch (stretch) {
                case 1:
                    w = image.Width;
                    h = image.Height;
                    break;
                case 2:
                    double dbl = (double)image.Width / (double)image.Height;
                    if ((int)((double)bh * dbl) <= bw) {
                        w = (int)((double)bh * dbl);
                        h = bh;
                    } else {
                        w = bw;
                        h = (int)((double)bw / dbl);
                    }
                    break;
                case 3:
                    w = bw;
                    h = bh;
                    break;
                default:
                    throw new IOException("unknown mode: " + stretch);
            }

            PrintDocument pd = new PrintDocument();
            pd.DefaultPageSettings.Landscape = config.ReadBool("Landscape");

            pd.PrintPage += (o, e2) => {
                e2.Graphics.DrawImage(image, new Rectangle(xoff, yoff, w, h));
            };

            if (!config.ReadBool("PreviewOnly")) {

                PrintController printController = new StandardPrintController();
                pd.PrintController = printController;
                pd.Print();

                var myPrintServer = new LocalPrintServer();
                var pq = myPrintServer.DefaultPrintQueue;

                PrintJobInfoCollection jobs;
                do {
                    jobs = pq.GetPrintJobInfoCollection();
                    pq.Refresh();
                    Thread.Sleep(500);
                } while (jobs.Count() == 0);
                var job = jobs.First();
                var done = false;
                while (!done) {
                    pq.Refresh();
                    job.Refresh();
                    done = job.IsCompleted || job.IsDeleted || job.IsPrinted;
                    if (job.JobStatus == PrintJobStatus.PaperOut) {
                        return 68010;
                    } else if (job.JobStatus == PrintJobStatus.UserIntervention) {
                        return 68150;
                    }
                }

                Log.Write("Done");
                return 0;
            } else {

                PrintPreviewDialog printpreview = new PrintPreviewDialog();
                printpreview.Document = pd;
                printpreview.ShowDialog();
                return 0;
            }
        }
    }
}
