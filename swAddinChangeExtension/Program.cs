using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using SldWorks;
using SwConst;
using System.Drawing;
using System.Windows.Forms;

namespace swAddinChangeExtension
{
    class Program
    {
        static SldWorks.SldWorks swApp;
        static ModelDoc2 swModel;
        static string directory = System.Environment.GetFolderPath(System.Environment.SpecialFolder.DesktopDirectory) + @"\source";//変換するファイルのみを格納したフォルダ
        static string extension = ".step";//変換後の拡張子
        static bool thumbnail = false;
        static bool exportStep = false;
        /// <summary>
        /// 指定のフォルダ内にある３D モデルファイルを一括で別形式に変換するプログラム
        /// </summary>
        /// <param name="args"></param>
        static void Main(string[] args)
        {
            if (args.Count() > 0)
            {
                foreach (string arg in args)
                {
                    string argVal = arg.Substring(arg.IndexOf("=") + 1);
                    switch (arg.Substring(0, arg.IndexOf("=")))
                    {
                        case "--tn":
                            if (argVal == "true")
                            {
                                thumbnail = true;
                            }
                            break;
                        case "--step":
                            if (argVal == "true")
                            {
                                exportStep = true;
                            }
                            break;
                        case "--target":
                            directory = argVal;
                            break;
                        default:
                            break;
                    }
                }
            }

            string[] files = System.IO.Directory.GetFiles(directory, "*", searchOption: SearchOption.AllDirectories);

            foreach (string file in files)
            {
                string filename = file.Remove(0, directory.Length + 1).Replace(@"\", "_");
                string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(file);
                string extension = Path.GetExtension(file);
                if (extension != ".step " && extension != ".Step" && extension != ".STEP" && extension != ".x_t" && extension != ".X_T" && extension!=".stp")
                {
                    continue;
                }
                swApp = new SldWorks.SldWorks();
                swApp.Visible = thumbnail;
                Console.WriteLine(file);

                bool bRet = false;
                string strArg = null;
                int Err = 0;

                object importData = swApp.GetImportFileData(file);
                swModel = swApp.LoadFile4(file, strArg, importData, ref Err);

                if (thumbnail)
                {
                    swModel.ShowNamedView2("", 7);
                    swModel.ViewZoomtofit2();
                    swApp.FrameState = 1;

                    Thread.Sleep(1000);
                    Microsoft.VisualBasic.Interaction.AppActivate(fileNameWithoutExtension);

                    Bitmap bmp = new Bitmap(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height);
                    Graphics g = Graphics.FromImage(bmp);
                    g.CopyFromScreen(new Point(0, 0), new Point(0, 0), bmp.Size);

                    string bmpDirectory = directory + @"\bmp\";
                    if (!Directory.Exists(bmpDirectory))
                    {
                        Directory.CreateDirectory(bmpDirectory);
                    }
                    filename = bmpDirectory + filename + ".bmp";
                    bmp.Save(filename, System.Drawing.Imaging.ImageFormat.Bmp);
                    g.Dispose();
                }

                if (exportStep)
                {
                    string newfilename = file.Remove(file.IndexOf(".")) + extension;

                    swModel.ClearSelection2(true);
                    bRet = swApp.SetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swStepAP, 214);

                    ModelDocExtension swExt = swModel.Extension;
                    bRet = swExt.SaveAs(newfilename, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, 0, null, Err, 0);
                }

                swApp.ExitApp();
            }


        }

    }
}
