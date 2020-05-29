using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SldWorks;
using SwConst;

namespace swAddinChangeExtension
{
    class Program
    {
        static SldWorks.SldWorks swApp;
        static ModelDoc2 swModel;
        static string directory = @"C:\rename";//変換するファイルのみを格納したフォルダ
        static string extension = ".step";//変換後の拡張子
        
        /// <summary>
        /// 指定のフォルダ内にある３D モデルファイルを一括で別形式に変換するプログラム
        /// </summary>
        /// <param name="args"></param>
        static void Main(string[] args)
        {

            string filename = directory;

            string[] files = System.IO.Directory.GetFiles(filename);

            foreach(string file in files)
            {
                swApp = new SldWorks.SldWorks();
                swApp.Visible = false;
                Console.WriteLine(file);

                bool bRet = false;
                string strArg = null;
                int Err = 0;

                object importData = swApp.GetImportFileData(file);
                swModel = swApp.LoadFile4(file, strArg, importData, ref Err);


                string newfilename = file.Remove(file.IndexOf(".")) + extension;

                swModel.ClearSelection2(true);
                bRet = swApp.SetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swStepAP, 214);

                ModelDocExtension swExt = swModel.Extension;
                bRet = swExt.SaveAs(newfilename, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, 0, null, Err, 0);
                swApp.ExitApp();
            }


        }

    }
}
