using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows;
using System.IO.Compression;
using System.IO;
using System.Drawing.Drawing2D;
using CZiplib = ICSharpCode.SharpZipLib;
using CZip = ICSharpCode.SharpZipLib.Zip;
using CZipSum =ICSharpCode.SharpZipLib.Checksum;
using System.Threading;

namespace Excel_Comments_Fixer
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }


        private void Button1_Click(object sender, EventArgs e)
        {
            int i = 0;
            OpenFileDialog path = new OpenFileDialog();
            //文件类型过滤
            path.Filter = @"待修复Excel文档|*.xlsx|待修复带有宏的Excel文档|*.xlsm";
            //展开文件选择对话框
            DialogResult result = path.ShowDialog();
            if (result == DialogResult.OK)
            {
                //获取导出文件夹目录
                string expath = path.FileName + @"_Fix";
                if (Directory.Exists(expath))
                {
                    DialogResult fg = MessageBox.Show("导出数据已存在！是否覆盖？覆盖后将失去所有已导出数据！", "警告", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                    if (fg == DialogResult.Yes)//确认删除
                    {
                        //对重复文件夹进行删除操作
                        DirectoryInfo dl = new DirectoryInfo(expath);
                        dl.Delete(true);
                        DialogResult qr = MessageBox.Show("删除命令已执行，是否继续？", "提示", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (qr == DialogResult.Yes)
                        {
                            //继续执行命令
                        }
                        else
                        {
                            if (qr == DialogResult.No)//删除文件夹后不继续
                            {
                                i = 1;
                            }
                        }
                    }
                    else
                    {
                        if (fg == DialogResult.No)//确认不删除
                        {
                            i = 1;
                        }
                    }
                }
                //强制用户确认后才进行解压操作
                if (i == 0)
                {
                    //获取打开文件的路径（全路径，含文件名）
                    int fnlength = path.ToString().LastIndexOf("FileName");
                    string sFile = path.ToString().Substring(fnlength + 9, path.ToString().Length - fnlength - 9);
                    //获取打开文件所在目录，传递至sourcepath
                    string sourcepath = Path.GetDirectoryName(sFile);

                    //开始解压文件
                    ZipFile.ExtractToDirectory(path.FileName, expath);
                    //跳转至xl文件夹下
                    string copath = expath + "\\xl\\";
                    //获取xl文件夹下所有带有comments字符的xml文件路径
                    string[] pathFile = Directory.GetFiles(copath, "comments*.xml", SearchOption.TopDirectoryOnly);
                    //将中间文件命名为cobak.xml
                    string strcon = copath + @"cobak.xml";
                    //暂时将替换容器内容清空
                    string con = "";
                    foreach (string str in pathFile)
                    {
                        StreamReader reader = new StreamReader(str, Encoding.UTF8);
                        //读取至文件尾
                        con = reader.ReadToEnd();
                        //替换shapeId="0"为空值
                        con = con.Replace(@"shapeId=""0""", "");
                        //替换Excel2013的不正确的XML代码块
                        con = con.Replace(@"</commment></commentList></comments>", "");

                        //开始以UTF-8编码开始写操作
                        StreamWriter writer = new StreamWriter(strcon, false, Encoding.UTF8);
                        //写入已替换的内容
                        writer.Write(con);
                        //释放写入缓冲区，暂时关闭写入器
                        writer.Flush();
                        writer.Close();
                        reader.Close();

                        //将写入文件复制为原文件，删除写入文件
                        File.Copy(strcon, str, true);
                        File.Delete(strcon);
                    }
                    //到此替换写入操作已完成
                    //打开Excel文件和注释所在路径
                    //System.Diagnostics.Process.Start("explorer.exe", sourcepath);
                    //System.Diagnostics.Process.Start("explorer.exe", copath);

                    //开始修改文件扩展名
                    //Path.GetExtension(sFile);
                    //Path.GetFileName(sFile);
                    string nFile = Path.ChangeExtension(sFile, ".zip");
                    FileInfo fi = new FileInfo(sFile);
                    fi.MoveTo(nFile);



                    //MessageBox.Show(nFile, "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    CZip.ZipFile zip = new CZip.ZipFile(nFile);
                    string list = string.Empty;
                    foreach (CZip.ZipEntry entry in zip)
                    {
                        list += entry.Name + "\r\n";
                    }
                    //MessageBox.Show(list);
                    zip.BeginUpdate();
                    foreach (string str in pathFile)
                    {
                        string filename = Path.GetFileName(str);
                        zip.Add(str, @"/xl/" + filename);
                        //MessageBox.Show(@"/xl/" + filename, "T", MessageBoxButtons.OK, MessageBoxIcon.None);
                    }
                    zip.CommitUpdate();
                    zip.Close();
                    string endFile = Path.ChangeExtension(nFile, ".xlsx");
                    fi.CopyTo(endFile, true);
                    DirectoryInfo dl = new DirectoryInfo(expath);
                    dl.Delete(true);
                    File.Delete(nFile);

                    /*DirectoryInfo dl = new DirectoryInfo(expath);
                    dl.Delete(true);*/


                    System.Diagnostics.Process.Start("explorer.exe", sourcepath);
                }
            }
        }

        /*废弃代码存放处
        //MessageBox.Show(path.ToString().Length.ToString(), "info", MessageBoxButtons.OK, MessageBoxIcon.None);
        // string sourcepath = path.FileName.ToString();
        //MessageBox.Show(sourcepath, "source_path", MessageBoxButtons.OK, MessageBoxIcon.None);
        //string afterpath = path + @"Fixer.rar";
        //MessageBox.Show("文件夹已存在！是否覆盖？覆盖后将失去已解压的所有数据！", "警告", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
        //MessageBox.Show("删除命令已执行，是否继续？", "提示", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
        //MessageBox.Show(copath, "copath", MessageBoxButtons.OK, MessageBoxIcon.None);
        //
        //int[] a = new int[100];
        //string strFilePath = copath;
        //string strIndex= @"shapeId=""0""";
        //MessageBox.Show(strIndex, "info", MessageBoxButtons.OK, MessageBoxIcon.None);
        //string newValue="";
        //{
            if (System.IO.File.Exists(strFilePath))
            {
                string[] lines = System.IO.File.ReadAllLines(strFilePath);
                for (int i = 0; i < lines.Length; i++)
                {
                    if (lines[i].Contains(strIndex))
                    {
                        string[] str = lines[i].Split('=');
                        str[1] = newValue;
                        lines[i] = str[0] + " = " + str[1];
                    }
                }
                System.IO.File.WriteAllLines(strFilePath, lines);
            }
        }
        string oldString;
        string newString = "";
        FileStream fs = new FileStream(allpath,FileMode.Open,FileAccess.ReadWrite);
        StreamReader sr = new StreamReader(fs, Encoding.Default);
        oldString = sr.ReadToEnd();
        newString = oldString.Replace(@"shapeId=""0""", newString);
        sr.Close();
        StreamWriter sw = new StreamWriter(fs, Encoding.Default);
        sw.Flush(); //清除流内容
        sw.Write(newString); //重写替换内容后的字符串
        sw.Close();
        fs.Close();
        // int i = 1;
        // string a = i.ToString();
        // string endpath = "*.xml";
        // string allpath = copath + endpath;
        //
        //string aapath = allpath;
                        FileStream fs = new FileStream(str, FileMode.Open, FileAccess.ReadWrite);
                        MessageBox.Show(str, "str", MessageBoxButtons.OK, MessageBoxIcon.None);
                        StreamReader sr = new StreamReader(fs);
                        con = sr.ReadToEnd();
                        MessageBox.Show(con, "con_read_file_until_end", MessageBoxButtons.OK, MessageBoxIcon.None);
                        con = con.Replace(@"shapeId=""0""", "");
                        con = con.Replace(@"</commment></commentList></comments>", "");
                        MessageBox.Show(con, "con_after_trans", MessageBoxButtons.OK, MessageBoxIcon.None);
                        sr.Close();
                        fs.Close();
                        FileStream fs2 = new FileStream(str, FileMode.OpenOrCreate, FileAccess.Write);
                        MessageBox.Show(str, "fs2_str", MessageBoxButtons.OK, MessageBoxIcon.None);
                        StreamWriter sw = new StreamWriter(fs2,append:false);
                        sw.WriteLine(con);
                        MessageBox.Show(con, "fs2_write_con", MessageBoxButtons.OK, MessageBoxIcon.None);
                        sw.Close();
                        fs2.Close();
        //FileStream fs = new FileStream(str, FileMode.Open, FileAccess.ReadWrite);
        //MessageBox.Show(con, "con_del_shapeId", MessageBoxButtons.OK, MessageBoxIcon.None);
        //MessageBox.Show(path.ToString(), "con_del_comment", MessageBoxButtons.OK, MessageBoxIcon.None);
        //MessageBox.Show(copath, "comments_copath", MessageBoxButtons.OK, MessageBoxIcon.None);
        //MessageBox.Show(con, "con_del_comment", MessageBoxButtons.OK, MessageBoxIcon.None);
        //File.Copy(sourcepath, afterpath, true);
        
            //原button2_code
            DialogResult ztx = MessageBox.Show("请确认已将指定文件复制到压缩包内！继续请点是，其他请点否", "提示", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (ztx == DialogResult.Yes)
                {
                OpenFileDialog transpath = new OpenFileDialog();
                //文件类型过滤
                transpath.Filter = @"待转化压缩包|*.xlsx";
                DialogResult resultt = transpath.ShowDialog();
                if (resultt == DialogResult.OK)
                {
                    //获取打开文件的路径（全路径，含文件名）
                    int fnlength2 = transpath.ToString().LastIndexOf("FileName");
                    string sFile2 = transpath.ToString().Substring(fnlength2 + 9, transpath.ToString().Length - fnlength2 - 9);
                    //开始修改文件扩展名
                    Path.GetExtension(sFile2);
                    string dd = Path.GetExtension(sFile2);
                    Path.ChangeExtension(sFile2, "xlsx");
                    string fFile = Path.ChangeExtension(sFile2, ".xlsx");
                    FileInfo fn = new FileInfo(sFile2);
                    //fn.MoveTo(fFile);

                    string directory = "\\xl";
                    string expath = transpath.FileName + @"_Fix";
                    string copath = expath + "\\xl\\";
                    string[] pathFile = Directory.GetFiles(copath, "comments*.xml", SearchOption.TopDirectoryOnly);

                    try
                    {
                        CZip.ZipFile zip = new CZip.ZipFile(sFile2);
                        //zip.AddDirectory(pathFile,);
                        zip.BeginUpdate();
                        foreach (string str in pathFile)
                        {

                        }
                        zip.CommitUpdate();
                    }
                    catch(Exception ex)
                    {
                        throw ex;
                    }
                    MessageBox.Show("恢复完成！请使用Excel2007打开文件！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                   

            }
           */


        private void Label1_Click(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void Button2_Click(object sender, EventArgs e)
        //关于
        {
            MessageBox.Show("程序使用C#编写\n(需要安装.Net Framework 4.5环境才可正常运行)\nZip功能来自icsharpcode/SharpZipLib开源项目,项目作者及贡献者版权所有" +
                "\n支持修复注释的版本:Excel2013\n(Excel2016修复一次后仍需要Excel程序再次修复)\nBy Raven\n\n" +
                "历史更新:" +
                "\n(2019.10.16)修复程序逻辑性bug" +
                "\n(2019.10.17)修正替换内容后的代码遗留问题" +
                "\n(2019.10.18)修复<一键修复>功能存在的功能性错误" +
                "\n(2020.11.03)修正描述错误,修正由于疏漏没有xlsm格式可选的问题",
                "关于Excel批注修复工具", MessageBoxButtons.OK, MessageBoxIcon.None);
        }

        private void Button3_Click(object sender, EventArgs e)
        {

        }

        private void Button4_Click(object sender, EventArgs e)
        {
            MessageBox.Show("暂不支持对Excel2016以上版本xlsx文件进行批注修复", "注意事项", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void Button5_Click(object sender, EventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.Multiselect = true;//等于true表示可以选择多个文件
            dlg.DefaultExt = ".xlsx";
            dlg.Filter = "待删除批注的Excel文档|*.xlsx|待删除批注带有宏的Excel文档|*.xlsm";
            if (dlg.ShowDialog() == DialogResult.OK)
            {

                foreach (string file in dlg.FileNames)
                {
                    MessageBox.Show(file);
                    string sfile = Path.ChangeExtension(file, ".zip");
                    FileInfo fi = new FileInfo(file);
                    fi.MoveTo(sfile);
                    CZip.ZipFile zip = new CZip.ZipFile(sfile);
                    zip.BeginUpdate();
                    zip.Delete(@"xl\comments*.xml");
                    zip.CommitUpdate();
                }
            }
        }

        private void button3_Click_1(object sender, EventArgs e)
        {

        }
    }
}


        

