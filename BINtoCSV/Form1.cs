using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Data.OleDb;
using System.Diagnostics;
using System.Threading;
using System.Globalization;
using System.Net;

namespace BINtoCSV
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        List<double> kkk = new List<double>();
        float ppp;
        int Ttimes = 1, Count = 0;
        int Direction, timesCountCW = 0, timesCountCCW = 0, CWcount = 0, CCWcount = 0;
        int num = 1200;
        int pl = 1;
        private void button3_Click(object sender, EventArgs e)
        {
           
        }

        private void button4_Click(object sender, EventArgs e)
        {
            for(int i = 0; i < 2; i++)
            {
                FindLoadData1 = FindLoad(3+i);
            }
           
        }
        List<string> FindLoadData = new List<string>();
        List<string> FindLoadData1 = new List<string>();
        public List<string> FindLoad( int Num)
        {
            OleDbConnection oleDB = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\pool0\Desktop\pool\pool.mdb");
            string sql = "select * from LoadData WHERE times = 297 AND judgment = 'Max'";
            //獲取表1中暱稱為LanQ的內容
            OleDbDataAdapter dbDataAdapter = new OleDbDataAdapter(sql, oleDB); //建立適配物件
            DataTable dt = new DataTable(); //新建表物件
            dbDataAdapter.Fill(dt); //用適配物件填充表物件
            FindLoadData.Clear();
            foreach (DataRow item in dt.Rows)
            {
                // MessageBox.Show(item[0] + " | " + item[1] + " | " + item[2] + " | " + item[3] + " | " + item[4] + " | " + item[5] + " | " + item[6] + " | " + item[7]);
                FindLoadData.Add(item[Num].ToString());

            }
            return FindLoadData;
        }
        string SerialBuffer = "", StopCodeNum = "";
        private void button5_Click(object sender, EventArgs e)
        {
            SerialBuffer = "00000D000000";
            int BufferNum = SerialBuffer.IndexOf("0D");
            StopCodeNum = LeftCut(SerialBuffer, BufferNum+2, 12- (BufferNum+2));//切割出分割碼
        }
        bool clicked = false;
        int iOldX, iOldY, iClickX, iClickY;

        private void listBox1_MouseUp(object sender, MouseEventArgs e)
        {
           
        }

        private void listBox1_MouseMove(object sender, MouseEventArgs e)
        {
           
        }
        private void Gp_MouseMove(object sender, MouseEventArgs e)
        {
            if (clicked)
            {
                Point p = new Point(); // New Coordinate
                p.X = e.X + groupBox1.Left;
                p.Y = e.Y + groupBox1.Top;
                groupBox1.Left = p.X - iClickX;
                groupBox1.Top = p.Y - iClickY;
            }
        }
        TextBox[] tbx = new TextBox[100];
        
        private void Form1_Load(object sender, EventArgs e)
        {
            groupBox1.MouseDown += new MouseEventHandler(Gp_MouseDown);
            groupBox1.MouseUp += new MouseEventHandler(Gp_MouseUP);
            groupBox1.MouseMove += new MouseEventHandler(Gp_MouseMove);
           for(int i=0; i < 100; i++)
            {
                tbx[i] = new TextBox();
                tbx[i].Text = "0";
                tbx[i].Location = new Point(0, i*10);
                this.Controls.Add(tbx[i]);
                tbx[i].Visible = false;
            }

        }
        private void Gp_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                Point p = ConvertFromChildToForm(e.X, e.Y, groupBox1);
                iOldX = p.X;
                iOldY = p.Y;
                iClickX = e.X;
                iClickY = e.Y;
                clicked = true;
            }
        }
        private void Gp_MouseUP(object sender, MouseEventArgs e)
        {
            clicked = false;
        }
        List<double> pool = new List<double>();
        int[][] lkk = new int[10][]; 
        private void button6_Click(object sender, EventArgs e)
        {
            lkk[0][0] = 1;
            lkk[1][0] = 2;
            lkk[2][0] = 3;
            pool.Add(-1);
            pool.Add(-2);
            pool.Add(-3);
            pool.Add(-4);
            pool.Add(-5);
            MessageBox.Show(pool.IndexOf(2).ToString());
        }

        private void listBox1_MouseDown(object sender, MouseEventArgs e)
        {
           
        }
        private void button7_Click(object sender, EventArgs e)
        {
            LoadDataUpdate(1, "L", 1, 1, 1);
        }
        public void LoadDataUpdate(long times, string judgment, double angle, double test_data, int SpecType)
        {
            OleDbConnection oleDB = new OleDbConnection(@"Provider = Microsoft.Jet.OLEDB.4.0; Data Source = C:\Users\pool0\Desktop\test.mdb");
            string sql = @"update LoadData set judgment = '" + judgment + "'," + "angle = " + angle + "," + "test_data = " + test_data + "," + "SpecType = " + SpecType + "," + " where times =" + times;
            oleDB.Open();
            // string sql = @"insert into LoadData (times,judgment,angle,test_data,SpecType ) values (" + times + ",'" + judgment + "'," + angle + "," + test_data + "," + SpecType + ")";
            
            OleDbCommand oleDbCommand = new OleDbCommand(sql, oleDB);
            oleDbCommand.ExecuteNonQuery(); //返回被修改的數目
            oleDB.Close();
        } //儲存測定值
        long abs = 22222442;
        int kkl = 200;
        double dd1, dd2;
        private void button8_Click(object sender, EventArgs e)
        {
            dd1 = 10;
            dd2 = 4.4444;
            using (BinaryWriter bwr = new BinaryWriter(new FileStream(@"C:\Users\pool0\Desktop\123\123.set", FileMode.OpenOrCreate, FileAccess.ReadWrite)))
            {
                bwr.BaseStream.Seek(bwr.BaseStream.Length, 0);
                bwr.Write(dd1);
               // bwr.Write(abs);
               // bwr.Write(dd2);
               // bwr.Write("X");
               // bwr.Write(kkl.ToString("00000"));
             //   bwr.Write(5);
              //  bwr.Write(6);
             //   bwr.Write(7);
                bwr.Close();
                bwr.Dispose();
                timesCountCW += 1;
            }
        }
        byte[] buff = new byte[600];
        int nm, mm, mm1, mm2, mm3, mm4, mm5, mm6;
        List<string> momo = new List<string>();
        List<double> TestKK = new List<double>();
        private void button12_Click(object sender, EventArgs e)
        {
            TestKK.Add(3.3);
            TestKK.Add(3.4);
            TestKK.Add(3.5);
        }

        private void button13_Click(object sender, EventArgs e)
        {
            MessageBox.Show(TestKK.Count().ToString());

        }

        private void button14_Click(object sender, EventArgs e)
        {
            TestKK.RemoveAt(1);
        }

        private void button15_Click(object sender, EventArgs e)
        {
            MessageBox.Show(TestKK[1].ToString());
        }

        private void comboBox13_MouseDown(object sender, MouseEventArgs e)
        {
            comboBox12.Visible = true;
        }

        private void comboBox12_MouseDown(object sender, MouseEventArgs e)
        {
            comboBox11.Visible = true;
        }

        private void button16_Click(object sender, EventArgs e)
        {
            for(int i = 0; i < 100; i++)
            {
                tbx[i].Visible = true;
            }
        }
        List<double> kkkp = new List<double>();
        List<double> kkk1p = new List<double>();
        List<double> kkk2p = new List<double>();
        string localPath = @"C:\Users\pool0\Desktop\FTP\";
        string fileName;
        string ftpSeverIP = "192.168.0.1";
        double klk;
        private void button17_Click(object sender, EventArgs e)
        {
            /*
            OpenFileDialog open = new OpenFileDialog();
            open.CheckFileExists = false;
            if (open.ShowDialog() == DialogResult.OK)
            {
               
            }
            //Upload("456.txt");
            Upload(open.FileName);
            */
            for(double i = 0; i < 1000000; i++)
            {
                klk = i;
            }
        }
        FtpWebRequest reqFTP;
        private void Upload(string filename)
        {
            FileInfo fileinf = new FileInfo(filename);
            string uri = "ftp://" + ftpSeverIP + "/" + fileinf.Name;
            Connect(uri);
            reqFTP.KeepAlive = false;
            reqFTP.Method = WebRequestMethods.Ftp.UploadFile;
            reqFTP.ContentLength = fileinf.Length;
            int bufflength = 2048;
            byte[] buff = new byte[bufflength];
            int contentLen;
            FileStream fs = fileinf.OpenRead();
            try
            {
                Stream stm = reqFTP.GetRequestStream();
                contentLen = fs.Read(buff,0,bufflength);
                while(contentLen != 0)
                {
                    stm.Write(buff, 0, contentLen);
                    contentLen = fs.Read(buff,0,bufflength);
                }
                stm.Close();
                fs.Close();
                System.IO.File.Delete(filename);
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void Connect(string path)
        {
            reqFTP = (FtpWebRequest)FtpWebRequest.Create(new Uri(path));
            reqFTP.UseBinary = true;
            reqFTP.Credentials = new NetworkCredential("123","123");
        }
        Thread threadTest, threadTest1, threadTest2, threadTest3, threadTest4;
        private void button18_Click(object sender, EventArgs e)
        {
            threadTest = new Thread(new ThreadStart(Mes));
            threadTest.IsBackground = true;
            threadTest.Start();
        }
        delegate void Pool();
        private void Mes()
        {
            try
            {
                for (int i = 0; i < 10; i++)
                {
                    Pool pl = new Pool(AAA);
                    this.Invoke(pl);
                    Thread.Sleep(1);
                }
                for (int i = 0; i < 30; i++)
                {
                    Pool pl = new Pool(AAA);
                    this.Invoke(pl);
                    Thread.Sleep(1);
                }
            }
            catch (ThreadInterruptedException ex)
            {
              //  MessageBox.Show(ex.Message);
            }
               
            
        }
        private void Mes1()
        {
            try
            {
                while (true)
                {

                    for (int i = 0; i < 10; i++)
                    {
                        Pool pl = new Pool(AAA);
                        this.Invoke(pl);
                        Thread.Sleep(1);
                    }
                    for (int i = 0; i < 30; i++)
                    {
                        Pool pl = new Pool(AAA);
                        this.Invoke(pl);
                        Thread.Sleep(1);
                    }
                }
            }
             catch (ThreadAbortException ex)
            {
                //  MessageBox.Show(ex.Message);
            }

        }
        private void Mes2()
        {
            try
            {
                while (true)
                {
                  //  Thread.Sleep(1);
                }
            }
            catch (ThreadAbortException ex)
            {
                //  MessageBox.Show(ex.Message);
            }

        }
        private void Mes3()
        {
            try
            {
                while (true)
                {

                }
            }
            catch (ThreadInterruptedException ex)
            {
                //  MessageBox.Show(ex.Message);
            }
        }
        private void Mes4()
        {
            try
            {
                while (true)
                {

                }
            }
            catch (ThreadInterruptedException ex)
            {
                //  MessageBox.Show(ex.Message);
            }
        }

      
        int p = 0;

        private void button22_Click(object sender, EventArgs e)
        {
            threadTest2.Abort();
        }

        private void button23_Click(object sender, EventArgs e)
        {
            string klll = "";
        }

        private void AAA()
        {
            p += 1;
                listBox1.Items.Add(p);
                Thread.Sleep(100);
            
        }
        private void button21_Click(object sender, EventArgs e)
        {
            threadTest1.Abort();
        }


        private void button19_Click(object sender, EventArgs e)
        {
            threadTest.Interrupt();
            p = 0;
        }

        private void button20_Click(object sender, EventArgs e)
        {
            threadTest1 = new Thread(new ThreadStart(Mes1));
            threadTest1.IsBackground = true;
            threadTest1.Start();
            threadTest2 = new Thread(new ThreadStart(Mes2));
            threadTest2.IsBackground = true;
            threadTest2.Start();
           
        }

        public static void RefAdd(ref int num1)
        {
            num1 = 10;
        }
        Stopwatch stt = new Stopwatch();
        private void button11_Click(object sender, EventArgs e)
        {
            stt.Start();
            for (int i =0;i<1000;i++)
            {
                momo.Add(DateTime.Now.Ticks.ToString());
                Thread.Sleep(10);
            }
            stt.Stop();
            stt.Reset();
        }

        double aa, aa1, aa2, aa3;
        long len;
        string aa4,aa5;

        private void button10_Click(object sender, EventArgs e)
        {
            FileStream BinaryFile = File.Open(@"C:\Users\pool0\Desktop\123\123.set", FileMode.OpenOrCreate, FileAccess.ReadWrite);
            BinaryReader reader = new BinaryReader(BinaryFile);
            aa4 = reader.ReadString();
            aa1 = reader.ReadInt32();
            aa5 = reader.ReadString();

        }

        private void button9_Click(object sender, EventArgs e)
        {
            FileStream BinaryFile = File.Open(@"C:\Users\pool0\Desktop\123\123.set", FileMode.OpenOrCreate, FileAccess.ReadWrite);
            BinaryWriter bwr = new BinaryWriter(BinaryFile);
            BinaryReader reader = new BinaryReader(BinaryFile);
            len = bwr.BaseStream.Length;
            //  reader.BaseStream.Seek(8,0);
            //   aa = reader.ReadDouble();
            //   aa1 = reader.ReadInt64();
            //  aa2 = reader.ReadDouble();
            // aa3 = reader.ReadDouble();
            for (int i = 0; i < 100; i++)
                {
                   // aa4 = reader.ReadString();
                aa = reader.ReadDouble();
                    if (aa4 == "X")
                    {
                        bwr.Write(322);
                    bwr.Write(kkl.ToString("00000"));
                    break;
                    }
                }
                int kkkkkk = 0;
                int kkkkkk1 = 0;
                int kkkkkk2 = 0;
                int kkkkkk3 = 0;
          
           
            

            reader.Close();
            reader.Dispose();
            /*
            using (BinaryWriter bwr = new BinaryWriter(new FileStream(@"C:\Users\pool0\Desktop\123\123.set", FileMode.OpenOrCreate, FileAccess.ReadWrite)))
            {
                bwr.BaseStream.Seek(8, 0);
                len = bwr.BaseStream.Length;
                bwr.Write(8.0);
                bwr.Close();
                bwr.Dispose();
            }
            */
        }

        private Point ConvertFromChildToForm(int x, int y, Control control)
        {
            Point p = new Point(x, y);
            control.Location = p;
            return p;
        }




        private void button3_MouseDown(object sender, MouseEventArgs e)
        {
           
        }


        public string LeftCut(string str, int st, int len)
        {
            return str.Substring(st, len);
        } //從左側切割字元
        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog open = new OpenFileDialog();
            open.CheckFileExists = false;
            if (open.ShowDialog() == DialogResult.OK)
            {
                file_dbName = Path.GetFileNameWithoutExtension(open.FileName);//path.getfilename用於取出檔案名稱
            }
            FileStream BinaryFile = File.Open(file_dbName + ".data", FileMode.OpenOrCreate, FileAccess.ReadWrite);
            BinaryReader reader = new BinaryReader(BinaryFile);
            for (int i = 0; i < 1 * 2; i++)
            {
                if(i == 64)
                {
                    
                }
                kkk.Clear();
                ppp = 0;
                BinSt = reader.ReadString();
                Direction = reader.ReadInt32();
                BinSt = reader.ReadString();
                if (i <= 64)
                {
                    pok = reader.ReadDouble();
                }
                else
                {
                    pok = reader.ReadDouble();
                }
                if (pok == Ttimes)
                {
                    plk = reader.ReadDouble();
                    kkk.Add(plk);
                    for (int nn = 1; nn < 1200; nn++)
                    {
                        pok = reader.ReadInt32();
                        plk = reader.ReadDouble();
                        kkk.Add(plk);
                    }
                    using (sw = new StreamWriter(file_dbName + ".csv", true, Encoding.Default))
                    {
                        sw.Write("角度,正反轉(1為正2為反),扭力,次數\r\n");
                        foreach (double kl in kkk)
                        {
                            sw.Write(ppp + "," + Direction + "," + Math.Round(kl, 3) + " ," + pok + "," + "\r\n");
                            ppp += 0.1f;
                        }
                    }

                    sw.Close();
                    if (Direction == 1)
                    {
                        Ttimes += 1;
                    }
                }
            }
            reader.Close();
            BinaryFile.Close();
        }

        string BinSt, file_dbName;
        double pok, plk;
        StreamWriter sw;
        List<double> kkk1 = new List<double>();
        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog open = new OpenFileDialog();
            open.CheckFileExists = false;
            if (open.ShowDialog() == DialogResult.OK)
            {
                file_dbName = Path.GetFileNameWithoutExtension(open.FileName);//path.getfilename用於取出檔案名稱
            }
            FileStream BinaryFile = File.Open(file_dbName + ".data", FileMode.OpenOrCreate, FileAccess.ReadWrite);
            BinaryReader reader = new BinaryReader(BinaryFile);
            for (int i = 0; i < 1 * 2; i++)
            {
                kkk.Clear();
                ppp = 0;
                BinSt = reader.ReadString();
                Direction = reader.ReadInt32();
                BinSt = reader.ReadString();
                plk = reader.ReadDouble();
                pok = reader.ReadInt32();
                if (pok == Ttimes)
                {
                    plk = reader.ReadDouble();
                    kkk.Add(plk);
                    for (int nn = 1; nn < 1200; nn++)
                    {
                        pok = reader.ReadInt32();
                        plk = reader.ReadDouble();
                        kkk.Add(plk);
                      //  kkk1.Add();
                    }
                    using ( sw = new StreamWriter(file_dbName + ".csv", true, Encoding.Default))
                    {
                        sw.Write("角度,正反轉(1為正2為反),扭力,次數\r\n");
                        foreach (double kl in kkk)
                        {
                            sw.Write(ppp + "," + Direction + "," + Math.Round(kl, 3) + " ," + pok + "," + "\r\n");
                            ppp += 0.1f;
                        }
                    }

                    sw.Close();
                            if (Direction == 1)
                            {
                                Ttimes += 1;
                            }
                }
            }
            reader.Close();
            BinaryFile.Close();
        }
    }
}
