using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using System.Runtime.InteropServices;
using System.Windows;
using System.Diagnostics;

namespace AutoYES1
{
    public partial class Form1 : Form
    {
        // Win32 API
        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern IntPtr FindWindow(string strClassName, string strWindowName);

        [DllImport("user32.dll")]
        public static extern bool GetWindowRect(IntPtr hwnd, ref Rect rectangle);

        public struct Rect
        {
            public int Left { get; set; }
            public int Top { get; set; }
            public int Right { get; set; }
            public int Bottom { get; set; }
        }

        [DllImport("USER32.DLL")]
        private static extern void mouse_event(int dwFlags, int dx, int dy, int cButtons, int dwExtraInfo);


        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool SetForegroundWindow(IntPtr hWnd);



        // kernel32 APIを使用する。
        [DllImport("kernel32.dll")]
        extern static ExecutionState SetThreadExecutionState(ExecutionState esFlags);

        //引数のExecutionState列挙体
        [FlagsAttribute]
        public enum ExecutionState : uint
        {
            // Return value of failed.
            Null = 0,

            // Anti standby.
            SystemRequired = 1,

            // Anti Display-off
            DisplayRequired = 2,

            // continuous
            Continuous = 0x80000000,
        }

        public int dialogLeft=0, dialogTop = 0;


        int currentPointX, currentPointY; //現在のカーソルの座標
        string messageKanshichu;

        conf loadedConf;    //iniから読み込んだ設定情報


        public Form1()
        {
            InitializeComponent();
            messageKanshichu = lblStatus.Text;

            loadedConf = loadIni();
        }


        private void timerAutoClick_Tick(object sender, EventArgs e)
        {
         
            // セキュリティダイアログの座標を取得
            //IntPtr dialog = FindWindow(null, "Microsoft Office Outlook");
            IntPtr dialog = FindWindow(null, loadedConf.windowName);
            
            Rect dialogRect = new Rect();
            GetWindowRect(dialog, ref dialogRect);
            toolStripStatusLabel1.Text = "LEFT:" + dialogRect.Left + "   TOP:" + dialogRect.Top + "   RIGHT:"+dialogRect.Right + "   BOTTOM:"+dialogRect.Bottom;

            //ダイアログのサイズ
            int dialogHeight = dialogRect.Bottom - dialogRect.Top;
            //int dialogWidth = dialogRect.Right = dialogRect.Left;
            
            //Console.WriteLine("dialogRect.Bottom:" + dialogRect.Bottom);
            //Console.WriteLine("dialogRect.Top:" + dialogRect.Top);

            //Console.WriteLine("CW:" + loadedConf.windowWide + " CH:" + loadedConf.windowHigh);
            //Console.WriteLine(" TH:" + dialogHeight);


            // 座標が0,0であれば、ダイアログは出ていないと考えられるので、処理を終了。
            if (dialogRect.Right == 0 && dialogRect.Bottom == 0)
            {
                if (lblStatus.Text != messageKanshichu)
                {
                    lblStatus.Text = messageKanshichu;
                    this.BackColor = Color.White;
                    this.Opacity = 0.75;
                }
                return;
            } else if(
                // ダイアログは所定のサイズを満たしているか確認
                loadedConf.windowHigh+loadedConf.windowMargin >= dialogHeight &&
                loadedConf.windowHigh - loadedConf.windowMargin <= dialogHeight /* &&
                loadedConf.windowWide + loadedConf.windowMargin >= dialogWidth &&
                loadedConf.windowWide - loadedConf.windowMargin <= dialogWidth */)
            {
                // GO-ON
            }
            else
            {
                if (lblStatus.Text != messageKanshichu)
                {
                    lblStatus.Text = messageKanshichu;
                    this.BackColor = Color.White;
                    this.Opacity = 0.75;
                }
               return;
            }//*/

            

            lblStatus.Text = "ダイアログ発見\r\nクリック中！";
            this.BackColor = Color.Yellow;
            this.Opacity = 100;


            // Outlookをアクティブにする
            activateOutlook();


            //カーソルを"はい"ボタン上まで移動させ、クリックする
            //setPoint(dialogRect.Left + 116, dialogRect.Top + 156);
            //doClick();

            dialogLeft = dialogRect.Left;
            dialogTop = dialogRect.Top;
            timerDelayClick.Start();

        }


        void activateOutlook()
        {

            try
            {
                
                //user32.dll
                foreach (System.Diagnostics.Process p
                    in System.Diagnostics.Process.GetProcesses())
                {
                    //"Outlook"がメインウィンドウのタイトルに含まれているか調べる
                    if (0 <= p.MainWindowTitle.IndexOf("Outlook"))
                    {
                        //ウィンドウをアクティブにする
                        SetForegroundWindow(p.MainWindowHandle);
                        break;
                    }
                }

            }
            catch (Exception ex)
            {

            }
        }


        //クリックする処理
        void doClick()
        {
            
            int MOUSEEVENTF_LEFTDOWN = 0x2;
            int MOUSEEVENTF_LEFTUP = 0x4;
            mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0);
            mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
        }

        //カーソルを移動させる処理
        void setPoint(int newX, int newY)
        {
            System.Windows.Forms.Cursor.Position = new System.Drawing.Point(newX, newY);
        }




        // 画面ロックを50秒ごとに抑止する
        private void timerAutolockCancel_Tick(object sender, EventArgs e)
        {
            // DisplayRequiredをSetThreadExecutionStateへ送信.(スクリーンロックを抑止)
            ExecutionState es = new ExecutionState();
            es = ExecutionState.DisplayRequired;
            SetThreadExecutionState(es);
        }

        private void timerDelayClick_Tick(object sender, EventArgs e)
        {
            //setPoint(dialogLeft + 116, dialogTop + 156);
            setPoint(dialogLeft + loadedConf.LeftOffset, dialogTop + loadedConf.TopOffset);

            doClick();
            doClick();

            timerDelayClick.Stop();
        }


        // iniファイルを読み込み、内容を返す。
        public conf loadIni()
        {
            System.IO.StreamReader sr = new System.IO.StreamReader(@"settings.ini",
                System.Text.Encoding.GetEncoding("shift_jis"));
            
            //内容をすべて読み込む
            string s = sr.ReadToEnd();
            
            //閉じる
            sr.Close();

            string[] stArrayData = s.Split(',','\r');

            conf config = new conf();
            config.windowName = stArrayData[0];
            config.LeftOffset = int.Parse(stArrayData[1]);
            config.TopOffset = int.Parse(stArrayData[2]);
            config.windowWide = int.Parse(stArrayData[3]);
            config.windowHigh = int.Parse(stArrayData[4]);
            config.windowMargin = int.Parse(stArrayData[5]);
            
            return config;
        }


        // iniファイルの中身を読み込むインスタンス
        public class conf
        {
            public string windowName = "";
            public int LeftOffset = 0;
            public int TopOffset = 0;
            public int windowWide = 0;
            public int windowHigh = 0;
            public int windowMargin = 0;
        }
    }
}
