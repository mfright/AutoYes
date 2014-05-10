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


        int currentPointX, currentPointY; //現在のカーソルの座標
        string messageKanshichu;


        public Form1()
        {
            InitializeComponent();
            messageKanshichu = lblStatus.Text;
        }


        private void timerAutoClick_Tick(object sender, EventArgs e)
        {

            // セキュリティダイアログの座標を取得
            IntPtr dialog = FindWindow("#32770", "Microsoft Office Outlook");
            Rect dialogRect = new Rect();
            GetWindowRect(dialog, ref dialogRect);
            toolStripStatusLabel1.Text = "LEFT:" + dialogRect.Left + "   TOP:" + dialogRect.Top + "   RIGHT:"+dialogRect.Right + "   BOTTOM:"+dialogRect.Bottom;

            // 座標が0,0であれば、ダイアログは出ていないと考えられるので、処理を終了。
            if (dialogRect.Right == 0 && dialogRect.Bottom == 0)
            {
                if (lblStatus.Text != messageKanshichu)
                {
                    lblStatus.Text = messageKanshichu;
                }
                return;
            }

            lblStatus.Text = "ダイアログ発見\r\nクリック中！";

            //現在のカーソル位置を記憶しておく
            //currentPointX = System.Windows.Forms.Cursor.Position.X;
            //currentPointY = System.Windows.Forms.Cursor.Position.Y;

            //カーソルを"はい"ボタン上まで移動させ、クリックする
            setPoint(dialogRect.Left + 116, dialogRect.Top + 156);
            doClick();

        }


        //クリックする処理
        void doClick()
        {
            //[System.Runtime.InteropServices.DllImport("USER32.DLL")]
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
        
    }
}
