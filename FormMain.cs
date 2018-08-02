using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using DevExpress.XtraSplashScreen;
using Microsoft.Win32;
using DevExpress.LookAndFeel;
using ParacBp.dlg;
using System.Threading;
using System.Net.NetworkInformation;

namespace ParacBp
{
    public partial class FormMain : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        private DateTime startTime; // 软件启动的时间
        private DlgNull dlgNull = new DlgNull();
        // User Model
        private DlgSingleMonitor dlgSingleMonitor = new DlgSingleMonitor();
        private DlgGroupMonitor dlgGroupMonitor = new DlgGroupMonitor();
        private DlgModelMonitor dlgModelMonitor = new DlgModelMonitor();
        private DlgMapMonitor dlgMapMonitor = new DlgMapMonitor();
        private DlgSystemLog dlgSystemLog = new DlgSystemLog();

        // Admin Model
        private DlgOpiuSetting dlgOpuiSetting = new DlgOpiuSetting();
        private DlgSingleSetting dlgSingleSetting = new DlgSingleSetting();
        private DlgGroupSetting dlgGroupSetting = new DlgGroupSetting();
        private DlgModelSetting dlgModelSetting = new DlgModelSetting();
        private DlgMapSetting dlgMapSetting = new DlgMapSetting();
        private DlgSchedueSetting dlgSchedueSetting = new DlgSchedueSetting();
        DataTable dt;

        public FormMain()
        {
            InitializeComponent();


          

            Thread t = new Thread(ini);
            t.IsBackground = true;
            t.Start();

           

        }
        private void ini()
        {
            string queryString = "select SysId,IpAddr,ArmName from SysArm where IsUse = '1' order by cast(SysId as int) asc";
            dt = CommonlyFunctions.getSchedue(queryString);

            Ping ping = new Ping();
            if (dt.Rows.Count > 0)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    string ip = dt.Rows[i][1].ToString();
                    PingReply pingReply = ping.Send(ip);
                    if (pingReply.Status == IPStatus.Success)
                    {
                        Thread matchthread = new Thread(new ParameterizedThreadStart(match));
                        matchthread.IsBackground = true;
                        matchthread.Start(i);
                    }
                }

            }

            Thread t = new Thread(PingOpiu);
            t.IsBackground = true;
            t.Start();
           
           
        }

        private void PingOpiu()
        {
            while (true)
            {
                if (dt.Rows.Count > 0)
                {
                    try
                    {
                        Ping ping = new Ping();
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            string ip = dt.Rows[i][1].ToString();
                            string armname = dt.Rows[i][2].ToString();

                            PingReply pingReply = ping.Send(ip);
                            if (pingReply.Status == IPStatus.Success)
                            {
                                if(this.IsHandleCreated)
                                this.Invoke((MethodInvoker)delegate
                                                {
                                                    this.barStaticItemStatus.Caption = armname + "信号正常";
                                                    this.barStaticItemStatus.ItemAppearance.Normal.ForeColor = Color.LawnGreen;
                                                });


                            }
                            else
                            {
                                if (this.IsHandleCreated)
                                this.Invoke((MethodInvoker)delegate
                                                {
                                                    this.barStaticItemStatus.Caption = armname + "信号异常";
                                                    this.barStaticItemStatus.ItemAppearance.Normal.ForeColor = Color.Red;
                                                });

                            }
                            Thread.Sleep(3000);
                        }
                    }
                    catch(Exception ex)
                    {
                        CommonlyFunctions.insertlog(Guid.NewGuid().ToString(),"-","0","匹配OPIU连接状态异常",DateTime.Now,"-","-",ex.Message.ToString());
                    }

                }
                Thread.Sleep(10);

            }

        }
           
        private void match(object i)
        {

  
            string armid="-";
            try
            {
                int[] status = new int[455];
                //List<string> updatestrings = new List<string> { };
                int row = (int)i;
                string ip = dt.Rows[row][1].ToString();
                armid = dt.Rows[row][0].ToString();
               
                ModbusTcpHelper mt = new ModbusTcpHelper(ip, 502);
                 mt.Connect();
                 
                 if (mt.GetConnect())
                 {
                    byte[] return_bytes;
                    mt.discover_loopandgroup();
             
                    while (true)
                    {
                        Thread.Sleep(10);
                        if (mt.GetReturnData() != null)
                        {
                            return_bytes = mt.GetReturnData();
                            break;
                        }
                    } 
                
                    if (return_bytes.Length == 57)
                    {
                      
                        byte[] loop_bytes = new byte[48];
                       
                        for (int j = 0; j < loop_bytes.Length;j++)
                        {

                            if (j < 47)
                            {
                                loop_bytes[j] = return_bytes[j + 9];
                                string temp = Convert.ToString(loop_bytes[j], 2);
                                char[] bin = stringtobin(temp);
                                for (int k = 0; k < 8; k++)
                                {

                                    status[8 * j + k] = Convert.ToInt32(bin[7 - k].ToString());
                                   


                                }
                            }
                            else if (j == 47)
                            {
                                loop_bytes[j] = return_bytes[j + 9];
                                string temp = Convert.ToString(loop_bytes[j], 2);
                                char[] bin = stringtobin(temp);
                                for (int k = 0; k < 7; k++)
                                {

                                    status[8 * j + k] = Convert.ToInt32(bin[7 - k].ToString());
                                   

                                }
                            }
                         
                        }
                       
                        
                    }

                 }
             
                 mt.disconnect();
                 Thread.Sleep(50);

                 ModbusTcpHelper mt_model = new ModbusTcpHelper(ip, 502);
                 mt_model.Connect();
                 if (mt_model.GetConnect())
                 {
               
                     byte[] return_bytes_model;
                     mt_model.discover_model();
                  
                     while (true)
                     {
                         Thread.Sleep(10);
                         if (mt_model.GetReturnData() != null)
                         {
                             return_bytes_model = mt_model.GetReturnData();
                             break;
                         }
                     }
    
                     if (return_bytes_model.Length == 18)
                     {
                          byte[] model_bytes = new byte[9];

                          for (int j = 0; j < model_bytes.Length; j++)
                          {


                              model_bytes[j] = return_bytes_model[j + 9];
                              string temp = Convert.ToString(model_bytes[j], 2);
                              char[] bin = stringtobin(temp);
                              for (int k = 0; k < 8; k++)
                              {
                                  status[8 * j + k + 383] = Convert.ToInt32(bin[7 - k].ToString());
                              }
                          }

                     }
                 }
            
                 CommonlyFunctions.ExecuteSqlTran(armid, status);
                
                
            }
            catch (Exception ex)
            {
                CommonlyFunctions.insertlog(Guid.NewGuid().ToString(), armid, "0", "匹配OPIU寄存器状态异常", DateTime.Now, "-", "-", ex.Message.ToString());
            }
           
        }

        private char[] stringtobin(string s)
        {
            char[] j = new char[8];
            char[] k = s.ToCharArray();
            int n = j.Length - k.Length;
            if (n == 0)
            {
                for (int i = 0; i < j.Length; i++)
                {
                    j[i] = k[i];

                }
            }
            else if (n > 0)
            {
                for (int i = 0; i < n; i++)
                {
                    j[i] = '0';
                }
                for (int i = n; i < 8; i++)
                {
                    j[i] = k[i - n];
                }
            }
            return j;
        }
        private void auto_run()
        {
            try
            {
                string path = Application.ExecutablePath;
                RegistryKey rk = Registry.LocalMachine;
                RegistryKey rk2 = rk.CreateSubKey(@"Software\Microsoft\Windows\CurrentVersion\Run");
                rk2.SetValue("ParacBp_autoRun", path);
                rk2.Close();
                rk.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("设置软件自启动，请关闭360等安全防护软件！");
            }
           
        }

        private void cancel_autorun()
        {
            string path = Application.ExecutablePath;
            RegistryKey rk = Registry.LocalMachine;
            RegistryKey rk2 = rk.CreateSubKey(@"Software\Microsoft\Windows\CurrentVersion\Run");
            rk2.DeleteValue("ParacBp_autoRun", false);
            rk2.Close();
            rk.Close();
        }

        private void FormMain_Load(object sender, EventArgs e)
        {
            try
            {
                // 软件发布的时候放开这段被注释的代码
                //if (Properties.Settings.Default.AutoRun)
                //{
                //    auto_run();
                //}
                //else
                //{
                //    cancel_autorun();
                //}

                //this.Text = Properties.Settings.Default.SoftwareTitle;


                // 空窗体
                //dlgNull.Show();
                //dlgNull.TopLevel = false;
                //dlgNull.Parent = panelMain;
                //dlgNull.Dock = DockStyle.Fill;
                //panelMain.Controls.Add(dlgNull);

        
                // 地图设定
                dlgMapSetting.Show();
                dlgMapSetting.TopLevel = false;
                dlgMapSetting.Parent = panelMain;
                dlgMapSetting.Dock = DockStyle.Fill;
                panelMain.Controls.Add(dlgMapSetting);
  



     
                // 地图监视
                dlgMapMonitor.Show();
                dlgMapMonitor.TopLevel = false;
                dlgMapMonitor.Parent = panelMain;
                dlgMapMonitor.Dock = DockStyle.Fill;
                panelMain.Controls.Add(dlgMapMonitor);
          

                // 系统日志
                dlgSystemLog.Show();
                dlgSystemLog.TopLevel = false;
                dlgSystemLog.Parent = panelMain;
                dlgSystemLog.Dock = DockStyle.Fill;
                panelMain.Controls.Add(dlgSystemLog);

 
                // OPIU连接设定
                dlgOpuiSetting.Show();
                dlgOpuiSetting.TopLevel = false;
                dlgOpuiSetting.Parent = panelMain;
                dlgOpuiSetting.Dock = DockStyle.Fill;
                panelMain.Controls.Add(dlgOpuiSetting);

      
                // 负载单个设定
                dlgSingleSetting.Show();
                dlgSingleSetting.TopLevel = false;
                dlgSingleSetting.Parent = panelMain;
                dlgSingleSetting.Dock = DockStyle.Fill;
                panelMain.Controls.Add(dlgSingleSetting);

            
                // 负载群组设定
                dlgGroupSetting.Show();
                dlgGroupSetting.TopLevel = false;
                dlgGroupSetting.Parent = panelMain;
                dlgGroupSetting.Dock = DockStyle.Fill;
                panelMain.Controls.Add(dlgGroupSetting);

         
                // 负载模式设定
                dlgModelSetting.Show();
                dlgModelSetting.TopLevel = false;
                dlgModelSetting.Parent = panelMain;
                dlgModelSetting.Dock = DockStyle.Fill;
                panelMain.Controls.Add(dlgModelSetting);


            
                // 日程设定

                dlgSchedueSetting.Show();
                dlgSchedueSetting.TopLevel = false;
                dlgSchedueSetting.Parent = panelMain;
                dlgSchedueSetting.Dock = DockStyle.Fill;
                panelMain.Controls.Add(dlgSchedueSetting);


          
                // 单个监视
                dlgSingleMonitor.Show();
                dlgSingleMonitor.TopLevel = false;
                dlgSingleMonitor.Parent = panelMain;
                dlgSingleMonitor.Dock = DockStyle.Fill;
                panelMain.Controls.Add(dlgSingleMonitor);

       
                // 群组监视
                dlgGroupMonitor.Show();
                dlgGroupMonitor.TopLevel = false;
                dlgGroupMonitor.Parent = panelMain;
                dlgGroupMonitor.Dock = DockStyle.Fill;
                panelMain.Controls.Add(dlgGroupMonitor);

     
                // 模式监视
                dlgModelMonitor.Show();
                dlgModelMonitor.TopLevel = false;
                dlgModelMonitor.Parent = panelMain;
                dlgModelMonitor.Dock = DockStyle.Fill;
                panelMain.Controls.Add(dlgModelMonitor);

       
                //// 程序刚启动的时候显示空白窗体
                //panelMain.Controls.SetChildIndex(dlgNull, 0);
                panelMain.Controls.SetChildIndex(dlgMapMonitor, 0);
                this.Text = Properties.Settings.Default.SoftwareTitle + " — 地图监视";


              
                if (CommonlyFunctions.LoginUserType == true)
                {
                    ribbonSetting.Visible = true;
                    ribbonControlMain.SelectedPage = ribbonMain;
                }
                else
                {
                    ribbonSetting.Visible = false;
                    ribbonControlMain.SelectedPage = ribbonMain;
                }
                startTime = DateTime.Now;
                timerRun.Enabled = true;
                timerRun.Start();

                SplashScreenManager.CloseForm();

            }
            catch (Exception ex)
            {
                CommonlyFunctions.insertlog(Guid.NewGuid().ToString(), "-", "0", "软件开启加载窗体异常", DateTime.Now, "-", "-", ex.Message.ToString());
            }
        }

        private void ribbonControlMain_Paint(object sender, PaintEventArgs e)
        {
            DevExpress.XtraBars.Ribbon.ViewInfo.RibbonViewInfo ribbonViewInfo = ribbonControlMain.ViewInfo;
            if (ribbonViewInfo == null)
            {
                return;
            }

            DevExpress.XtraBars.Ribbon.ViewInfo.RibbonPanelViewInfo panelViewInfo = ribbonViewInfo.Panel;
            if (panelViewInfo == null)
            {
                return;
            }

            Rectangle bounds = panelViewInfo.Bounds;
            int minX = bounds.X;
            DevExpress.XtraBars.Ribbon.ViewInfo.RibbonPageGroupViewInfoCollection groups = panelViewInfo.Groups;
            if (groups == null)
            {
                return;
            }
            if (groups.Count > 0)
            {
                minX = groups[groups.Count - 1].Bounds.Right;
            }

            SizeF sizeF = e.Graphics.MeasureString(CommonlyFunctions.loginUser, new Font("仿宋", 16, FontStyle.Bold));

            int offset = (int)((bounds.Height - sizeF.Height) / 2);
            int width = (int)sizeF.Width + 5;
            bounds.X = bounds.Width - width;
            if (bounds.X < minX)
            {
                return;
            }

            bounds.Width = width;
            bounds.Y += offset + 10;
            bounds.Height = (int)sizeF.Height;

            Font drawFont = new Font("仿宋", 16, FontStyle.Bold);
            SolidBrush drawBrush = new SolidBrush(Color.Black);
            e.Graphics.DrawString(CommonlyFunctions.loginUser, drawFont, drawBrush, (PointF)bounds.Location);

        }

        private void ribbonStatusBarMain_Paint(object sender, PaintEventArgs e)
        {
            Rectangle statusRect = ribbonStatusBarMain.Bounds;
            SizeF sizeF = e.Graphics.MeasureString(Properties.Settings.Default.VersionString, new Font("宋体", 9, FontStyle.Regular));

            Font drawFont = new Font("宋体", 9, FontStyle.Regular);
            SolidBrush drawBrush = new SolidBrush(Color.Black);
            e.Graphics.DrawString(Properties.Settings.Default.VersionString, drawFont, drawBrush, new PointF((statusRect.Width - sizeF.Width) / 2, 2 + (statusRect.Height - sizeF.Height) / 2));

        }

        private void skinRibbonGalleryBarItem1_GalleryItemClick(object sender, DevExpress.XtraBars.Ribbon.GalleryItemClickEventArgs e)
        {
            // 换皮肤
            string SkinValue = e.Item.Caption;
            Properties.Settings.Default.SoftwareSkin = SkinValue;
            Properties.Settings.Default.Save();
            DefaultLookAndFeel Custom = new DefaultLookAndFeel();
            Custom.LookAndFeel.SetSkinStyle(SkinValue);
        }

        private void timerRun_Tick(object sender, EventArgs e)
        {
            // 软件运行时间
            TimeSpan tSpan = DateTime.Now - startTime;
            string spanTime;
            if (tSpan.Days > 0)
            {
                spanTime = string.Format("{0}天 {1:D2}:{2:D2}:{3:D2}", tSpan.Days, tSpan.Hours, tSpan.Minutes, tSpan.Seconds);
            }
            else
            {
                spanTime = string.Format("{0:D2}:{1:D2}:{2:D2}", tSpan.Hours, tSpan.Minutes, tSpan.Seconds);
            }

            barStaticItemRun.Caption = spanTime;
        }

        private void FormMain_FormClosing(object sender, FormClosingEventArgs e)
        {
            timerRun.Stop();
            timerRun.Enabled = false;

            CommonlyFunctions.closeDatabase();
        }

        private void barButtonItemSwitch_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            // 用户切换
            LoginForm dlg = new LoginForm();
            dlg.setRunModel(false);
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                if (CommonlyFunctions.LoginUserType == true)
                {
                    ribbonSetting.Visible = true;
                    ribbonControlMain.SelectedPage = ribbonMain;
                }
                else
                {
                    ribbonSetting.Visible = false;
                    ribbonControlMain.SelectedPage = ribbonMain;
                }

                panelMain.Controls.SetChildIndex(dlgNull, 0);
            }
        }

        private void barButtonItemBugReport_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            // 报告bug

        }

        private void barButtonItemHelp_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            // 使用帮助

        }

        private void barButtonItemAbout_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            // 关于
            FormAbout dlg = new FormAbout();
            dlg.ShowDialog();
        }

        private void barButtonItemSingleMonitor_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            // 单个监视
            panelMain.Controls.SetChildIndex(dlgSingleMonitor, 0);
            this.Text = Properties.Settings.Default.SoftwareTitle + " — 单个监视";
        }

        private void barButtonItemMultiMonitor_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            // 群组监视
            panelMain.Controls.SetChildIndex(dlgGroupMonitor, 0);
            this.Text = Properties.Settings.Default.SoftwareTitle + " — 群组监视";
        }

        private void barButtonItemModelMonitor_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            // 模式监视
            panelMain.Controls.SetChildIndex(dlgModelMonitor, 0);
            this.Text = Properties.Settings.Default.SoftwareTitle + " — 模式监视";
        }

        private void barButtonItemMapMonitor_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            // 地图监视
            panelMain.Controls.SetChildIndex(dlgMapMonitor, 0);
            this.Text = Properties.Settings.Default.SoftwareTitle + " — 地图监视";
        }

        private void barButtonItemSystemLog_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            // 系统记录
            panelMain.Controls.SetChildIndex(dlgSystemLog, 0);
            this.Text = Properties.Settings.Default.SoftwareTitle + " — 系统记录";
        }

        private void barButtonItemOpiuSetting_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            // OPIU连接设定
            panelMain.Controls.SetChildIndex(dlgOpuiSetting, 0);
            this.Text = Properties.Settings.Default.SoftwareTitle + " — OPIU连接设定";
        }

        private void barButtonItemSingleSetting_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            // 负载单个设定
            panelMain.Controls.SetChildIndex(dlgSingleSetting, 0);
            this.Text = Properties.Settings.Default.SoftwareTitle + " — 负载(单个)设定";
        }

        private void barButtonItemMultiSetting_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            // 负载群组设定
            panelMain.Controls.SetChildIndex(dlgGroupSetting, 0);
            this.Text = Properties.Settings.Default.SoftwareTitle + " — 负载(群组)设定";
        }

        private void barButtonItemModelSetting_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            // 负载模式设定
            panelMain.Controls.SetChildIndex(dlgModelSetting, 0);
            this.Text = Properties.Settings.Default.SoftwareTitle + " — 负载(模式)设定";
        }

        private void barButtonItemMapSetting_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            // 地图设定
            panelMain.Controls.SetChildIndex(dlgMapSetting, 0);
            this.Text = Properties.Settings.Default.SoftwareTitle + " — 地图设定";
        }

        private void barButtonItemScheduleSetting_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            // 日程设定
            panelMain.Controls.SetChildIndex(dlgSchedueSetting, 0);
            this.Text = Properties.Settings.Default.SoftwareTitle + " — 日程设定";
        }

        private void barButtonItemSystemSetting_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            // 系统设定
            SystemSetting dlg = new SystemSetting();
            if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                this.Text = Properties.Settings.Default.SoftwareTitle;
            }
        }

        private void barButtonItemUserSetting_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            // 用户设定
            UserSetting dlg = new UserSetting();
            dlg.ShowDialog();
        }

        private void barButtonItemLogoSetting_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            // 灯具图标设定
            LampSetting dlg = new LampSetting();
            dlg.ShowDialog();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            try
            {
                string queryString = string.Empty;
                string updatestring = string.Empty;
                queryString = "select SysId from SysArm where IsUse = '1' order by cast(SysId as int) asc";
                DataTable dt = CommonlyFunctions.getSchedue(queryString);

                if (dt.Rows.Count > 0)
                {
                    updatestring = "update SysOneLoad set LightingMin = (LightingMin + 1)  where IsUse ='1' and LoadStatus='1' and (";
                    if (dt.Rows.Count == 1)
                    {
                        updatestring += " ArmId =  '" + dt.Rows[0][0].ToString() + "')";

                    }
                    else
                    {
                        for (int i = 0; i < dt.Rows.Count - 1; i++)
                        {
                            updatestring += " ArmId =  '" + dt.Rows[i][0].ToString() + "' or ";

                        }
                        updatestring += " ArmId =  '" + dt.Rows[dt.Rows.Count - 1][0].ToString() + "')";

                    }
                    CommonlyFunctions.UpdateLightTime(updatestring);
                }
            }
            catch (Exception ex)
            {
                CommonlyFunctions.insertlog(Guid.NewGuid().ToString(), "-", "0", "更新OPIU点灯时间异常", DateTime.Now, "-", "-", ex.Message.ToString());
            }

        }

      

       
    }
}
