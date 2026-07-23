using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using System.Drawing.Imaging;
using System.IO;
using ComLib;
using System.Data.SqlClient;
using System.Drawing.Printing;
using System.Drawing.Drawing2D;

namespace AccAdm
{
    public partial class AccDealerCdDevimage : DevExpress.XtraEditors.XtraForm
    {
        public AccDealerCdDevimage()
        {
            InitializeComponent();
        }

        public delegate void SendDataHandler(DataRow row);
        public delegate void SendDataHandler2(string SLINO);
        public event SendDataHandler2 DataRowSendEvent2;
        public string _dealercd;
        public string _check;
        private string PROCEDURE_ID = "DP_IMAGE";
        public enum ModeGubun { Add, Modi }
        public ModeGubun _ModeGubun = ModeGubun.Add;

        
        private void EQ001F00_POP01_Load(object sender, EventArgs e)
        {
            Dictionary<string, string> dicParams = new Dictionary<string, string>();

            if (_check.Equals("1"))
            {
                dicParams.Clear();
                dicParams.Add("CMD", "PRINT1");
                dicParams.Add("DEALER_CD", _dealercd);

                DataTable dt = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);
                if (dt.Rows.Count > 0)
                {
                    if (string.IsNullOrEmpty(dt.Rows[0]["IMIMAGE"]?.ToString()))
                    {
                        BtnPrint.Enabled = false;
                    }
                }
                else
                {
                    BtnPrint.Enabled = false;
                }
                    
            }
            else if (_check.Equals("2"))
            {
                dicParams.Clear();
                dicParams.Add("CMD", "PRINT2");
                dicParams.Add("DEALER_CD", _dealercd);

                DataTable dt = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);
                if (dt.Rows.Count > 0)
                {
                    if (string.IsNullOrEmpty(dt.Rows[0]["IMIMAGE"]?.ToString()))
                    {
                        BtnPrint.Enabled = false;
                    }
                }
                else
                {
                    BtnPrint.Enabled = false;
                }

            }
            else if (_check.Equals("3"))
            {
                dicParams.Clear();
                dicParams.Add("CMD", "PRINT3");
                dicParams.Add("DEALER_CD", _dealercd);

                DataTable dt = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);
                if (dt.Rows.Count > 0)
                {
                    if (string.IsNullOrEmpty(dt.Rows[0]["IMIMAGE"]?.ToString()))
                    {
                        BtnPrint.Enabled = false;
                    }
                }
                else
                {
                    BtnPrint.Enabled = false;
                }
            }
            
            try
            {
                dicParams.Clear();
                dicParams.Add("CMD", "LIST3");
                dicParams.Add("DEALER_CD", _dealercd);
            
                DataTable dt = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);
            
                if (dt.Rows.Count > 0)
                {
                    PicInsa.Image = null;
                    if (_check.Equals("1"))
                    {
                        tx_MANAM.EditValue = dt.Rows[0]["IMTXT1"];
                    }
                    else if (_check.Equals("2"))
                    {
                        tx_MANAM.EditValue = dt.Rows[0]["IMTXT2"];
                    }
                    else if (_check.Equals("3"))
                    {
                        tx_MANAM.EditValue = dt.Rows[0]["IMTXT3"];
                    }
                    PicInsa.Image = null;
                    try
                    {
                        if (_check.Equals("1"))
                        {
                            byte[] ImgData = (byte[])dt.Rows[0]["IMIMAGE1"];
                            MemoryStream ms = new MemoryStream(ImgData);
                            Image returnImage = Image.FromStream(ms);
                            returnImage = FixImageOrientation(returnImage);
                            //PicInsa.Image = ResizeImage(returnImage, 300, 200);
                            PicInsa.Image = returnImage;
                        }
                        else if (_check.Equals("2"))
                        {
                            byte[] ImgData = (byte[])dt.Rows[0]["IMIMAGE2"];
                            MemoryStream ms = new MemoryStream(ImgData);
                            Image returnImage = Image.FromStream(ms);
                            returnImage = FixImageOrientation(returnImage);
                            //PicInsa.Image = ResizeImage(returnImage, 300, 200);
                            PicInsa.Image = returnImage;
                        }
                        else if (_check.Equals("3"))
                        {
                            byte[] ImgData = (byte[])dt.Rows[0]["IMIMAGE3"];
                            MemoryStream ms = new MemoryStream(ImgData);
                            Image returnImage = Image.FromStream(ms);
                            returnImage = FixImageOrientation(returnImage);
                            //PicInsa.Image = ResizeImage(returnImage, 300, 200);
                            PicInsa.Image = returnImage;
                        }
                    }
                    catch
                    {
            
                    }
                }
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.ToString(), "기본사항 조회 오류");
            }
        }
        private Image ResizeImage(Image image, int width, int height)
        {
            // 새로운 이미지 생성
            Bitmap resizedImage = new Bitmap(width, height);

            // 그래픽 개체 생성
            using (Graphics graphics = Graphics.FromImage(resizedImage))
            {
                graphics.CompositingQuality = CompositingQuality.HighQuality;
                graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
                graphics.SmoothingMode = SmoothingMode.HighQuality;

                // 이미지를 새로운 크기로 그립니다.
                graphics.DrawImage(image, 0, 0, width, height);
            }

            return resizedImage;
        }
        #region [Function]
        //Edit초기화

        private void SaveInsaF()
        {
            StringBuilder strSql = new StringBuilder();
            string sMANAM = tx_MANAM.EditValue?.ToString();
            
            if (PicInsa.Image != null)
            {
                if (string.IsNullOrEmpty(sMANAM))
                {
                    XtraMessageBox.Show("이미지 정보를 입력하세요.", "");
                    tx_MANAM.SelectAll();
                    tx_MANAM.Focus();
                    return;
                }
            }

            if (!string.IsNullOrEmpty(sMANAM))
            {
                if (PicInsa.Image == null)
                {
                    XtraMessageBox.Show("이미지가 없습니다.", "");
                    return;
                }
            }
            
            byte[] byteImg = null;
            if (PicInsa.Image != null)
            {
                Image img = PicInsa.Image;
                MemoryStream ms = new MemoryStream();
                img.Save(ms, img.RawFormat);
                byteImg = ms.ToArray();
            }

            if (_check.Equals("1"))
            {
                Dictionary<string, Object> dicParams = new Dictionary<string, Object>();

                dicParams.Clear();
                dicParams.Add("CMD", "SAVE1");
                dicParams.Add("DEALER_CD", _dealercd);
                dicParams.Add("IMTXT1", sMANAM);
                dicParams.Add("IMIMAGE1", byteImg);

                DataTable dt = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);
                string sResult = dt.Rows[0]["RESULT"]?.ToString();


                if (sResult.Equals("-1"))
                {
                    XtraMessageBox.SmartTextWrap = true;
                    XtraMessageBox.Show(dt.Rows[0]["MSG"]?.ToString(), "등록오류");
                }
                else
                {
                    if (this.Owner.Name.Equals("AccDealerCdDev"))
                    {
                        AccDealerCdDev pFrm = (AccDealerCdDev)this.Owner;
                        pFrm._dealercd = dt.Rows[0]["DEALER_CD"]?.ToString();
                        DialogResult = DialogResult.OK;
                    }
                    else
                    {
                        //DataRowSendEvent2(dt.Rows[0]["DEALER_CD"]?.ToString());
                        Dispose();
                    }
                }
            }
            else if (_check.Equals("2"))
            {
                Dictionary<string, Object> dicParams = new Dictionary<string, Object>();

                dicParams.Clear();
                dicParams.Add("CMD", "SAVE2");
                dicParams.Add("DEALER_CD", _dealercd);
                dicParams.Add("IMTXT2", sMANAM);
                dicParams.Add("IMIMAGE2", byteImg);

                DataTable dt = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);
                string sResult = dt.Rows[0]["RESULT"]?.ToString();


                if (sResult.Equals("-1"))
                {
                    XtraMessageBox.SmartTextWrap = true;
                    XtraMessageBox.Show(dt.Rows[0]["MSG"]?.ToString(), "등록오류");
                }
                else
                {
                    
                    if (this.Owner.Name.Equals("AccDealerCdDev"))
                    {
                        AccDealerCdDev pFrm = (AccDealerCdDev)this.Owner;
                        pFrm._dealercd = dt.Rows[0]["DEALER_CD"]?.ToString();
                        DialogResult = DialogResult.OK;
                    }
                    else
                    {
                        //DataRowSendEvent2(dt.Rows[0]["DEALER_CD"]?.ToString());
                        Dispose();
                    }
                }
            }
            else if (_check.Equals("3"))
            {
                Dictionary<string, Object> dicParams = new Dictionary<string, Object>();

                dicParams.Clear();
                dicParams.Add("CMD", "SAVE3");
                dicParams.Add("DEALER_CD", _dealercd);
                dicParams.Add("IMTXT3", sMANAM);
                dicParams.Add("IMIMAGE3", byteImg);

                DataTable dt = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);
                string sResult = dt.Rows[0]["RESULT"]?.ToString();


                if (sResult.Equals("-1"))
                {
                    XtraMessageBox.SmartTextWrap = true;
                    XtraMessageBox.Show(dt.Rows[0]["MSG"]?.ToString(), "등록오류");
                }
                else
                {
                    if (this.Owner.Name.Equals("AccDealerCdDev"))
                    {
                        AccDealerCdDev pFrm = (AccDealerCdDev)this.Owner;
                        pFrm._dealercd = dt.Rows[0]["DEALER_CD"]?.ToString();
                        DialogResult = DialogResult.OK;
                    }
                    else
                    {
                        //DataRowSendEvent2(dt.Rows[0]["DEALER_CD"]?.ToString());
                        Dispose();
                    }
                }
            }
        }


        #endregion

        #region [버튼이벤트]
        private void BtnSave_Click(object sender, EventArgs e)
        {
            SaveInsaF();
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
        }

        private void BtnAddPic_Click(object sender, EventArgs e)
        {

        }
        #endregion

        private void EQ001F00_POP01_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
                BtnClose.PerformClick();
            else if (e.KeyCode == Keys.F3)
                BtnSave.PerformClick();
        }
        
        private void RadiPaitg_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                BtnSave.Focus();
            }
        }
        

        private void layoutControlGroup1_CustomButtonClick(object sender, DevExpress.XtraBars.Docking2010.BaseButtonEventArgs e)
        {
            if (e.Button.Properties.Tag.Equals("T_ADD"))
            {
                string sFile = null;
                OpenFileDialog dialog = new OpenFileDialog();
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    sFile = dialog.FileName;
                    //PicInsa.Image = Image.FromFile(sFile);
                    Image loadedImage = Image.FromFile(sFile);
                    loadedImage = FixImageOrientation(loadedImage);
                    PicInsa.Image = loadedImage;
                }
                else
                {
                    return;
                }
            }
            else if (e.Button.Properties.Tag.Equals("T_DEL"))
            {
                PicInsa.EditValue = null;
                tx_MANAM.EditValue = null;
            }
        }

        private void BtnPrint_Click(object sender, EventArgs e)
        {
            if (_check.Equals("1"))
            {
                Dictionary<string, string> dicParams = new Dictionary<string, string>();

                dicParams.Clear();
                dicParams.Add("CMD", "PRINT1");
                dicParams.Add("DEALER_CD", _dealercd);

                DataTable dt = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);
                

                ReportViewer sr4 = new ReportViewer(dt, "RptDealerImage");
                sr4.StartPosition = FormStartPosition.WindowsDefaultLocation;
                sr4.WindowState = FormWindowState.Maximized;
                sr4.ShowDialog();
            }
            else if (_check.Equals("2"))
            {
                Dictionary<string, string> dicParams = new Dictionary<string, string>();

                dicParams.Clear();
                dicParams.Add("CMD", "PRINT2");
                dicParams.Add("DEALER_CD", _dealercd);

                DataTable dt = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);
               
                ReportViewer sr4 = new ReportViewer(dt, "RptDealerImage");
                sr4.StartPosition = FormStartPosition.WindowsDefaultLocation;
                sr4.WindowState = FormWindowState.Maximized;
                sr4.ShowDialog();
            }
            else if (_check.Equals("3"))
            {
                Dictionary<string, string> dicParams = new Dictionary<string, string>();

                dicParams.Clear();
                dicParams.Add("CMD", "PRINT3");
                dicParams.Add("DEALER_CD", _dealercd);

                DataTable dt = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);
               
                ReportViewer sr4 = new ReportViewer(dt, "RptDealerImage");
                sr4.StartPosition = FormStartPosition.WindowsDefaultLocation;
                sr4.WindowState = FormWindowState.Maximized;
                sr4.ShowDialog();
            }
        }
        private Image FixImageOrientation(Image img)
        {
            const int ExifOrientationId = 0x0112; // EXIF 'Orientation' 태그 ID
            if (img.PropertyIdList.Contains(ExifOrientationId))
            {
                var prop = img.GetPropertyItem(ExifOrientationId);
                int val = BitConverter.ToUInt16(prop.Value, 0);
                RotateFlipType rotateFlip = RotateFlipType.RotateNoneFlipNone;

                switch (val)
                {
                    case 2:
                        rotateFlip = RotateFlipType.RotateNoneFlipX;
                        break;
                    case 3:
                        rotateFlip = RotateFlipType.Rotate180FlipNone;
                        break;
                    case 4:
                        rotateFlip = RotateFlipType.Rotate180FlipX;
                        break;
                    case 5:
                        rotateFlip = RotateFlipType.Rotate90FlipX;
                        break;
                    case 6:
                        rotateFlip = RotateFlipType.Rotate90FlipNone;
                        break;
                    case 7:
                        rotateFlip = RotateFlipType.Rotate270FlipX;
                        break;
                    case 8:
                        rotateFlip = RotateFlipType.Rotate270FlipNone;
                        break;
                }

                if (rotateFlip != RotateFlipType.RotateNoneFlipNone)
                    img.RotateFlip(rotateFlip);

                // EXIF 정보 초기화 (중복 회전 방지)
                img.RemovePropertyItem(ExifOrientationId);
            }

            return img;
        }

    }
}