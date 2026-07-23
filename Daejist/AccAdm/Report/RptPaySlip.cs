using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using DevExpress.XtraReports.UI;
using System.Data;
using System.Collections.Generic;

namespace AccAdm
{
    public partial class RptPaySlip : DevExpress.XtraReports.UI.XtraReport
    {
        public RptPaySlip()
        {
            InitializeComponent();
        }

        public RptPaySlip(DataTable dt)
        {
            InitializeComponent();

            string sBASYM = dt.Rows[0]["BASYM"]?.ToString();

            TxtBASYM.Text = sBASYM.Substring(0, 4) + "년 " + sBASYM.Substring(5, 2) + "월 급여명세서";
            TxtPAYDT.Text = "[ 지급일 : " + dt.Rows[0]["PAYDT"]?.ToString() + " ]";
            TxtENTRANCE_YMD.Text = dt.Rows[0]["ENTRANCE_YMD"]?.ToString();
            TxtEMP_NM.Text = dt.Rows[0]["EMP_NM"]?.ToString();
            TxtDEPT_NM.Text = dt.Rows[0]["DEPT_NM"]?.ToString();
            TxtJIKWI.Text = dt.Rows[0]["JIKWI"]?.ToString();
            TxtPWKTM.Text = dt.Rows[0]["PWKTM"]?.ToString();
            TxtPYJTM.Text = dt.Rows[0]["PYJTM"]?.ToString();
            TxtJANUP.Text = dt.Rows[0]["JANUP"]?.ToString();
            //TxtTKGUN.Text = dt.Rows[0]["TKGUN"]?.ToString();

            if (dt.Rows[0]["GTGJTM"].Equals("0.0"))
                TxtGTGJTM.Text = null;
            else
                TxtGTGJTM.Text = dt.Rows[0]["GTGJTM"]?.ToString();

            TxtTOTWKTM.Text = dt.Rows[0]["TOTWKTM"]?.ToString();

            string sPSIGUB1= dt.Rows[0]["PSIGUB1"]?.ToString();
            string sPGJSD1 = dt.Rows[0]["PGJSD1"]?.ToString();    //기본급
            string sPGJSD2 = dt.Rows[0]["PGJSD2"]?.ToString();    //고정연장수당
            string sPGJSD3 = dt.Rows[0]["PGJSD3"]?.ToString();
            string sPGJSD4 = dt.Rows[0]["PGJSD4"]?.ToString();
            string sPGJSD5 = dt.Rows[0]["PGJSD5"]?.ToString();
            string sPGJSD6 = dt.Rows[0]["PGJSD6"]?.ToString();
            string sPGJSD7 = dt.Rows[0]["PGJSD7"]?.ToString();
            string sPCHSD1 = dt.Rows[0]["PCHSD1"]?.ToString();    //잔업수당
            string sPCHSD2 = dt.Rows[0]["PCHSD2"]?.ToString();
            string sPCHSD3 = dt.Rows[0]["PCHSD3"]?.ToString();
            string sPCHSD4 = dt.Rows[0]["PCHSD4"]?.ToString();
            string sPCHSD5 = dt.Rows[0]["PCHSD5"]?.ToString();
            string sPCHSD6 = dt.Rows[0]["PCHSD6"]?.ToString();    //인센티브
            string sPCHSD7 = dt.Rows[0]["PCHSD7"]?.ToString();    //통신비
            string sPCHSD8 = dt.Rows[0]["PCHSD8"]?.ToString();
            string sPGROSS = dt.Rows[0]["PGROSS"]?.ToString();    //급여합계
            string sPGJGJ1 = dt.Rows[0]["PGJGJ1"]?.ToString();    //소득세 
            string sPGJGJ2 = dt.Rows[0]["PGJGJ2"]?.ToString();    //주민세
            string sPGJGJ3 = dt.Rows[0]["PGJGJ3"]?.ToString();    //국민연금
            string sPGJGJ4 = dt.Rows[0]["PGJGJ4"]?.ToString();    //건강보험
            string sPGJGJ5 = dt.Rows[0]["PGJGJ5"]?.ToString();    //고정회비
            string sPGJGJ6 = dt.Rows[0]["PGJGJ6"]?.ToString();
            string sPGJGJ7 = dt.Rows[0]["PGJGJ7"]?.ToString();
            string sPCHGJ1 = dt.Rows[0]["PCHGJ1"]?.ToString();    //고용보험
            string sPCHGJ2 = dt.Rows[0]["PCHGJ2"]?.ToString();    //건강보험연말정산
            string sPCHGJ3 = dt.Rows[0]["PCHGJ3"]?.ToString();    //가불금
            string sPCHGJ4 = dt.Rows[0]["PCHGJ4"]?.ToString();    //보수총액
            string sPCHGJ5 = dt.Rows[0]["PCHGJ5"]?.ToString();    //기존건강보험료
            string sPCHGJ6 = dt.Rows[0]["PCHGJ6"]?.ToString();    
            string sPCHGJ7 = dt.Rows[0]["PCHGJ7"]?.ToString();
            string sPTOTGO = dt.Rows[0]["PTOTGO"]?.ToString();    //공제합계
            string sPCHAIN = dt.Rows[0]["PCHAIN"]?.ToString();    //실지급액
                                                                  
            double dPSIGUB1 = 0;                                  
            double dPGJSD1 = 0;
            double dPGJSD2 = 0;
            double dPGJSD3 = 0;
            double dPGJSD4 = 0;
            double dPGJSD5 = 0;
            double dPGJSD6 = 0;
            double dPGJSD7 = 0;
            double dPCHSD1 = 0;
            double dPCHSD2 = 0;
            double dPCHSD3 = 0;
            double dPCHSD4 = 0;
            double dPCHSD5 = 0;
            double dPCHSD6 = 0;
            double dPCHSD7 = 0;
            double dPCHSD8 = 0;
            double dPGROSS = 0;
            double dPGJGJ1 = 0;
            double dPGJGJ2 = 0;
            double dPGJGJ3 = 0;
            double dPGJGJ4 = 0;
            double dPGJGJ5 = 0;
            double dPGJGJ6 = 0;
            double dPGJGJ7 = 0;
            double dPCHGJ1 = 0;
            double dPCHGJ2 = 0;
            double dPCHGJ3 = 0;
            double dPCHGJ4 = 0;
            double dPCHGJ5 = 0;
            double dPCHGJ6 = 0;
            double dPCHGJ7 = 0;
            double dPTOTGO = 0;
            double dPCHAIN = 0;

            double.TryParse(sPSIGUB1, out dPSIGUB1);
            double.TryParse(sPGJSD1 , out dPGJSD1 );
            double.TryParse(sPGJSD2 , out dPGJSD2 );
            double.TryParse(sPGJSD3 , out dPGJSD3 );
            double.TryParse(sPGJSD4 , out dPGJSD4 );
            double.TryParse(sPGJSD5 , out dPGJSD5 );
            double.TryParse(sPGJSD6 , out dPGJSD6 );
            double.TryParse(sPGJSD7 , out dPGJSD7 );
            double.TryParse(sPCHSD1 , out dPCHSD1 );
            double.TryParse(sPCHSD2 , out dPCHSD2 );
            double.TryParse(sPCHSD3 , out dPCHSD3 );
            double.TryParse(sPCHSD4 , out dPCHSD4 );
            double.TryParse(sPCHSD5 , out dPCHSD5 );
            double.TryParse(sPCHSD6 , out dPCHSD6 );
            double.TryParse(sPCHSD7 , out dPCHSD7 );
            double.TryParse(sPCHSD8 , out dPCHSD8 );
            double.TryParse(sPGROSS , out dPGROSS );
            double.TryParse(sPGJGJ1 , out dPGJGJ1 );
            double.TryParse(sPGJGJ2 , out dPGJGJ2 );
            double.TryParse(sPGJGJ3 , out dPGJGJ3 );
            double.TryParse(sPGJGJ4 , out dPGJGJ4 );
            double.TryParse(sPGJGJ5 , out dPGJGJ5 );
            double.TryParse(sPGJGJ6 , out dPGJGJ6 );
            double.TryParse(sPGJGJ7 , out dPGJGJ7 );
            double.TryParse(sPCHGJ1 , out dPCHGJ1 );
            double.TryParse(sPCHGJ2 , out dPCHGJ2 );
            double.TryParse(sPCHGJ3 , out dPCHGJ3 );
            double.TryParse(sPCHGJ4 , out dPCHGJ4 );
            double.TryParse(sPCHGJ5 , out dPCHGJ5 );
            double.TryParse(sPCHGJ6 , out dPCHGJ6 );
            double.TryParse(sPCHGJ7 , out dPCHGJ7 );
            double.TryParse(sPTOTGO , out dPTOTGO );
            double.TryParse(sPCHAIN, out dPCHAIN);

            TxtPSIGUB1.Text = dPSIGUB1.ToString("n0");
            TxtPGJSD1.Text = dPGJSD1.ToString("n0");       //기본급
            TxtPGJSD2.Text = dPGJSD2.ToString("n0");       //고정연장수당
            //TxtPGJSD3.Text = dPGJSD3;       //
            //TxtPGJSD4.Text = dPGJSD4;       //
            //TxtPGJSD5.Text = dPGJSD5;       //
            //TxtPGJSD6.Text = dPGJSD6;       //
            //TxtPGJSD7.Text = dPGJSD7;       //
            TxtPCHSD1.Text = dPCHSD1.ToString("n0");       //잔업수당
            //TxtPCHSD2.Text = dPCHSD2;       //
            //TxtPCHSD3.Text = dPCHSD3;       //
            //TxtPCHSD4.Text = dPCHSD4;       //
            //TxtPCHSD5.Text = dPCHSD5;       //
            TxtPCHSD6.Text = dPCHSD6.ToString("n0");       //인센티브
            TxtPCHSD7.Text = dPCHSD7.ToString("n0");       //통신비
            //TxtPCHSD8.Text = dPCHSD8;       //
            TxtPGROSS.Text = dPGROSS.ToString("n0");       //급여합계
            TxtPGROSS2.Text = dPGROSS.ToString("n0");       //급여합계
            TxtPGJGJ1.Text = dPGJGJ1.ToString("n0");       //소득세
            TxtPGJGJ2.Text = dPGJGJ2.ToString("n0");       //주민세
            TxtPGJGJ3.Text = dPGJGJ3.ToString("n0");       //국민연금
            TxtPGJGJ4.Text = dPGJGJ4.ToString("n0");       //건강보험
            //TxtPGJGJ5.Text = dPGJGJ5;       //
            //TxtPGJGJ6.Text = dPGJGJ6;       //
            //TxtPGJGJ7.Text = dPGJGJ7;       //
            TxtPCHGJ1.Text = dPCHGJ1.ToString("n0");       //고용보험
            TxtPCHGJ2.Text = dPCHGJ2.ToString("n0");       //건강보험연말정산
            TxtPCHGJ3.Text = (dPCHGJ3+ dPGJGJ5).ToString("n0");       //가불금= 고정회비+가불금
            TxtPCHGJ4.Text = dPCHGJ4.ToString("n0");       //보수총액
            TxtPCHGJ5.Text = dPCHGJ5.ToString("n0");       //기존건강보험료
            //TxtPCHGJ6.Text = dPCHGJ6;       //
            //TxtPCHGJ7.Text = dPCHGJ7;       //
            TxtPTOTGO.Text = dPTOTGO.ToString("n0");       //공제합계
            TxtPCHAIN.Text = dPCHAIN.ToString("n0");       //실지급액
            TxtPTOTGO2.Text = dPTOTGO.ToString("n0");       //공제합계
        }

    }
}
