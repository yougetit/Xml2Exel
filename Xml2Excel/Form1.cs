using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using System.IO;
using System.Xml;

namespace Xml2Excel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //openFileDialog1.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory); //預設開啟資料夾
            openFileDialog1.FileName = string.Empty; //預設檔名
            openFileDialog1.Filter = "XML、EXCEL File(*.xml; *.xls)|*.xml; *.xls"; //過濾檔案類型
            openFileDialog1.RestoreDirectory = true;
            if (openFileDialog1.ShowDialog() == DialogResult.Cancel)
            {
                textBox1.Text = string.Empty;
            }
            else
            {
                textBox1.Text = openFileDialog1.FileName;
            }
        }

        int n, rowNumber;
        private string year, caseNo, serNo, pID, orderList, orderCode, amount, cashPoint, doID;
        private void button2_Click(object sender, EventArgs e)
        {
            //建立新excel檔

            IWorkbook workbook = new HSSFWorkbook();
            ISheet sheet = workbook.CreateSheet("new sheet");

            sheet.CreateRow(0).CreateCell(0).SetCellValue("身分證");
            sheet.GetRow(0).CreateCell(1).SetCellValue("案件分類");
            sheet.GetRow(0).CreateCell(2).SetCellValue("流水號");
            sheet.GetRow(0).CreateCell(3).SetCellValue("醫令序");
            sheet.GetRow(0).CreateCell(4).SetCellValue("醫令代碼");
            sheet.GetRow(0).CreateCell(5).SetCellValue("申報數量");
            sheet.GetRow(0).CreateCell(6).SetCellValue("申報金額");
            sheet.GetRow(0).CreateCell(7).SetCellValue("醫事人員");
            sheet.GetRow(0).CreateCell(8).SetCellValue("門住申報");

            n = 0;
            rowNumber = 0;

            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(textBox1.Text);

            XmlNode tdataNode = xmlDoc.SelectSingleNode("inpatient/tdata");
            XmlNodeList tdataNodeChildNodeList = tdataNode.ChildNodes;

            foreach (XmlNode tNode in tdataNodeChildNodeList)
            {
                if (tNode.Name == "t3")
                {
                    year = tNode.InnerText;
                }
            }

            XmlNodeList nodeLists = xmlDoc.SelectNodes("inpatient/ddata");
            foreach (XmlNode node in nodeLists)
            {

                XmlNode childNodeHead = node.SelectSingleNode("dhead");
                XmlNodeList dheadNodeList = childNodeHead.ChildNodes;
                foreach (XmlNode child in dheadNodeList)
                {
                    if (child.Name == "d2")//流水號
                        serNo = child.InnerText;
                    if (child.Name == "d1") //案件分類
                        caseNo = child.InnerText;
                }

                XmlNode childNodeBody = node.SelectSingleNode("dbody");
                XmlNodeList dbodyNodeList = childNodeBody.ChildNodes;

                foreach (XmlNode dbodyChild in dbodyNodeList)
                {
                    if (dbodyChild.Name == "d3") //身份證
                    {
                        pID = dbodyChild.InnerText;
                    }

                    if (dbodyChild.Name == "pdata")
                    {
                        XmlNodeList pdataList = dbodyChild.ChildNodes;
                        foreach (XmlNode childPdata in pdataList)
                        {

                            string nodePdataName = childPdata.Name;
                            string nodePdataValue = childPdata.InnerText;
                            //MessageBox.Show(nodePdataName + "_" + nodePdataValue);
                            if (nodePdataName == "p1")
                                orderList = nodePdataValue.Trim();
                            if (nodePdataName == "p3")
                                orderCode = nodePdataValue.Trim();
                            if (nodePdataName == "p16")
                                amount = nodePdataValue.Trim();
                            if (nodePdataName == "p18")
                                cashPoint = nodePdataValue.Trim();
                            if (nodePdataName == "p20")
                                doID = nodePdataValue.Trim();
                        }
                        string subOrderCode;
                        if (orderCode != null)
                        {
                            subOrderCode = orderCode.Substring(0, 2);

                            if (subOrderCode == "06" || subOrderCode == "07" || subOrderCode == "08" || subOrderCode == "09" || subOrderCode == "10" || subOrderCode == "12" || subOrderCode == "13" || subOrderCode == "14" || subOrderCode == "18" || subOrderCode == "85" || subOrderCode == "21" || subOrderCode == "22")
                            {
                                n++;
                                rowNumber++;
                                sheet.CreateRow(rowNumber).CreateCell(0).SetCellValue(pID);
                                sheet.GetRow(rowNumber).CreateCell(1).SetCellValue(caseNo);
                                sheet.GetRow(rowNumber).CreateCell(2).SetCellValue(serNo);
                                sheet.GetRow(rowNumber).CreateCell(3).SetCellValue(orderList);
                                sheet.GetRow(rowNumber).CreateCell(4).SetCellValue(orderCode);
                                sheet.GetRow(rowNumber).CreateCell(5).SetCellValue(amount);
                                sheet.GetRow(rowNumber).CreateCell(6).SetCellValue(cashPoint);
                                sheet.GetRow(rowNumber).CreateCell(7).SetCellValue(doID);
                                sheet.GetRow(rowNumber).CreateCell(8).SetCellValue("住");
                            }
                        }
                    }

                    //MessageBox.Show(pID + "_" + caseNo + "_" + serNo + "_" + orderList + "_" + orderCode + "_" + amount + "_" + cashPoint + "_" + doID);

                }
            }

            string filename = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + @"\檢驗報表\" + year + "住院明細" + DateTime.Now.ToString("yyyy-M-d" + "HH-mm-ss") + ".xls";
            //MessageBox.Show(filename);
            FileStream file = new FileStream(filename, FileMode.Create, FileAccess.Write);
            workbook.Write(file);
            file.Close();
            MessageBox.Show("done!共" + n);


        }

        private void button3_Click(object sender, EventArgs e)
        {
            //建立新excel檔

            IWorkbook workbook = new HSSFWorkbook();
            ISheet sheet = workbook.CreateSheet("new sheet");

            sheet.CreateRow(0).CreateCell(0).SetCellValue("身分證");
            sheet.GetRow(0).CreateCell(1).SetCellValue("案件分類");
            sheet.GetRow(0).CreateCell(2).SetCellValue("流水號");
            sheet.GetRow(0).CreateCell(3).SetCellValue("醫令序");
            sheet.GetRow(0).CreateCell(4).SetCellValue("醫令代碼");
            sheet.GetRow(0).CreateCell(5).SetCellValue("申報數量");
            sheet.GetRow(0).CreateCell(6).SetCellValue("申報金額");
            sheet.GetRow(0).CreateCell(7).SetCellValue("醫事人員");
            sheet.GetRow(0).CreateCell(8).SetCellValue("門住申報");

            n = 0;
            rowNumber = 0;

            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(textBox1.Text);

            XmlNode tdataNode = xmlDoc.SelectSingleNode("outpatient/tdata");
            XmlNodeList tdataNodeChildNodeList = tdataNode.ChildNodes;

            foreach (XmlNode tNode in tdataNodeChildNodeList)
            {
                if (tNode.Name == "t3")
                {
                    year = tNode.InnerText;
                }
            }

            XmlNodeList nodeLists = xmlDoc.SelectNodes("outpatient/ddata");
            foreach (XmlNode ddataNode in nodeLists)
            {
                XmlNode childNodeHead = ddataNode.SelectSingleNode("dhead");
                XmlNodeList dheadNodeList = childNodeHead.ChildNodes;
                foreach (XmlNode child in dheadNodeList)
                {
                    if (child.Name == "d2")//流水號
                        serNo = child.InnerText;
                    if (child.Name == "d1") //案件分類
                        caseNo = child.InnerText;
                }

                XmlNode childNodeBody = ddataNode.SelectSingleNode("dbody");
                XmlNodeList dbodyNodeList = childNodeBody.ChildNodes;

                foreach (XmlNode dbodyChild in dbodyNodeList)
                {
                    if (dbodyChild.Name == "d3") //身份證
                    {
                        pID = dbodyChild.InnerText;
                    }

                    if (dbodyChild.Name == "pdata")
                    {
                        XmlNodeList pdataList = dbodyChild.ChildNodes;
                        foreach (XmlNode childPdata in pdataList)
                        {

                            string nodePdataName = childPdata.Name;
                            string nodePdataValue = childPdata.InnerText;
                            //MessageBox.Show(nodePdataName + "_" + nodePdataValue);
                            if (nodePdataName == "p13")
                                orderList = nodePdataValue.Trim();
                            if (nodePdataName == "p4")
                                orderCode = nodePdataValue.Trim();
                            if (nodePdataName == "p10")
                                amount = nodePdataValue.Trim();
                            if (nodePdataName == "p12")
                                cashPoint = nodePdataValue.Trim();
                            if (nodePdataName == "p16")
                                doID = nodePdataValue.Trim();
                        }
                        string subOrderCode;
                        if (orderCode != null)
                        {
                            subOrderCode = orderCode.Substring(0, 2);

                            if (subOrderCode == "06" || subOrderCode == "07" || subOrderCode == "08" || subOrderCode == "09" || subOrderCode == "10" || subOrderCode == "12" || subOrderCode == "13" || subOrderCode == "14" || subOrderCode == "18" || subOrderCode == "85" || subOrderCode == "21" || subOrderCode == "22")
                            {
                                n++;
                                rowNumber++;
                                sheet.CreateRow(rowNumber).CreateCell(0).SetCellValue(pID);
                                sheet.GetRow(rowNumber).CreateCell(1).SetCellValue(caseNo);
                                sheet.GetRow(rowNumber).CreateCell(2).SetCellValue(serNo);
                                sheet.GetRow(rowNumber).CreateCell(3).SetCellValue(orderList);
                                sheet.GetRow(rowNumber).CreateCell(4).SetCellValue(orderCode);
                                sheet.GetRow(rowNumber).CreateCell(5).SetCellValue(amount);
                                sheet.GetRow(rowNumber).CreateCell(6).SetCellValue(cashPoint);
                                sheet.GetRow(rowNumber).CreateCell(7).SetCellValue(doID);
                                sheet.GetRow(rowNumber).CreateCell(8).SetCellValue("門");
                            }
                        }
                    }

                    //MessageBox.Show(pID + "_" + caseNo + "_" + serNo + "_" + orderList + "_" + orderCode + "_" + amount + "_" + cashPoint + "_" + doID);

                }
            }

            string filename = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + @"\檢驗報表\" + year + "門診明細" + DateTime.Now.ToString("yyyy-M-d" + "HH-mm-ss") + ".xls";
            FileStream file = new FileStream(filename, FileMode.Create, FileAccess.Write);
            workbook.Write(file);
            file.Close();
            MessageBox.Show("done!共" + n);

        }

        private string name, indate, outdate, icd1, icd2, icd3, icd4, icd5, icd6, icd7, icd8, icd9, icd10, dr, bedno, stime, etime;

        private void button4_Click(object sender, EventArgs e)
        {
            //建立新excel檔



            IWorkbook workbook = new HSSFWorkbook();
            ISheet sheet = workbook.CreateSheet("new sheet");

            sheet.CreateRow(0).CreateCell(0).SetCellValue("身分證"); //d3
            sheet.GetRow(0).CreateCell(1).SetCellValue("案件分類"); //d1
            sheet.GetRow(0).CreateCell(2).SetCellValue("流水號");  //d2
            sheet.GetRow(0).CreateCell(3).SetCellValue("姓名");   //d103
            sheet.GetRow(0).CreateCell(4).SetCellValue("入院日期"); //d10
            sheet.GetRow(0).CreateCell(5).SetCellValue("出院日期"); //d11
            sheet.GetRow(0).CreateCell(6).SetCellValue("主診斷"); //d25
            sheet.GetRow(0).CreateCell(7).SetCellValue("副診斷1"); //d26
            sheet.GetRow(0).CreateCell(8).SetCellValue("副診斷2"); //d27
            sheet.GetRow(0).CreateCell(9).SetCellValue("副診斷3"); //d28
            sheet.GetRow(0).CreateCell(10).SetCellValue("副診斷4"); //d29
            sheet.GetRow(0).CreateCell(11).SetCellValue("副診斷5"); //d30
            sheet.GetRow(0).CreateCell(12).SetCellValue("副診斷6"); //d31
            sheet.GetRow(0).CreateCell(13).SetCellValue("副診斷7"); //d32
            sheet.GetRow(0).CreateCell(14).SetCellValue("副診斷8"); //d33
            sheet.GetRow(0).CreateCell(15).SetCellValue("副診斷9"); //d34

            n = 0;
            rowNumber = 0;

            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(textBox1.Text);

            XmlNode tdataNode = xmlDoc.SelectSingleNode("inpatient/tdata");
            XmlNodeList tdataNodeChildNodeList = tdataNode.ChildNodes;

            foreach (XmlNode tNode in tdataNodeChildNodeList)
            {
                if (tNode.Name == "t3")
                {
                    year = tNode.InnerText;
                }
            }

            XmlNodeList nodeLists = xmlDoc.SelectNodes("inpatient/ddata");
            foreach (XmlNode ddataNode in nodeLists)
            {
                XmlNode childNodeHead = ddataNode.SelectSingleNode("dhead");
                XmlNodeList dheadNodeList = childNodeHead.ChildNodes;
                foreach (XmlNode child in dheadNodeList)
                {
                    if (child.Name == "d2")//流水號
                        serNo = child.InnerText;
                    if (child.Name == "d1") //案件分類
                        caseNo = child.InnerText;
                }

                XmlNode childNodeBody = ddataNode.SelectSingleNode("dbody");
                XmlNodeList dbodyNodeList = childNodeBody.ChildNodes;

                foreach (XmlNode dbodyChild in dbodyNodeList)
                {
                    if (dbodyChild.Name == "d3") //身份證
                    {
                        pID = dbodyChild.InnerText;
                    }

                    if (dbodyChild.Name == "d103") //姓名
                    {
                        name = dbodyChild.InnerText;
                    }

                    if (dbodyChild.Name == "d10") //入院日期
                    {
                        indate = dbodyChild.InnerText;
                    }

                    if (dbodyChild.Name == "d11") //出院日期
                    {
                        outdate = dbodyChild.InnerText;
                    }

                    if (dbodyChild.Name == "d25") //主診斷
                    {
                        icd1 = dbodyChild.InnerText;
                    }

                    if (dbodyChild.Name == "d26") //副診斷1
                    {
                        icd2 = dbodyChild.InnerText;
                    }

                    if (dbodyChild.Name == "d27") //副診斷2
                    {
                        icd3 = dbodyChild.InnerText;
                    }

                    if (dbodyChild.Name == "d28") //副診斷3
                    {
                        icd4 = dbodyChild.InnerText;
                    }

                    if (dbodyChild.Name == "d29") //副診斷4
                    {
                        icd5 = dbodyChild.InnerText;
                    }

                    if (dbodyChild.Name == "d30") //副診斷5
                    {
                        icd6 = dbodyChild.InnerText;
                    }

                    if (dbodyChild.Name == "d31") //副診斷6
                    {
                        icd7 = dbodyChild.InnerText;
                    }

                    if (dbodyChild.Name == "d32") //副診斷7
                    {
                        icd8 = dbodyChild.InnerText;
                    }

                    if (dbodyChild.Name == "d33") //副診斷8
                    {
                        icd9 = dbodyChild.InnerText;
                    }

                    if (dbodyChild.Name == "d34") //副診斷9
                    {
                        icd10 = dbodyChild.InnerText;
                    }
                }

                n++;
                rowNumber++;
                sheet.CreateRow(rowNumber).CreateCell(0).SetCellValue(pID);
                sheet.GetRow(rowNumber).CreateCell(1).SetCellValue(caseNo);
                sheet.GetRow(rowNumber).CreateCell(2).SetCellValue(serNo);
                sheet.GetRow(rowNumber).CreateCell(3).SetCellValue(name);
                sheet.GetRow(rowNumber).CreateCell(4).SetCellValue(indate);
                sheet.GetRow(rowNumber).CreateCell(5).SetCellValue(outdate);
                sheet.GetRow(rowNumber).CreateCell(6).SetCellValue(icd1);
                sheet.GetRow(rowNumber).CreateCell(7).SetCellValue(icd2);
                sheet.GetRow(rowNumber).CreateCell(8).SetCellValue(icd3);
                sheet.GetRow(rowNumber).CreateCell(9).SetCellValue(icd4);
                sheet.GetRow(rowNumber).CreateCell(10).SetCellValue(icd5);
                sheet.GetRow(rowNumber).CreateCell(11).SetCellValue(icd6);
                sheet.GetRow(rowNumber).CreateCell(12).SetCellValue(icd7);
                sheet.GetRow(rowNumber).CreateCell(13).SetCellValue(icd8);
                sheet.GetRow(rowNumber).CreateCell(14).SetCellValue(icd9);
                sheet.GetRow(rowNumber).CreateCell(15).SetCellValue(icd10);
            }

            string filename = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + @"\檢驗報表\" + year + "住院診斷" + DateTime.Now.ToString("yyyy-M-d" + "HH-mm-ss") + ".xls";
            FileStream file = new FileStream(filename, FileMode.Create, FileAccess.Write);
            workbook.Write(file);
            file.Close();

        }

        private void button7_Click(object sender, EventArgs e)
        {
            //建立新excel檔
            IWorkbook workbook = new HSSFWorkbook();
            ISheet sheet = workbook.CreateSheet("new sheet");

            sheet.CreateRow(0).CreateCell(0).SetCellValue("身分證"); //d3
            sheet.GetRow(0).CreateCell(1).SetCellValue("案件分類"); //d1
            sheet.GetRow(0).CreateCell(2).SetCellValue("流水號");  //d2
            sheet.GetRow(0).CreateCell(3).SetCellValue("姓名");   //d103
            sheet.GetRow(0).CreateCell(4).SetCellValue("入院日期"); //d10
            sheet.GetRow(0).CreateCell(5).SetCellValue("出院日期"); //d11
            sheet.GetRow(0).CreateCell(6).SetCellValue("主治醫師"); //d20
            sheet.GetRow(0).CreateCell(7).SetCellValue("床號"); //p9
            sheet.GetRow(0).CreateCell(8).SetCellValue("時間起"); //p14
            sheet.GetRow(0).CreateCell(9).SetCellValue("時間迄"); //d28

            n = 0;
            rowNumber = 0;

            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(textBox1.Text);

            XmlNode tdataNode = xmlDoc.SelectSingleNode("inpatient/tdata");
            XmlNodeList tdataNodeChildNodeList = tdataNode.ChildNodes;

            foreach (XmlNode tNode in tdataNodeChildNodeList)
            {
                if (tNode.Name == "t3")
                {
                    year = tNode.InnerText;
                }
            }

            XmlNodeList nodeLists = xmlDoc.SelectNodes("inpatient/ddata");
            foreach (XmlNode node in nodeLists)
            {

                XmlNode childNodeHead = node.SelectSingleNode("dhead");
                XmlNodeList dheadNodeList = childNodeHead.ChildNodes;
                foreach (XmlNode child in dheadNodeList)
                {
                    if (child.Name == "d2")//流水號
                        serNo = child.InnerText;
                    if (child.Name == "d1") //案件分類
                        caseNo = child.InnerText;
                }

                XmlNode childNodeBody = node.SelectSingleNode("dbody");
                XmlNodeList dbodyNodeList = childNodeBody.ChildNodes;

                foreach (XmlNode dbodyChild in dbodyNodeList)
                {
                    if (dbodyChild.Name == "d3") //身份證
                    {
                        pID = dbodyChild.InnerText;
                    }
                    if (dbodyChild.Name == "d103") //姓名
                    {
                        name = dbodyChild.InnerText;
                    }

                    if (dbodyChild.Name == "d10") //入院日期
                    {
                        indate = dbodyChild.InnerText;
                    }

                    if (dbodyChild.Name == "d11") //出院日期
                    {
                        outdate = dbodyChild.InnerText;
                    }

                    if (dbodyChild.Name == "d20") //主治醫師
                    {
                        dr = dbodyChild.InnerText;
                    }

                    if (dbodyChild.Name == "pdata")
                    {
                        XmlNodeList pdataList = dbodyChild.ChildNodes;
                        foreach (XmlNode childPdata in pdataList)
                        {

                            string nodePdataName = childPdata.Name;
                            string nodePdataValue = childPdata.InnerText;
                            //MessageBox.Show(nodePdataName + "_" + nodePdataValue);
                            if (nodePdataName == "p1") //醫令序
                                orderList = nodePdataValue.Trim();
                            if (nodePdataName == "p3") //醫令
                                orderCode = nodePdataValue.Trim();
                            if (nodePdataName == "p9") //床號
                                bedno = nodePdataValue.Trim();
                            if (nodePdataName == "p14") //時間起
                                stime = nodePdataValue.Trim().Substring(0, 7);
                            if (nodePdataName == "p15") //時間迄
                                etime = nodePdataValue.Trim().Substring(0, 7);
                        }

                        if (orderCode != null)
                        {
                            if ((orderCode == "03057B" || orderCode == "04002B" || orderCode == "04011B") && Convert.ToInt32(stime) < 1031100)
                            {
                                n++;
                                rowNumber++;
                                sheet.CreateRow(rowNumber).CreateCell(0).SetCellValue(pID);
                                sheet.GetRow(rowNumber).CreateCell(1).SetCellValue(caseNo);
                                sheet.GetRow(rowNumber).CreateCell(2).SetCellValue(serNo);
                                sheet.GetRow(rowNumber).CreateCell(3).SetCellValue(name);
                                sheet.GetRow(rowNumber).CreateCell(4).SetCellValue(indate);
                                sheet.GetRow(rowNumber).CreateCell(5).SetCellValue(outdate);
                                sheet.GetRow(rowNumber).CreateCell(6).SetCellValue(dr);
                                sheet.GetRow(rowNumber).CreateCell(7).SetCellValue(bedno);
                                sheet.GetRow(rowNumber).CreateCell(8).SetCellValue(stime);
                                sheet.GetRow(rowNumber).CreateCell(9).SetCellValue(etime);
                            }
                        }
                    }

                    //MessageBox.Show(pID + "_" + caseNo + "_" + serNo + "_" + orderList + "_" + orderCode + "_" + amount + "_" + cashPoint + "_" + doID);

                }
            }

            string filename = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + @"\轉檔\" + year + "住院明細" + DateTime.Now.ToString("yyyy-M-d" + "HH-mm-ss") + ".xls";
            //MessageBox.Show(filename);
            FileStream file = new FileStream(filename, FileMode.Create, FileAccess.Write);
            workbook.Write(file);
            file.Close();
            MessageBox.Show("done!共" + n);


        }

        private void button8_Click(object sender, EventArgs e)
        {
            IWorkbook workbook = new HSSFWorkbook();
            ISheet sheet = workbook.CreateSheet("new sheet");

            sheet.CreateRow(0).CreateCell(0).SetCellValue("日期");
            sheet.CreateRow(1).CreateCell(0).SetCellValue("許珮珊");
            sheet.CreateRow(2).CreateCell(0).SetCellValue("陳柏偉");
            sheet.CreateRow(3).CreateCell(0).SetCellValue("林典雍");
            sheet.CreateRow(4).CreateCell(0).SetCellValue("詹永騰");
            for (int i = 1; i <= 31; i++)
            {
                sheet.GetRow(0).CreateCell(i).SetCellValue(i);
                sheet.GetRow(1).CreateCell(i).SetCellValue(0);
                sheet.GetRow(2).CreateCell(i).SetCellValue(0);
                sheet.GetRow(3).CreateCell(i).SetCellValue(0);
                sheet.GetRow(4).CreateCell(i).SetCellValue(0);
            }

            var sworkbook = InitializeWorkbook(textBox1.Text);
            int sdate, edate; //時間起迄
            int n;


            for (int rowNumber1 = 1; rowNumber1 < sworkbook.GetSheetAt(0).PhysicalNumberOfRows; rowNumber1++)
            {
                var cell1 = sworkbook.GetSheetAt(0).GetRow(rowNumber1).GetCell(4); //主治醫師
                var cell2 = sworkbook.GetSheetAt(0).GetRow(rowNumber1).GetCell(6); //主治醫師
                var cell3 = sworkbook.GetSheetAt(0).GetRow(rowNumber1).GetCell(7); //主治醫師


                if (cell1 != null)
                {
                    //MessageBox.Show(cell1.ToString());
                    if (cell1.ToString() == "許珮珊")
                    {
                        sdate = Convert.ToInt32(cell2.ToString());
                        edate = Convert.ToInt32(cell3.ToString());

                        if (sdate < 1031101)
                        {
                            sdate = 1;
                        }
                        else
                        {
                            sdate = sdate - 1031100;
                        }
                        if (edate > 1031130)
                        {
                            edate = 30;
                        }
                        else
                        {
                            edate = edate - 1031100;
                        }

                        for (int i = sdate; i <= edate; i++)
                        {
                            n = Convert.ToInt32(sheet.GetRow(1).GetCell(i).ToString()) + 1;
                            sheet.GetRow(1).GetCell(i).SetCellValue(n);
                        }
                    }

                    if (cell1.ToString() == "陳柏偉")
                    {
                        sdate = Convert.ToInt32(cell2.ToString());
                        edate = Convert.ToInt32(cell3.ToString());

                        if (sdate < 1031101)
                        {
                            sdate = 1;
                        }
                        else
                        {
                            sdate = sdate - 1031100;
                        }
                        if (edate > 1031130)
                        {
                            edate = 30;
                        }
                        else
                        {
                            edate = edate - 1031100;
                        }

                        for (int i = sdate; i <= edate; i++)
                        {
                            n = Convert.ToInt32(sheet.GetRow(2).GetCell(i).ToString()) + 1;
                            sheet.GetRow(2).GetCell(i).SetCellValue(n);
                        }
                    }

                    if (cell1.ToString() == "林典雍")
                    {
                        sdate = Convert.ToInt32(cell2.ToString());
                        edate = Convert.ToInt32(cell3.ToString());

                        if (sdate < 1031101)
                        {
                            sdate = 1;
                        }
                        else
                        {
                            sdate = sdate - 1031100;
                        }
                        if (edate > 1031130)
                        {
                            edate = 30;
                        }
                        else
                        {
                            edate = edate - 1031100;
                        }

                        for (int i = sdate; i <= edate; i++)
                        {
                            n = Convert.ToInt32(sheet.GetRow(3).GetCell(i).ToString()) + 1;
                            sheet.GetRow(3).GetCell(i).SetCellValue(n);
                        }
                    }

                    if (cell1.ToString() == "詹永騰")
                    {
                        sdate = Convert.ToInt32(cell2.ToString());
                        edate = Convert.ToInt32(cell3.ToString());

                        if (sdate < 1031101)
                        {
                            sdate = 1;
                        }
                        else
                        {
                            sdate = sdate - 1031100;
                        }
                        if (edate > 1031130)
                        {
                            edate = 30;
                        }
                        else
                        {
                            edate = edate - 1031100;
                        }

                        for (int i = sdate; i <= edate; i++)
                        {
                            n = Convert.ToInt32(sheet.GetRow(4).GetCell(i).ToString()) + 1;
                            sheet.GetRow(4).GetCell(i).SetCellValue(n);
                        }
                    }

                }
            }

            string filename = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + @"\轉檔\" + year + "住院加總" + DateTime.Now.ToString("yyyy-M-d" + "HH-mm-ss") + ".xls";
            //MessageBox.Show(filename);
            FileStream file = new FileStream(filename, FileMode.Create, FileAccess.Write);
            workbook.Write(file);
            file.Close();

            MessageBox.Show("done");

        }

        static IWorkbook InitializeWorkbook(string filepath)
        {
            IWorkbook workbook;
            FileStream file = new FileStream(filepath, FileMode.Open, FileAccess.Read);
            workbook = new HSSFWorkbook(file);
            file.Close();
            return workbook;
        }

        private void button9_Click(object sender, EventArgs e)
        {
            //建立新excel檔

            IWorkbook workbook = new HSSFWorkbook();
            ISheet sheet = workbook.CreateSheet("new sheet");

            sheet.CreateRow(0).CreateCell(0).SetCellValue("身分證");
            sheet.GetRow(0).CreateCell(1).SetCellValue("案件分類");
            sheet.GetRow(0).CreateCell(2).SetCellValue("流水號");
            sheet.GetRow(0).CreateCell(3).SetCellValue("醫令序");
            sheet.GetRow(0).CreateCell(4).SetCellValue("醫令代碼");
            sheet.GetRow(0).CreateCell(5).SetCellValue("申報數量");
            sheet.GetRow(0).CreateCell(6).SetCellValue("申報金額");
            sheet.GetRow(0).CreateCell(7).SetCellValue("醫事人員");
            sheet.GetRow(0).CreateCell(8).SetCellValue("主治醫師");

            n = 0;
            rowNumber = 0;

            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(textBox1.Text);

            XmlNode tdataNode = xmlDoc.SelectSingleNode("outpatient/tdata");
            XmlNodeList tdataNodeChildNodeList = tdataNode.ChildNodes;

            foreach (XmlNode tNode in tdataNodeChildNodeList)
            {
                if (tNode.Name == "t3")
                {
                    year = tNode.InnerText;
                }
            }

            XmlNodeList nodeLists = xmlDoc.SelectNodes("outpatient/ddata");
            foreach (XmlNode ddataNode in nodeLists)
            {
                XmlNode childNodeHead = ddataNode.SelectSingleNode("dhead");
                XmlNodeList dheadNodeList = childNodeHead.ChildNodes;
                foreach (XmlNode child in dheadNodeList)
                {
                    if (child.Name == "d2")//流水號
                        serNo = child.InnerText;
                    if (child.Name == "d1") //案件分類
                        caseNo = child.InnerText;
                }

                XmlNode childNodeBody = ddataNode.SelectSingleNode("dbody");
                XmlNodeList dbodyNodeList = childNodeBody.ChildNodes;

                foreach (XmlNode dbodyChild in dbodyNodeList)
                {
                    if (dbodyChild.Name == "d3") //身份證
                    {
                        pID = dbodyChild.InnerText;
                    }
                    if (dbodyChild.Name == "d30") //身份證
                    {
                        dr = dbodyChild.InnerText;
                    }

                    if (dbodyChild.Name == "pdata")
                    {
                        XmlNodeList pdataList = dbodyChild.ChildNodes;
                        foreach (XmlNode childPdata in pdataList)
                        {

                            string nodePdataName = childPdata.Name;
                            string nodePdataValue = childPdata.InnerText;
                            //MessageBox.Show(nodePdataName + "_" + nodePdataValue);
                            if (nodePdataName == "p13")
                                orderList = nodePdataValue.Trim();
                            if (nodePdataName == "p4")
                                orderCode = nodePdataValue.Trim();
                            if (nodePdataName == "p10")
                                amount = nodePdataValue.Trim();
                            if (nodePdataName == "p12")
                                cashPoint = nodePdataValue.Trim();
                            if (nodePdataName == "p16")
                                doID = nodePdataValue.Trim();
                        }
                       
                        if (orderCode != null)
                        {


                            if (orderCode == "45010C" || orderCode == "45087C" || orderCode == "45013C" || orderCode == "45046C" || orderCode == "45085B" || orderCode == "45082B" || orderCode == "45102C")
                            {
                                n++;
                                rowNumber++;
                                sheet.CreateRow(rowNumber).CreateCell(0).SetCellValue(pID);
                                sheet.GetRow(rowNumber).CreateCell(1).SetCellValue(caseNo);
                                sheet.GetRow(rowNumber).CreateCell(2).SetCellValue(serNo);
                                sheet.GetRow(rowNumber).CreateCell(3).SetCellValue(orderList);
                                sheet.GetRow(rowNumber).CreateCell(4).SetCellValue(orderCode);
                                sheet.GetRow(rowNumber).CreateCell(5).SetCellValue(amount);
                                sheet.GetRow(rowNumber).CreateCell(6).SetCellValue(cashPoint);
                                sheet.GetRow(rowNumber).CreateCell(7).SetCellValue(doID);
                                sheet.GetRow(rowNumber).CreateCell(8).SetCellValue(dr);
                            }
                        }
                    }

                    //MessageBox.Show(pID + "_" + caseNo + "_" + serNo + "_" + orderList + "_" + orderCode + "_" + amount + "_" + cashPoint + "_" + doID);

                }
            }

            string filename = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + @"\轉檔\" + year + "門診明細" + DateTime.Now.ToString("yyyy-M-d" + "HH-mm-ss") + ".xls";
            FileStream file = new FileStream(filename, FileMode.Create, FileAccess.Write);
            workbook.Write(file);
            file.Close();
            MessageBox.Show("done!共" + n);

        }
    }
}
