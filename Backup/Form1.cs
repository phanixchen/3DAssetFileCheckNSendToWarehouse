using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;
using CS_Class;

namespace _F_Shelves
{
    public partial class Form1 : Form
    {
        private string sql;
        private SQLConn cn;
        private const string _databasePath = "\\\\10.91.21.40\\Database";
        private const string _amPath = "\\\\10.91.21.40\\AmDatabase";
        private const string _mocapPath = "\\\\10.91.21.40\\MoCapDb";
        private const string _indexPath = "\\\\10.91.21.40\\Index";
        //private const string _databasePath = "\\\\140.113.124.12\\f$\\Database";
        //private const string _amPath = "\\\\140.113.124.12\\f$\\Database";
        //private const string _mocapPath = "\\\\140.113.124.12\\f$\\Database";
        //private const string _indexPath = "\\\\140.113.124.12\\f$\\Index";
        
        public Form1()
        {
            InitializeComponent();

            tbCreator.Text = Environment.UserName;

            try
            {
                StreamReader sr = new StreamReader("config.ini");
                cn = new SQLConn(sr.ReadLine().Trim(), sr.ReadLine().Trim(), sr.ReadLine().Trim(), sr.ReadLine().Trim());
                sr.Close();
            }
            catch
            {
                MessageBox.Show("讀取 config.ini 失敗，無法建立連線!\n程式即將關閉");
                Environment.Exit(0);
            }
            

            DataGridViewCellStyle style = new DataGridViewCellStyle();
            style.BackColor = Color.FromArgb(180, 180, 180);

            dg1.Columns["type1"].DefaultCellStyle = style;
            dg1.Columns["type2"].DefaultCellStyle = style;
            dg1.Columns["type3"].DefaultCellStyle = style;
            dg1.Columns["type4"].DefaultCellStyle = style;
            dg1.Columns["type5"].DefaultCellStyle = style;
            dg1.Columns["type6"].DefaultCellStyle = style;

            dg1.Columns["char_value1"].DefaultCellStyle = style;
            dg1.Columns["char_value2"].DefaultCellStyle = style;

            dg1.Columns["tag"].DefaultCellStyle = style;
        }

        private void cbDB_SelectedIndexChanged(object sender, EventArgs e)
        {
            //generate tree view structure
            tree1.Nodes.Clear();
            if (cbDB.SelectedItem != null && cbDB.SelectedItem.ToString() != "")
            {
                funcGenerateTree(null, 0);
            }

            if (tabControl1.SelectedTab == tabPage1)    //多筆新增
            {
                funcNewRow(0, dg1.Rows.Count - 1);
            }
            else if (tabControl1.SelectedTab == tabPage3)   //檢視資料
            {
                funcGenerate_dgView();
            }
            else if (tabControl1.SelectedTab == tabPage4)
            {
            }
            else if (tabControl1.SelectedTab == tabPage5)
            {
                //強制退回前一個 tab
                tabControl1.SelectedTab = tabPage4;
            }
        }

        private void funcGenerateTree(TreeNode tn, int _parent)
        {
            DataTable dt = new DataTable();
            TreeNode newtn;
            sql = "select * from " + cbDB.SelectedItem.ToString() + "_structure where parent = '" + _parent + "' order by folderid;";
            cn.Exec(sql, ref dt);

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                if (tn == null)
                {
                    tree1.Nodes.Add(dt.Rows[i]["folderid"].ToString(), dt.Rows[i]["Cht"].ToString() + "_" + dt.Rows[i]["Target"].ToString());
                    newtn = tree1.Nodes[tree1.Nodes.Count - 1];
                    funcGenerateTree(newtn, (int)dt.Rows[i]["folderid"]);
                }
                else
                {
                    tn.Nodes.Add(dt.Rows[i]["folderid"].ToString(), dt.Rows[i]["Cht"].ToString() + "_" + dt.Rows[i]["Target"].ToString());
                    newtn = tn.Nodes[tn.Nodes.Count - 1];
                    funcGenerateTree(newtn, (int)dt.Rows[i]["folderid"]);
                }
            }
        }


        #region Change Focus
        private void tbPath1_Enter(object sender, EventArgs e)
        {
            tbPath1.SelectAll();
        }

        private void tbPath2_Enter(object sender, EventArgs e)
        {
            tbPath2.SelectAll();
        }

        private void tbPath3_Enter(object sender, EventArgs e)
        {
            tbPath3.SelectAll();
        }

        private void tbPath4_Enter(object sender, EventArgs e)
        {
            tbPath4.SelectAll();
        }

        private void tbPath5_Enter(object sender, EventArgs e)
        {
            tbPath5.SelectAll();
        }

        private void tbPath6_Enter(object sender, EventArgs e)
        {
            tbPath6.SelectAll();
        }

        private void tbFilename_Enter(object sender, EventArgs e)
        {
            tbFilename.SelectAll();
        }

        private void tbName_Enter(object sender, EventArgs e)
        {
            tbName.SelectAll();
        }

        private void tbType1_Enter(object sender, EventArgs e)
        {
            tbType1.SelectAll();
        }

        private void tbType2_Enter(object sender, EventArgs e)
        {
            tbType2.SelectAll();
        }

        private void tbType3_Enter(object sender, EventArgs e)
        {
            tbType3.SelectAll();
        }

        private void tbType4_Enter(object sender, EventArgs e)
        {
            tbType4.SelectAll();
        }

        private void tbType5_Enter(object sender, EventArgs e)
        {
            tbType5.SelectAll();
        }

        private void tbType6_Enter(object sender, EventArgs e)
        {
            tbType6.SelectAll();
        }

        private void tbcvalue1_Enter(object sender, EventArgs e)
        {
            tbcvalue1.SelectAll();
        }

        private void tbcvalue2_Enter(object sender, EventArgs e)
        {
            tbcvalue2.SelectAll();
        }

        private void tbKeyword_Enter(object sender, EventArgs e)
        {
            tbKeyword.SelectAll();
        }

        private void tbTag_Enter(object sender, EventArgs e)
        {
            tbTag.SelectAll();
        }
        #endregion


        
        //新增到資料庫 (單筆)
        private void bAdd_Click(object sender, EventArgs e)
        {

        }

        //新增資料 (多筆)
        private void bAddMultiple_Click(object sender, EventArgs e)
        {
            int i, j;
            string keyword;
            string strTag;
            string[] path = new string[6];
            string[] type = new string[6];
            string[] char_value = new string[2];

            if (cbDB.SelectedItem == "")
            {
                MessageBox.Show("請選擇資料庫");
            }
            else if (tbCreator.Text.Trim() == "")
            {
                MessageBox.Show("請填寫資料建立人");
            }
            else
            {
                sql = "";
                for (i = 0; i < dg1.Rows.Count - 1; i++)
                {
                    if (dg1.Rows[i].Cells["keyword"].Value == null)
                    {
                        MessageBox.Show("有些資料沒有指定關鍵字喔！");
                        return;
                    }
                }

                for (i = 0; i < dg1.Rows.Count - 1; i++)
                {
                    keyword = dg1.Rows[i].Cells["keyword"].Value.ToString();
                    while (keyword.Substring(0, 1) == ";")
                    {
                        keyword = keyword.Substring(1);
                    }
                    keyword = ";" + dg1.Rows[i].Cells["filename"].Value.ToString() + ";" + dg1.Rows[i].Cells["name"].Value.ToString() + ";" + keyword;


                    //把 path 補上 default value (empty string)
                    for (j = 1; j <= 6; j++)
                    {
                        if (dg1.Rows[i].Cells["path" + j.ToString()].Value == null)
                        {
                            path[j - 1] = "NULL";
                        }
                        else if (dg1.Rows[i].Cells["path" + j.ToString()].Value == "")
                        {
                            path[j - 1] = "NULL";
                        }
                        else
                        {
                            path[j - 1] = "'" + dg1.Rows[i].Cells["path" + j.ToString()].Value + "'";
                        }
                    }

                    //把 type 補上 default value (empty string)
                    for (j = 1; j <= 6; j++)
                    {
                        if (dg1.Rows[i].Cells["type" + j.ToString()].Value == null)
                        {
                            type[j - 1] = "NULL";
                        }
                        else if (dg1.Rows[i].Cells["type" + j.ToString()].Value == "")
                        {
                            type[j - 1] = "NULL";
                        }
                        else
                        {
                            type[j - 1] = "'" + dg1.Rows[i].Cells["type" + j.ToString()].Value + "'";
                        }
                    }

                    //把 char_value 補上 default value (empty string)
                    for (j = 1; j <= 2; j++)
                    {
                        if (dg1.Rows[i].Cells["char_value" + j.ToString()].Value == null)
                        {
                            char_value[j - 1] = "NULL";
                        }
                        else if (dg1.Rows[i].Cells["char_value" + j.ToString()].Value == "")
                        {
                            char_value[j - 1] = "NULL";
                        }
                        else
                        {
                            char_value[j - 1] = "'" + dg1.Rows[i].Cells["char_value" + j.ToString()].Value + "'";
                        }
                    }

                    //Tag
                    if (dg1.Rows[i].Cells["tag"].Value == null)
                    {
                        strTag = "NULL";
                    }
                    else if (dg1.Rows[i].Cells["tag"].Value.ToString() == "")
                    {
                        strTag = "NULL";
                    }
                    else
                    {
                        strTag = "'" + dg1.Rows[i].Cells["tag"].Value + "'";
                    }


                    sql = sql + "Insert into " + cbDB.SelectedItem.ToString() + "_data " + 
                                "(path1, path2, path3, path4, path5, path6, filename, name, type1, type2, type3, type4, type5, type6, " +
                                " char_value1, char_value2, keyword, tag, creator) values " +
                                "(" + path[0] + ", " + path[1] + ", " + path[2] + ", " + path[3] + ", " + path[4] + ", " + path[5] + ", " +
                                "'" + dg1.Rows[i].Cells["filename"].Value.ToString() + "', '" + dg1.Rows[i].Cells["name"].Value.ToString() + "', " +
                                "" + type[0] + ", " + type[1] + ", " + type[2] + ", " + type[3] + ", " + type[4] + ", " + type[5] + ", " +
                                "" + char_value[0] + ", " + char_value[1] + ", " +
                                "'" + keyword + "', " + strTag + ", '" + tbCreator.Text.Trim() + "'); ";


                }

                try
                {
                    cn.Exec(sql);

                    MessageBox.Show("資料建立完畢！");
                }
                catch (Exception ex)
                {
                    StreamWriter sw = new StreamWriter("error.txt", false, Encoding.Default);
                    sw.WriteLine(ex.Message);
                    sw.WriteLine("");
                    sw.WriteLine(sql);
                    sw.Close();

                    MessageBox.Show("程式執行錯誤, 錯誤訊息在 error.txt");
                }
            }

        }

        #region 控制滑鼠點選以及滑鼠右鍵選單
        private int _iClickedRow = -1;  //記錄點選哪一個 row
        private int _iClickedCell = -1;
        private int _iCopyRow = -1;     //記錄複製哪一個 row

        private void dg1_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                _iClickedRow = e.RowIndex;
                _iClickedCell = e.ColumnIndex;
                contextMenuStrip1.Show(MousePosition.X, MousePosition.Y);
            }
            else
            {
                return;
            }
        }

        //滑鼠右鍵選單開啟，需要控制一下 copy & paste 選項 enable/disable
        private void contextMenuStrip1_Opened(object sender, EventArgs e)
        {
            if (_iClickedRow == -1)
            {
                return;
            }
            else if (dg1.Rows.Count == 1)
            {
                contextMenuStrip1.Items[0].Enabled = false;
                contextMenuStrip1.Items[1].Enabled = false;
            }
            else if (_iClickedRow == dg1.Rows.Count - 1)
            {
                contextMenuStrip1.Items[0].Enabled = false;
                contextMenuStrip1.Items[1].Enabled = true;
            }
            else
            {
                contextMenuStrip1.Items[0].Enabled = true;
                contextMenuStrip1.Items[1].Enabled = false;
            }
        }


        //copy clicked
        private void copyRowToolStripMenuItem_Click(object sender, EventArgs e)
        {
            _iCopyRow = _iClickedRow;
        }

        //paste clicked
        private void pasteRowToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (_iCopyRow == -1)
            {
                return;
            }
            else if (_iClickedRow != dg1.Rows.Count - 1)
            {
                return;
            }
            else
            {
                dg1.Rows.Add(1);
                //dg1.Rows.Add(dg1.Rows[_iCopyRow].Clone());
                for (int i = 0; i < dg1.Columns.Count; i++)
                {
                    dg1.Rows[_iClickedRow].Cells[i].Value = dg1.Rows[_iCopyRow].Cells[i].Value;
                }
            }
        }

        //貼上 excel data
        private void pasteExcelDataToolStripMenuItem_Click(object sender, EventArgs e)
        {
            IDataObject t;
            StreamReader sr;
            string[] row;

            try
            {
                t = Clipboard.GetDataObject();
                sr = new StreamReader((MemoryStream)t.GetData("csv"), Encoding.Default);
                //sr = new StreamReader(Convert.ChangeType(t.GetData("csv"), System.Type.GetType("System.IO.MemoryStream")));

                int i = 0;
                int j, k, iOriginalRowCount;

                iOriginalRowCount = dg1.Rows.Count;
                while (sr.Peek() != -1)
                {
                    row = sr.ReadLine().Split(',');

                    if (_iClickedRow + i == dg1.Rows.Count - 1)
                    {
                        dg1.Rows.Add();
                    }

                    for (j = 0, k = 0; j < row.Length; j++, k++)
                    {
                        try
                        {
                            if (row[j].Substring(0, 1) == "\"" && row[j + 1].Substring(row[j + 1].Length - 1) == "\"")
                            {
                                dg1.Rows[_iClickedRow + i].Cells[_iClickedCell + k].Value = row[j].Substring(1) + "," + row[j + 1].Substring(0, row[j + 1].Length - 1);
                                j++;
                            }
                            else
                            {
                                dg1.Rows[_iClickedRow + i].Cells[_iClickedCell + k].Value = row[j];
                            }
                        }
                        catch
                        {
                            dg1.Rows[_iClickedRow + i].Cells[_iClickedCell + k].Value = row[j];
                        }
                    }
                    i++;
                }

                if (_iClickedRow + i < iOriginalRowCount)
                {
                }
                else
                {
                    dg1.Rows.RemoveAt(dg1.Rows.Count - 2);
                }
            }
            catch
            {
                MessageBox.Show("不是 Excel 資料");
            }
        }

        private void dg1_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            funcNewRow(dg1.Rows.Count - 2, dg1.Rows.Count - 1);
        }

        private void funcNewRow(int start, int end)
        {
            DataGridViewCellStyle style = new DataGridViewCellStyle();
            int i;

            if (cbDB.SelectedItem == "Mocap" || cbDB.SelectedItem == "AM")
            {
                for (i = start; i < end; i++)
                {
                    dg1.Rows[i].Cells["type1"].Value = "1";
                    dg1.Rows[i].Cells["type2"].Value = "1";
                    dg1.Rows[i].Cells["type3"].Value = "1";
                    dg1.Rows[i].Cells["type4"].Value = "8";
                    dg1.Rows[i].Cells["type5"].Value = "1";
                    dg1.Rows[i].Cells["type6"].Value = "1";
                    dg1.Rows[i].Cells["char_value1"].Value = "0";
                    dg1.Rows[i].Cells["char_value2"].Value = "1";
                }
            }
            else if (cbDB.SelectedItem == "avatar")
            {
                for (i = start; i < end; i++)
                {
                    dg1.Rows[i].Cells["type1"].Value = "1";
                    dg1.Rows[i].Cells["type2"].Value = "1";
                    dg1.Rows[i].Cells["type3"].Value = "2";
                    dg1.Rows[i].Cells["type4"].Value = "8";
                    dg1.Rows[i].Cells["type5"].Value = "5";
                    dg1.Rows[i].Cells["type6"].Value = "0";
                }
            }
            else if (cbDB.SelectedItem == "closet")
            {
                for (i = start; i < end; i++)
                {
                    dg1.Rows[i].Cells["type1"].Value = "1";
                    dg1.Rows[i].Cells["type2"].Value = "1";
                    dg1.Rows[i].Cells["type3"].Value = "1";
                    dg1.Rows[i].Cells["type4"].Value = "1";
                    dg1.Rows[i].Cells["type5"].Value = "1";
                    dg1.Rows[i].Cells["type6"].Value = "1";
                    dg1.Rows[i].Cells["char_value1"].Value = "";
                    dg1.Rows[i].Cells["char_value2"].Value = "";
                }
            }
            else if (cbDB.SelectedItem == "trans")
            {
                for (i = start; i < end; i++)
                {
                    dg1.Rows[i].Cells["type1"].Value = "3";
                    dg1.Rows[i].Cells["type2"].Value = "1";
                    dg1.Rows[i].Cells["type3"].Value = "2";
                    dg1.Rows[i].Cells["type4"].Value = "8";
                    dg1.Rows[i].Cells["type5"].Value = "5";
                    dg1.Rows[i].Cells["type6"].Value = "0";
                    dg1.Rows[i].Cells["char_value1"].Value = "0";
                }
            }
            else if (cbDB.SelectedItem == "model" || cbDB.SelectedItem == "scene")
            {
                for (i = start; i < end; i++)
                {
                    dg1.Rows[i].Cells["type1"].Value = "3";
                    dg1.Rows[i].Cells["type2"].Value = "1";
                    dg1.Rows[i].Cells["type3"].Value = "2";
                    dg1.Rows[i].Cells["type4"].Value = "8";
                    dg1.Rows[i].Cells["type5"].Value = "5";
                    dg1.Rows[i].Cells["type6"].Value = "0";
                }
            }
            else if (cbDB.SelectedItem == "viewpoint")
            {
                for (i = start; i < end; i++)
                {
                    dg1.Rows[i].Cells["type1"].Value = "3";
                    dg1.Rows[i].Cells["type2"].Value = "1";
                    dg1.Rows[i].Cells["type3"].Value = "3";
                    dg1.Rows[i].Cells["type4"].Value = "8";
                    dg1.Rows[i].Cells["type5"].Value = "4";
                    dg1.Rows[i].Cells["type6"].Value = "0";
                    dg1.Rows[i].Cells["char_value1"].Value = "0";
                    dg1.Rows[i].Cells["char_value1"].Value = "1";
                }
            }

        }
        #endregion

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabControl1.SelectedTab == tabPage3)
            {
                funcGenerate_dgView();
            }
            else if (tabControl1.SelectedTab == tabPage4)
            {
                if (cbDB.SelectedItem != null && cbDB.SelectedItem.ToString() != "")
                {
                }
            }
            else if (tabControl1.SelectedTab == tabPage5)
            {
                flowPanel.Controls.Clear();
                if (cbDB.SelectedItem == null)
                {
                    MessageBox.Show("要先選擇資料表喔！");
                    tabControl1.SelectedTab = tabPage4;
                }
                else if (cbDB.SelectedItem.ToString() == "")
                {
                    MessageBox.Show("要先選擇資料表喔！");
                    tabControl1.SelectedTab = tabPage4;
                }
                else if (tree1.SelectedNode == null)
                {
                    MessageBox.Show("要先選擇分類喔！");
                    tabControl1.SelectedTab = tabPage4;
                }
                else if (checkFolders() == false)
                {
                    MessageBox.Show("要先選擇分類喔！");
                    tabControl1.SelectedTab = tabPage4;
                }
                else
                {
                    funcAddPanel();
                }
            }
        }

        //顯示"上傳項目" panel
        private void funcAddPanel()
        {
            Panel p;
            CheckBox cb;
            PictureBox pb;
            Label l;
            TextBox tb, tbChtName;
            ComboBox ddl;
            TreeNode tn;
            string[] tmp;
            string[] sep = {"\\"};
            int i, j, iSerial;

            for (i = 0; i < lbFolders.Items.Count; i++)
            {
                p = new Panel();
                p.Width = 842;
                p.Height = 117;
                p.BorderStyle = BorderStyle.FixedSingle;
                p.Font = new Font("新細明體", (float)9);

                cb = new CheckBox();
                cb.Text = "確認";
                cb.Checked = true;
                cb.Appearance = Appearance.Button;
                cb.BackColor = Color.Orange;
                cb.MinimumSize = new Size(50, 50);
                cb.Size = new Size(50, 50);
                cb.Left = 14;
                cb.Top = 30;
                cb.TextAlign = ContentAlignment.MiddleCenter;
                cb.CheckStateChanged += new System.EventHandler(this.cb_CheckedChanged);
                cb.Parent = p;

                p.Controls.Add(cb);


                pb = new PictureBox();
                pb.Top = 11;
                pb.Left = 70;
                pb.Size = new Size(120, 90);
                pb.ImageLocation = lbFolders.Items[i].ToString() + "\\render\\small.jpg"; //Load Image

                pb.Parent = p;
                p.Controls.Add(pb);


                #region labels
                l = new Label();
                l.Text = "檔案來源";
                l.Width = 53;
                l.Top = 10;
                l.Left = 206;
                l.Parent = p;
                p.Controls.Add(l);

                l = new Label();
                l.Text = "物件名稱";
                l.Width = 53;
                l.Top = 64;
                l.Left = 206;
                l.Parent = p;
                p.Controls.Add(l);

                l = new Label();
                l.Text = "分類項目";
                l.Width = 53;
                l.Top = 36;
                l.Left = 206;
                l.Parent = p;
                p.Controls.Add(l);

                l = new Label();
                l.Text = "關鍵字(以;分隔)";
                l.Width = 90;
                l.Top = 64;
                l.Left = 596;
                l.Parent = p;
                p.Controls.Add(l);

                l = new Label();
                l.Text = "精緻度";
                l.Width = 41;
                l.Top = 36;
                l.Left = 721;
                l.Parent = p;
                p.Controls.Add(l);

                l = new Label();
                l.Text = "預設關鍵字";
                l.Width = 65;
                l.Top = 64;
                l.Left = 400;
                l.Parent = p;
                p.Controls.Add(l);

                l = new Label();
                l.Text = "目的目錄";
                l.Width = 53;
                l.Top = 92;
                l.Left = 206;
                l.Parent = p;
                p.Controls.Add(l);

                #endregion

                #region textbox
                tb = new TextBox();
                tb.ReadOnly = true;
                tb.Size = new Size(561, 22);
                tb.Top = 5;
                tb.Left = 265;
                tb.Parent = p;
                tb.Text = lbFolders.Items[i].ToString();
                p.Controls.Add(tb);

                tb = new TextBox();
                tb.ReadOnly = true;
                tb.Size = new Size(450, 22);
                tb.Top = 33;
                tb.Left = 265;
                tb.Parent = p;
                tb.Text = "";
                tn = tree1.SelectedNode;
                while (tn != null)
                {
                    tb.Text = "\\" + tn.Text + tb.Text;
                    tn = tn.Parent;
                }
                p.Controls.Add(tb);

                //物件名稱
                tbChtName = new TextBox();
                tbChtName.Size = new Size(129, 22);
                tbChtName.Top = 59;
                tbChtName.Left = 265;
                tbChtName.Parent = p;
                tmp = lbFolders.Items[i].ToString().Split(sep, StringSplitOptions.RemoveEmptyEntries);
                tbChtName.Text = tmp[tmp.Length - 2];
                tbChtName.TextChanged += new System.EventHandler(this.tbChtName_TextChanged);
                p.Controls.Add(tbChtName);

                //預設關鍵字
                tb = new TextBox();
                tb.ReadOnly = true;
                tb.Size = new Size(119, 22);
                tb.Top = 58;
                tb.Left = 471;
                tb.Parent = p;
                tn = tree1.SelectedNode;
                while (tn != null)
                {
                    tb.Text = tn.Text.Split('_').GetValue(0).ToString() + ";" + tb.Text;
                    tn = tn.Parent;
                }
                p.Controls.Add(tb);
                p.Controls.Add(tb);


                //自訂關鍵字
                tb = new TextBox();
                tb.Size = new Size(139, 22);
                tb.Top = 59;
                tb.Left = 687;
                tb.Parent = p;
                p.Controls.Add(tb);

                //目錄名稱
                tb = new TextBox();
                tb.Size = new Size(561, 22);
                tb.Top = 87;
                tb.Left = 265;
                tb.Parent = p;

                tmp = lbFolders.Items[i].ToString().Split(sep, StringSplitOptions.RemoveEmptyEntries);
                tmp = tmp[tmp.Length - 1].Split('_');
                tb.Text = "\\B_" + tmp[1] + "_" + tmp[2] + "_" + tmp[3];
                iSerial = Convert.ToInt32(tmp[3]);
                    
                tn = tree1.SelectedNode;
                while (tn != null)
                {
                    tb.Text = "\\" + tn.Text.Split('_').GetValue(1).ToString() + tb.Text;
                    tn = tn.Parent;
                }
                tb.Text = "\\" + cbDB.SelectedItem.ToString() + tb.Text;
                while (Directory.Exists(_databasePath + tb.Text) == true)
                {
                    iSerial++;
                    j = tb.Text.LastIndexOf(tmp[3]);
                    tb.Text = tb.Text.Substring(0, j) + iSerial.ToString().PadLeft(3, '0');
                }
                p.Controls.Add(tb);

                
                #endregion

                #region combobox
                ddl = new ComboBox();
                ddl.DropDownStyle = ComboBoxStyle.DropDownList;
                ddl.Size = new Size(58, 20);
                ddl.Top = 33;
                ddl.Left = 768;
                ddl.Items.AddRange(new object[] { "1", "2", "3", "4", "5" });
                ddl.SelectedIndex = 0;

                ddl.Parent = p;
                p.Controls.Add(ddl);
                #endregion

                p.Parent = flowPanel;
                flowPanel.Controls.Add(p);
            }

            flowPanel.Controls.Remove(pContainer);
        }


        #region 檢視/編輯資料 Tab
        //檢視資料
        private void funcGenerate_dgView()
        {
            DataTable dt = new DataTable();
            if (cbDB.SelectedIndex >= 1)
            {
                sql = "Select * from " + cbDB.SelectedItem + "_data order by modelid;";
                cn.Exec(sql, ref dt);

                dgView.DataSource = dt;
                dgView.Update();


                //設定 datagrid scroll down 到最後一筆資料
                dgView.FirstDisplayedScrollingRowIndex = dt.Rows.Count - 1;
                dgView.Refresh();
                dgView.CurrentCell = dgView.Rows[dt.Rows.Count - 1].Cells[0];
                dgView.Rows[dt.Rows.Count - 1].Selected = true;
            }
            else
            {
                dgView.DataSource = null;
                dgView.Update();
            }
        }

        private void dgView_UserDeletingRow(object sender, DataGridViewRowCancelEventArgs e)
        {
            sql = "Delete from " + cbDB.SelectedItem + "_data where modelid = '" + e.Row.Cells["modelid"].Value.ToString() + "';";
            cn.Exec(sql);
        }

        private void dgView_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            sql = "Update " + cbDB.SelectedItem + "_data " + 
                  "set " + dgView.Columns[e.ColumnIndex].Name + " = '" + dgView.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString() + "' " +
                  "where modelid = '" + dgView.Rows[e.RowIndex].Cells["modelid"].Value.ToString() + "';";
            cn.Exec(sql);
        }
        #endregion

        #region 選擇上架資料夾 tab
        private void bAddFolder_Click(object sender, EventArgs e)
        {
            if (folderB1.ShowDialog() == DialogResult.OK)
            {
                checkFolders(folderB1.SelectedPath);
            }
        }

        //檢查 folder 命名規則
        private void checkFolders(string _path)
        {
            int i = 0;
            bool bExist;

            string[] Dirs = Directory.GetDirectories(_path);
            string[] tmp;
            char[] Sep = { '\\' };

            #region 所選擇的 folder 是否就符合規則
            foreach (string d in Dirs)
            {
                tmp = d.Split(Sep, StringSplitOptions.RemoveEmptyEntries);
                if (tmp[tmp.Length - 1].ToUpper() == "WORK" || tmp[tmp.Length - 1].ToUpper() == "RENDER" || tmp[tmp.Length - 1].ToUpper() == "EXPORT" || tmp[tmp.Length - 1].ToUpper() == "TEXTURE")
                {
                    i++;
                }
            }

            if (i == 4)
            {
                bExist = false;
                for (i = 0; i < lbFolders.Items.Count; i++)
                {
                    if (lbFolders.Items[i].ToString() == _path)
                    {
                        bExist = true;
                        break;
                    }
                }

                if (bExist == false)
                {
                    lbFolders.Items.Add(_path);
                }

                return;
            }
            #endregion

            #region 所選的 folder 不符合規則，那就找其下的 sub folders
            foreach (string d in Dirs)
            {
                checkFolders(d);
            }
            #endregion
        }

        //處理按下 delete 時，要清除所選擇的 folder
        private void lbFolders_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Delete)
            {
                try
                {
                    lbFolders.Items.RemoveAt(lbFolders.SelectedIndex);
                }
                catch
                {
                    //do nothing
                }
            }
        }

        //下一步
        private void bNext_Click(object sender, EventArgs e)
        {
            flowPanel.Controls.Clear();
            if (cbDB.SelectedItem == null)
            {
                MessageBox.Show("要先選擇資料表喔！");
                tabControl1.SelectedTab = tabPage4;
            }
            else if (cbDB.SelectedItem.ToString() == "")
            {
                MessageBox.Show("要先選擇資料表喔！");
                tabControl1.SelectedTab = tabPage4;
            }
            else if (tree1.SelectedNode == null)
            {
                MessageBox.Show("要先選擇分類喔！");
                tabControl1.SelectedTab = tabPage4;
            }
            else if (checkFolders() == false)
            {
                MessageBox.Show("要先選擇分類喔！");
                tabControl1.SelectedTab = tabPage4;
            }
            else
            {
                tabControl1.SelectedTab = tabPage5;
            }
        }

        #endregion


        private void lbErr_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                try
                {
                    if (Directory.Exists(lbErr.SelectedValue.ToString()) == true)
                    {
                        Process.Start(lbErr.SelectedValue.ToString());
                    }
                }
                catch
                {
                }
            }
        }

        private void bCheckErr_Click(object sender, EventArgs e)
        {
        }

        private bool checkFolders()
        {
            int i;
            lbErr.Items.Clear();

            if (cb1.Checked == true)
            {
                //檢查空白路徑與檔案名稱

                for (i = 0; i < lbFolders.Items.Count; i++)
                {
                    if (lbFolders.Items[i].ToString().IndexOf(" ") >= 0)
                    {
                        lbErr.Items.Add(lbFolders.Items[i].ToString());
                        lbErr.Items.Add("檔案路徑名稱包含有空白字元!");
                    }
                }
            }

            if (cb2.Checked == true)
            {
                //貼圖類型
                for (i = 0; i < lbFolders.Items.Count; i++)
                {
                }
            }

            if (cb3.Checked == true)
            {
                //目錄格式
                for (i = 0; i < lbFolders.Items.Count; i++)
                {
                }
            }

            if (cb4.Checked == true)
            {
                //縮圖檔案
                for (i = 0; i < lbFolders.Items.Count; i++)
                {
                }
            }

            if (cb5.Checked == true)
            {
                //AT檔案
                for (i = 0; i < lbFolders.Items.Count; i++)
                {
                }
            }

            if (cb6.Checked == true)
            {
                //max檔案
                for (i = 0; i < lbFolders.Items.Count; i++)
                {
                }
            }

            return false;
        }

        private void bPre_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedTab = tabPage4;
        }

        

        //dynamically control the color of botton-style checkbox
        private void cb_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox cb = (CheckBox)sender;

            if (cb.Checked == false)
            {
                cb.BackColor = SystemColors.Control;
                cb.Text = "未確認";
            }
            else
            {
                cb.BackColor = Color.Orange;
                cb.Text = "確認";
            }
        }

        
        private void tbChtName_TextChanged(object sender, EventArgs e)
        {
            #region 中文名稱變更要補上關鍵字
            TextBox tb = (TextBox)sender;
            TextBox tbKeyword = null;
            TextBox tbClass = null;
            Panel p = (Panel)tb.Parent;
            int i, j, k;
            string[] tmp1, tmp2;
            char[] sep1 = {'\\'};
            char[] sep2 = { '_' };

            j = 0;
            for (i = 0; i < p.Controls.Count; i++)
            {
                if (p.Controls[i].GetType().ToString() == "System.Windows.Forms.TextBox")
                {
                    if (j == 1)
                    {
                        tbClass = (TextBox)p.Controls[i];
                    }
                    if (j == 3)
                    {
                        tbKeyword = (TextBox)p.Controls[i];
                    }
                    j++;
                }
            }

            if (tbKeyword == null || tbClass == null)
                return;

            if ( (tbKeyword.Text.IndexOf(";") >= 0 && tbKeyword.Text.IndexOf(tb.Text + ";") < 0) ||
                 (tbKeyword.Text.IndexOf(";") < 0 && tbKeyword.Text.IndexOf(tb.Text) < 0) )
            {
                tbKeyword.Text = "";
                tmp1 = tbClass.Text.Split(sep1, StringSplitOptions.RemoveEmptyEntries);

                for (i = 0; i < tmp1.Length; i++)
                {
                    tmp2 = tmp1[i].Split(sep2, StringSplitOptions.RemoveEmptyEntries);
                    tbKeyword.Text = tbKeyword.Text + tmp2[0] + ";";
                }

                tbKeyword.Text = tbKeyword.Text + tb.Text + ";";
            }
            #endregion
        }



        private void bCheckFolder_Click(object sender, EventArgs e)
        {
            Panel p;
            string[] tmp;
            TextBox tb;
            int i, j, k, iSerial;

            foreach (Control c in flowPanel.Controls)
            {
                p = (Panel)c;
                tb = null;
                j = 0;

                for (i = 0; i < p.Controls.Count; i++)
                {
                    if (p.Controls[i].GetType().ToString() == "System.Windows.Forms.TextBox")
                    {
                        if (j == 5)
                        {
                            tb = (TextBox)p.Controls[i];
                            break;
                        }
                        j++;
                    }
                }
                if (tb == null)
                    continue;

                tmp = tb.Text.Split('_');
                iSerial = Convert.ToInt32(tmp[tmp.Length - 1]);

                while (Directory.Exists(_databasePath + tb.Text) == true)
                {
                    iSerial++;
                    j = tb.Text.LastIndexOf("_");
                    tb.Text = tb.Text.Substring(0, j) + "_" + iSerial.ToString().PadLeft(3, '0');
                }
            }
        }

        private void bUpload_Click(object sender, EventArgs e)
        {
            TextBox tblocal = null;
            TextBox tbremote = null;
            Panel p;
            int i, j, k;
            
            
            foreach (Control c in flowPanel.Controls)
            {
                p = (Panel)c;
                j = 0;

                for (i = 0; i < p.Controls.Count; i++)
                {
                    if (p.Controls[i].GetType().ToString() == "System.Windows.Forms.TextBox")
                    {
                        if (j == 0)
                        {
                            tblocal = (TextBox)p.Controls[i];
                        }
                        if (j == 5)
                        {
                            tbremote = (TextBox)p.Controls[i];
                        }
                        j++;
                    }

                }

                if (tblocal == null || tbremote == null)
                    continue;

                FileSystem.copyDirectory(tblocal.Text, _databasePath + tbremote.Text);
                FileSystem.copyDirectory(tblocal.Text + "\\render", _indexPath + tbremote.Text +"\\render");
            }
        }

       
    }
}