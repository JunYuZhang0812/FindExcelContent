using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace FindExcelContent
{
    public partial class FindExcelContent : Form
    {
        public class ResultInfo
        {
            public string m_sourceName;
            public string m_filePath;
            public int m_sheetIndex;
            public string m_sheetName;
            public int m_row;
            public int m_col;
            public string m_content;
            public string GetContent()
            {
                return string.Format("行:{0} 列:{1} 疑似内容：{2}", m_row, m_col, m_content);
            }
        }
        public class DelayFunc
        {
            public float m_dt;
            public float m_time;
            public Action m_func;

            public DelayFunc(float time , Action func)
            {
                m_dt = 0;
                m_time = time;
                m_func = func;
            }
        }
        private bool isCancel = false;
        private List<string> m_fileList = null;
        //[sourceName][filePath][sheetName] = ResultInfo;
        private Dictionary<string, Dictionary<string, Dictionary<string, List<ResultInfo>>>> m_resultDic = new Dictionary<string, Dictionary<string, Dictionary<string, List<ResultInfo>>>>();

        private string m_currFilePath = "";
        private int m_currSheetIndex;
        private string m_currSheetName = "";

        private List<DelayFunc> m_delayFuncList = new List<DelayFunc>();
        public FindExcelContent()
        {
            InitializeComponent();
            InitPanel();
            //if( Logic.IsExternalCall )
            {
                Delay(0.2f, () =>
                {
                    m_btnBegin_Click(null, null);
                });
            }
        }
        private void InitPanel()
        {
            m_textProgress.Visible = false;
            m_sliderProgress.Visible = false;
            m_textDirPath.Text = Logic.DirPath;
            m_timer.Interval = 100;
            m_timer.Start();
        }

        private void m_textDirPath_TextChanged(object sender, EventArgs e)
        {
            Logic.DirPath = m_textDirPath.Text;
        }

        private void m_btnSelectDir_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            dialog.Description = "请选择文件夹路径";
            dialog.SelectedPath = m_textDirPath.Text;
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                string str = dialog.SelectedPath;
                m_textDirPath.Text = str;
            }
        }

        private void m_btnBegin_Click(object sender, EventArgs e)
        {
            if (m_asyncWorker.IsBusy)
            {
                isCancel = true;
                m_asyncWorker.CancelAsync();
                return;
            }
            if (!Logic.Check()) return;
            if(Logic.IsExternalCall && Logic.FindNameList.Count <= 0)
            {
                MessageBox.Show("请选择要查找的文件");
                return;
            }
            var result = MessageBox.Show("请确认配置表文件夹路径是否正确，如需修改，请点击[否]，然后手动选择路径，最后点击[开始]。\r\n另外请关闭当前文件夹下打开的所有Excel！", "提示", MessageBoxButtons.YesNo);
            if( result == DialogResult.No )
            {
                return;
            }
            m_fileList = Logic.GetFilePathList();
            m_sliderProgress.Maximum = m_fileList.Count;
            m_sliderProgress.Minimum = 0;
            m_sliderProgress.Visible = true;
            m_textProgress.Visible = true;
            m_textProgress.Text = "开始加载Excel";
            isCancel = false;
            m_btnBegin.Text = "停止";
            m_resultDic.Clear();
            m_tree.Nodes.Clear();

            if(Logic.IsExternalCall)
            {
                var findNameList = Logic.FindNameList;
                for (int i = 0; i < findNameList.Count; i++)
                {
                    var findName = findNameList[i];
                    m_resultDic.Add(findName, new Dictionary<string, Dictionary<string, List<ResultInfo>>>());
                }
            }

            m_asyncWorker.RunWorkerAsync();
        }

        private void m_asyncWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            for (int j = 0; j < m_fileList.Count; j++)
            {
                if (isCancel)
                    return;
                if (m_asyncWorker.WorkerReportsProgress)
                    m_asyncWorker.ReportProgress(j);
                var path = m_fileList[j];
                m_currFilePath = path;
                CheckExcel(path);
            }
        }
        private bool CheckExcel( string path )
        {
            IWorkbook workbook = null;
            try
            {
                using (FileStream file = new FileStream(path, FileMode.Open, FileAccess.Read))
                {
                    if (path.EndsWith(".xls"))
                        workbook = new HSSFWorkbook(file);
                    else
                        workbook = new XSSFWorkbook(file);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("[" + path + "]:" + ex.ToString());
                return false;
            }
            if (workbook != null)
            {
                for (int sheetIndex = 0; sheetIndex < workbook.NumberOfSheets; sheetIndex++)
                {
                    try
                    {
                        var sheet = workbook.GetSheetAt(sheetIndex);
                        if (sheet != null)
                        {
                            m_currSheetIndex = sheetIndex;
                            m_currSheetName = sheet.SheetName;
                            CheckSheet(sheet);
                        }
                    }
                    catch (Exception e)
                    {
                        MessageBox.Show(e.ToString());
                    }
                }
                workbook.Close();
            }
            return true;
        }

        private void CheckSheet( ISheet sheet )
        {
            var maxRow = sheet.LastRowNum;
            for (int row = 0; row < maxRow; row++)
            {
                var rowData = sheet.GetRow(row);
                if (rowData == null) break;
                var maxCol = rowData.LastCellNum;
                for (int col = 0; col < maxCol; col++)
                {
                    var cell = rowData.GetCell(col);
                    if (cell == null) continue;
                    var value = GetCellValue(cell);
                    if( Logic.IsExternalCall )
                    {
                        var findNameList = Logic.FindNameList;
                        for (int i = 0; i < findNameList.Count; i++)
                        {
                            var findName = findNameList[i];
                            if (value.Contains(findName))
                            {
                                AddResult(findName, row, col, value);
                            }
                        }
                    }
                }
            }
        }
        private string GetCellValue(ICell cell)
        {
            if (cell == null) return "";
            var type = cell.CellType;
            switch (type)
            {
                case CellType.Unknown:
                    return "";
                case CellType.Numeric:
                    return cell.NumericCellValue.ToString(); ;
                case CellType.String:
                    return cell.StringCellValue;
                case CellType.Formula:
                    return "";
                case CellType.Blank:
                    return "";
                case CellType.Boolean:
                    return cell.BooleanCellValue.ToString();
                case CellType.Error:
                    return cell.ErrorCellValue.ToString();
            }
            return "";
        }
        private void AddResult(string findName , int row , int col,  string value )
        {
            Dictionary<string, Dictionary<string, List<ResultInfo>>> fileDic;
            if ( !m_resultDic.ContainsKey(findName))
            {
                fileDic = new Dictionary<string, Dictionary<string, List<ResultInfo>>>();
                m_resultDic.Add(findName, fileDic);
            }
            else
            {
                fileDic = m_resultDic[findName];
            }
            Dictionary<string, List<ResultInfo>> sheetDic;
            if (!fileDic.ContainsKey(m_currFilePath))
            {
                sheetDic = new Dictionary<string, List<ResultInfo>>();
                fileDic.Add(m_currFilePath, sheetDic);
            }
            else
            {
                sheetDic = fileDic[m_currFilePath];
            }
            List<ResultInfo> resultList;
            if( !sheetDic.ContainsKey(m_currSheetName) )
            {
                resultList = new List<ResultInfo>();
                sheetDic.Add(m_currSheetName, resultList);
            }
            else
            {
                resultList = sheetDic[m_currSheetName];
            }
            resultList.Add(new ResultInfo()
            {
                m_sourceName = findName,
                m_filePath = m_currFilePath,
                m_sheetIndex = m_currSheetIndex,
                m_sheetName = m_currSheetName,
                m_row = row,
                m_col = col,
                m_content = value
            });
        }

        private void RefreshTree()
        {
            m_tree.Nodes.Clear();
            foreach (var sourceDic in m_resultDic)
            {
                var sourceName = sourceDic.Key;
                bool notHave = sourceDic.Value.Count <= 0;
                var name = sourceName;
                if (notHave)
                {
                    name += "    (未找到引用)";
                }
                var node = m_tree.Nodes.Add(sourceName, name);
                node.Tag = sourceName;
                if (notHave)
                {
                    node.ForeColor = Color.White;
                    node.BackColor = Color.Red;
                }
                else
                {
                    foreach (var fileDic in sourceDic.Value)
                    {
                        var fileName = Path.GetFileNameWithoutExtension(fileDic.Key);
                        var node2 = node.Nodes.Add(sourceName + "_" + fileName, fileName);
                        node2.Tag = fileName;
                        foreach (var sheetDic in fileDic.Value)
                        {
                            var sheetName = sheetDic.Key;
                            var node3 = node2.Nodes.Add(sourceName + "_" + fileName + "_" + sheetName, sheetName);
                            node3.Tag = sheetName;
                            var resuleList = sheetDic.Value;
                            for (int i = 0; i < resuleList.Count; i++)
                            {
                                var result = resuleList[i];
                                var node4 = node3.Nodes.Add(sourceName + "_" + fileName + "_" + sheetName + "_" + result.m_row.ToString() + "_" + result.m_col.ToString(), result.GetContent());
                                node4.Tag = result;
                            }
                        }
                    }
                }
            }
        }
        private void m_asynWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            int currPro = e.ProgressPercentage;
            if (currPro > m_sliderProgress.Maximum)
                currPro = m_sliderProgress.Maximum;
            m_sliderProgress.Value = currPro;
            var str = Path.GetFileName(m_fileList[currPro]) + "   " + Math.Floor(100d * currPro / m_sliderProgress.Maximum) + "%";
            m_textProgress.Text = str;
        }
        private void m_asynWorker_Complete(object sender, RunWorkerCompletedEventArgs e)
        {
            m_sliderProgress.Visible = false;
            m_textProgress.Visible = false;
            m_btnBegin.Text = "开始";
            if (isCancel)
            {
                isCancel = false;
                MessageBox.Show("用户取消");
                return;
            }
            RefreshTree();
        }
        private void m_tree_NodeClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            if (e.Button != MouseButtons.Right) return;
            var node = e.Node;
            var tag = node.Tag;
            var menu = new ContextMenu();
            string str;
            if( tag is ResultInfo )
            {
                var result = tag as ResultInfo;
                str = result.m_content;
                menu.MenuItems.Add("打开", (_s, _e) =>
                {
                    Process.Start(result.m_filePath);
                });
            }
            else
            {
                str = tag.ToString();
            }
            menu.MenuItems.Add("复制", (_s,_e) =>
             {
                 Clipboard.Clear();
                 Clipboard.SetData(DataFormats.Text, str);
             });
            node.ContextMenu = menu;
        }
        private void m_tree_NodeDoubleClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            if (e.Button != MouseButtons.Left) return;
            var node = e.Node;
            var result = node.Tag as ResultInfo;
            if (result == null) return;
            Process.Start(result.m_filePath );
        }

        private void Delay(float time , Action action)
        {
            m_delayFuncList.Add(new DelayFunc(time, action));
        }

        private void m_timer_Tick(object sender, EventArgs e)
        {
            var count = m_delayFuncList.Count;
            if (count <= 0) return;
            for (int i = count-1; i >= 0; i--)
            {
                var delay = m_delayFuncList[i];
                delay.m_dt += m_timer.Interval / 1000.0f;
                if(delay.m_dt >= delay.m_time)
                {
                    m_delayFuncList.RemoveAt(i);
                    delay.m_func.Invoke();
                }
                if (m_delayFuncList.Count <= 0)
                {
                    break;
                }
            }
        }
    }
}
