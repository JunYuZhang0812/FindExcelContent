
namespace FindExcelContent
{
    partial class FindExcelContent
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.label1 = new System.Windows.Forms.Label();
            this.m_textDirPath = new System.Windows.Forms.TextBox();
            this.m_btnSelectDir = new System.Windows.Forms.Button();
            this.m_tree = new System.Windows.Forms.TreeView();
            this.m_asyncWorker = new System.ComponentModel.BackgroundWorker();
            this.m_sliderProgress = new System.Windows.Forms.ProgressBar();
            this.m_textProgress = new System.Windows.Forms.Label();
            this.m_btnBegin = new System.Windows.Forms.Button();
            this.m_timer = new System.Windows.Forms.Timer(this.components);
            this.m_textSearch = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(77, 12);
            this.label1.TabIndex = 0;
            this.label1.Text = "文件夹路径：";
            // 
            // m_textDirPath
            // 
            this.m_textDirPath.Location = new System.Drawing.Point(95, 6);
            this.m_textDirPath.Name = "m_textDirPath";
            this.m_textDirPath.Size = new System.Drawing.Size(326, 21);
            this.m_textDirPath.TabIndex = 1;
            this.m_textDirPath.TextChanged += new System.EventHandler(this.m_textDirPath_TextChanged);
            // 
            // m_btnSelectDir
            // 
            this.m_btnSelectDir.Location = new System.Drawing.Point(427, 6);
            this.m_btnSelectDir.Name = "m_btnSelectDir";
            this.m_btnSelectDir.Size = new System.Drawing.Size(75, 21);
            this.m_btnSelectDir.TabIndex = 2;
            this.m_btnSelectDir.Text = "选择文件夹";
            this.m_btnSelectDir.UseVisualStyleBackColor = true;
            this.m_btnSelectDir.Click += new System.EventHandler(this.m_btnSelectDir_Click);
            // 
            // m_tree
            // 
            this.m_tree.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.m_tree.Location = new System.Drawing.Point(12, 33);
            this.m_tree.Name = "m_tree";
            this.m_tree.Size = new System.Drawing.Size(776, 379);
            this.m_tree.TabIndex = 3;
            this.m_tree.NodeMouseClick += new System.Windows.Forms.TreeNodeMouseClickEventHandler(this.m_tree_NodeClick);
            this.m_tree.NodeMouseDoubleClick += new System.Windows.Forms.TreeNodeMouseClickEventHandler(this.m_tree_NodeDoubleClick);
            // 
            // m_asyncWorker
            // 
            this.m_asyncWorker.WorkerReportsProgress = true;
            this.m_asyncWorker.WorkerSupportsCancellation = true;
            this.m_asyncWorker.DoWork += new System.ComponentModel.DoWorkEventHandler(this.m_asyncWorker_DoWork);
            this.m_asyncWorker.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(this.m_asynWorker_ProgressChanged);
            this.m_asyncWorker.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.m_asynWorker_Complete);
            // 
            // m_sliderProgress
            // 
            this.m_sliderProgress.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.m_sliderProgress.Location = new System.Drawing.Point(12, 418);
            this.m_sliderProgress.Name = "m_sliderProgress";
            this.m_sliderProgress.Size = new System.Drawing.Size(776, 23);
            this.m_sliderProgress.TabIndex = 4;
            // 
            // m_textProgress
            // 
            this.m_textProgress.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.m_textProgress.AutoSize = true;
            this.m_textProgress.BackColor = System.Drawing.Color.Transparent;
            this.m_textProgress.Location = new System.Drawing.Point(18, 424);
            this.m_textProgress.Name = "m_textProgress";
            this.m_textProgress.Size = new System.Drawing.Size(71, 12);
            this.m_textProgress.TabIndex = 5;
            this.m_textProgress.Text = "11111111111";
            // 
            // m_btnBegin
            // 
            this.m_btnBegin.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.m_btnBegin.Location = new System.Drawing.Point(713, 6);
            this.m_btnBegin.Name = "m_btnBegin";
            this.m_btnBegin.Size = new System.Drawing.Size(75, 21);
            this.m_btnBegin.TabIndex = 6;
            this.m_btnBegin.Text = "开始";
            this.m_btnBegin.UseVisualStyleBackColor = true;
            this.m_btnBegin.Click += new System.EventHandler(this.m_btnBegin_Click);
            // 
            // m_timer
            // 
            this.m_timer.Tick += new System.EventHandler(this.m_timer_Tick);
            // 
            // m_textSearch
            // 
            this.m_textSearch.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.m_textSearch.Location = new System.Drawing.Point(595, 6);
            this.m_textSearch.Name = "m_textSearch";
            this.m_textSearch.Size = new System.Drawing.Size(100, 21);
            this.m_textSearch.TabIndex = 7;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(524, 10);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(65, 12);
            this.label2.TabIndex = 8;
            this.label2.Text = "查找内容：";
            // 
            // FindExcelContent
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 443);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.m_textSearch);
            this.Controls.Add(this.m_textProgress);
            this.Controls.Add(this.m_btnBegin);
            this.Controls.Add(this.m_sliderProgress);
            this.Controls.Add(this.m_tree);
            this.Controls.Add(this.m_btnSelectDir);
            this.Controls.Add(this.m_textDirPath);
            this.Controls.Add(this.label1);
            this.Name = "FindExcelContent";
            this.Text = "查找Excel引用";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox m_textDirPath;
        private System.Windows.Forms.Button m_btnSelectDir;
        private System.Windows.Forms.TreeView m_tree;
        private System.ComponentModel.BackgroundWorker m_asyncWorker;
        private System.Windows.Forms.ProgressBar m_sliderProgress;
        private System.Windows.Forms.Label m_textProgress;
        private System.Windows.Forms.Button m_btnBegin;
        private System.Windows.Forms.Timer m_timer;
        private System.Windows.Forms.TextBox m_textSearch;
        private System.Windows.Forms.Label label2;
    }
}

