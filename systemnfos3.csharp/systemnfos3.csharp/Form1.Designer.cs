using System.Windows.Forms;

namespace systemnfos3.csharp
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.Windows.Forms.TreeNode treeNode1 = new System.Windows.Forms.TreeNode("Computers");
            System.Windows.Forms.TreeNode treeNode2 = new System.Windows.Forms.TreeNode("Collections");
            System.Windows.Forms.TreeNode treeNode3 = new System.Windows.Forms.TreeNode("Task");
            System.Windows.Forms.TreeNode treeNode4 = new System.Windows.Forms.TreeNode("Query");
            System.Windows.Forms.TreeNode treeNode5 = new System.Windows.Forms.TreeNode("Results");
            System.Windows.Forms.TreeNode treeNode6 = new System.Windows.Forms.TreeNode("Resource Explorer", new System.Windows.Forms.TreeNode[] {
            treeNode1,
            treeNode2,
            treeNode3,
            treeNode4,
            treeNode5});
            System.Windows.Forms.TreeNode treeNode7 = new System.Windows.Forms.TreeNode("Custom Actions");
            System.Windows.Forms.TreeNode treeNode8 = new System.Windows.Forms.TreeNode("Settings", new System.Windows.Forms.TreeNode[] {
            treeNode7});
            this.MainSplitContainer = new System.Windows.Forms.SplitContainer();
            this.ResourceExplorer = new System.Windows.Forms.TreeView();
            this.UserInputComboBox = new System.Windows.Forms.ComboBox();
            this.LabelResource = new System.Windows.Forms.Label();
            this.SubmitButton = new System.Windows.Forms.Button();
            this.ClearButton = new System.Windows.Forms.Button();
            this.NewButton = new System.Windows.Forms.Button();
            this.StatusStrip1 = new System.Windows.Forms.StatusStrip();
            this.UpTimeStatusLabel = new System.Windows.Forms.ToolStripStatusLabel();
            this.MemoryUsageStatusLabel = new System.Windows.Forms.ToolStripStatusLabel();
            this.CpuUsageStatusLabel = new System.Windows.Forms.ToolStripStatusLabel();
            this.ConnectionsStatusLabel = new System.Windows.Forms.ToolStripStatusLabel();
            this.DateTimeStatusLabel = new System.Windows.Forms.ToolStripStatusLabel();
            ((System.ComponentModel.ISupportInitialize)(this.MainSplitContainer)).BeginInit();
            this.MainSplitContainer.Panel1.SuspendLayout();
            this.MainSplitContainer.SuspendLayout();
            this.StatusStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // MainSplitContainer
            // 
            this.MainSplitContainer.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.MainSplitContainer.Location = new System.Drawing.Point(2, 35);
            this.MainSplitContainer.Name = "MainSplitContainer";
            // 
            // MainSplitContainer.Panel1
            // 
            this.MainSplitContainer.Panel1.Controls.Add(this.ResourceExplorer);
            // 
            // MainSplitContainer.Panel2
            // 
            this.MainSplitContainer.Panel2.BackColor = System.Drawing.Color.White;
            this.MainSplitContainer.Panel2.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom;
            this.MainSplitContainer.Size = new System.Drawing.Size(1012, 531);
            this.MainSplitContainer.SplitterDistance = 243;
            this.MainSplitContainer.TabIndex = 0;
            // 
            // ResourceExplorer
            // 
            this.ResourceExplorer.AllowDrop = true;
            this.ResourceExplorer.Dock = System.Windows.Forms.DockStyle.Fill;
            this.ResourceExplorer.HideSelection = false;
            this.ResourceExplorer.Location = new System.Drawing.Point(0, 0);
            this.ResourceExplorer.Name = "ResourceExplorer";
            treeNode1.Name = "Computers";
            treeNode1.Text = "Computers";
            treeNode2.Name = "Collections";
            treeNode2.Text = "Collections";
            treeNode3.Name = "Task";
            treeNode3.Text = "Task";
            treeNode4.Name = "Query";
            treeNode4.Text = "Query";
            treeNode5.Name = "Results";
            treeNode5.Text = "Results";
            treeNode6.Name = "RootNode";
            treeNode6.Text = "Resource Explorer";
            treeNode7.Name = "CustomActions";
            treeNode7.Text = "Custom Actions";
            treeNode8.Name = "Settings";
            treeNode8.Text = "Settings";
            this.ResourceExplorer.Nodes.AddRange(new System.Windows.Forms.TreeNode[] {
            treeNode6,
            treeNode8});
            this.ResourceExplorer.Size = new System.Drawing.Size(243, 531);
            this.ResourceExplorer.TabIndex = 0;
            // 
            // UserInputComboBox
            // 
            this.UserInputComboBox.AllowDrop = true;
            this.UserInputComboBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.UserInputComboBox.FormattingEnabled = true;
            this.UserInputComboBox.Location = new System.Drawing.Point(107, 8);
            this.UserInputComboBox.Name = "UserInputComboBox";
            this.UserInputComboBox.Size = new System.Drawing.Size(652, 21);
            this.UserInputComboBox.TabIndex = 0;
            // 
            // LabelResource
            // 
            this.LabelResource.AutoSize = true;
            this.LabelResource.Location = new System.Drawing.Point(12, 11);
            this.LabelResource.Name = "LabelResource";
            this.LabelResource.Size = new System.Drawing.Size(89, 13);
            this.LabelResource.TabIndex = 1;
            this.LabelResource.Text = "Select Resource:";
            // 
            // SubmitButton
            // 
            this.SubmitButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.SubmitButton.Location = new System.Drawing.Point(765, 6);
            this.SubmitButton.Name = "SubmitButton";
            this.SubmitButton.Size = new System.Drawing.Size(75, 23);
            this.SubmitButton.TabIndex = 2;
            this.SubmitButton.Text = "Submit";
            this.SubmitButton.UseVisualStyleBackColor = true;
            // 
            // ClearButton
            // 
            this.ClearButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.ClearButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.ClearButton.Location = new System.Drawing.Point(846, 6);
            this.ClearButton.Name = "ClearButton";
            this.ClearButton.Size = new System.Drawing.Size(75, 23);
            this.ClearButton.TabIndex = 3;
            this.ClearButton.Text = "Clear";
            this.ClearButton.UseVisualStyleBackColor = true;
            // 
            // NewButton
            // 
            this.NewButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.NewButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.NewButton.Location = new System.Drawing.Point(927, 6);
            this.NewButton.Name = "NewButton";
            this.NewButton.Size = new System.Drawing.Size(75, 23);
            this.NewButton.TabIndex = 4;
            this.NewButton.Text = "New";
            this.NewButton.UseVisualStyleBackColor = true;
            // 
            // StatusStrip1
            // 
            this.StatusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.UpTimeStatusLabel,
            this.MemoryUsageStatusLabel,
            this.CpuUsageStatusLabel,
            this.ConnectionsStatusLabel,
            this.DateTimeStatusLabel});
            this.StatusStrip1.LayoutStyle = System.Windows.Forms.ToolStripLayoutStyle.HorizontalStackWithOverflow;
            this.StatusStrip1.Location = new System.Drawing.Point(0, 569);
            this.StatusStrip1.Name = "StatusStrip1";
            this.StatusStrip1.Size = new System.Drawing.Size(1014, 22);
            this.StatusStrip1.TabIndex = 5;
            this.StatusStrip1.Text = "StatusStrip1";
            // 
            // UpTimeStatusLabel
            // 
            this.UpTimeStatusLabel.BorderStyle = System.Windows.Forms.Border3DStyle.Etched;
            this.UpTimeStatusLabel.Name = "UpTimeStatusLabel";
            this.UpTimeStatusLabel.Size = new System.Drawing.Size(108, 17);
            this.UpTimeStatusLabel.Text = "UpTimeStatusLabel";
            // 
            // MemoryUsageStatusLabel
            // 
            this.MemoryUsageStatusLabel.Name = "MemoryUsageStatusLabel";
            this.MemoryUsageStatusLabel.Size = new System.Drawing.Size(144, 17);
            this.MemoryUsageStatusLabel.Text = "MemoryUsageStatusLabel";
            // 
            // CpuUsageStatusLabel
            // 
            this.CpuUsageStatusLabel.Name = "CpuUsageStatusLabel";
            this.CpuUsageStatusLabel.Size = new System.Drawing.Size(121, 17);
            this.CpuUsageStatusLabel.Text = "CpuUsageStatusLabel";
            // 
            // ConnectionsStatusLabel
            // 
            this.ConnectionsStatusLabel.Name = "ConnectionsStatusLabel";
            this.ConnectionsStatusLabel.Size = new System.Drawing.Size(134, 17);
            this.ConnectionsStatusLabel.Text = "ConnectionsStatusLabel";
            // 
            // DateTimeStatusLabel
            // 
            this.DateTimeStatusLabel.Name = "DateTimeStatusLabel";
            this.DateTimeStatusLabel.Size = new System.Drawing.Size(133, 17);
            this.DateTimeStatusLabel.Text = "CurrentTimeStatusLabel";
            // 
            // Form1
            // 
            this.AcceptButton = this.SubmitButton;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.ClearButton;
            this.ClientSize = new System.Drawing.Size(1014, 591);
            this.Controls.Add(this.StatusStrip1);
            this.Controls.Add(this.NewButton);
            this.Controls.Add(this.ClearButton);
            this.Controls.Add(this.SubmitButton);
            this.Controls.Add(this.LabelResource);
            this.Controls.Add(this.UserInputComboBox);
            this.Controls.Add(this.MainSplitContainer);
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "SysTool";
            this.MainSplitContainer.Panel1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.MainSplitContainer)).EndInit();
            this.MainSplitContainer.ResumeLayout(false);
            this.StatusStrip1.ResumeLayout(false);
            this.StatusStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }
        internal System.Windows.Forms.SplitContainer MainSplitContainer;
        internal System.Windows.Forms.ComboBox UserInputComboBox;
        internal System.Windows.Forms.Label LabelResource;
        internal System.Windows.Forms.Button SubmitButton;
        public System.Windows.Forms.TreeView ResourceExplorer;
        internal System.Windows.Forms.Button ClearButton;
        internal System.Windows.Forms.Button NewButton;
        internal StatusStrip StatusStrip1;
        internal ToolStripStatusLabel UpTimeStatusLabel;
        internal ToolStripStatusLabel CpuUsageStatusLabel;
        internal ToolStripStatusLabel MemoryUsageStatusLabel;
        internal ToolStripStatusLabel ConnectionsStatusLabel;
        internal ToolStripStatusLabel DateTimeStatusLabel;

        #endregion
    }
}

