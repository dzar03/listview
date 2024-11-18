using System;
using System.IO;
using System.Windows.Forms;
using System.Collections;
using System.Reflection;
using OfficeOpenXml;

namespace WindowsFormsApp2

{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();


            // Thiết lập chế độ hiển thị
            listView1.View = View.Details;
            listView1.FullRowSelect = true;
            listView1.GridLines = true;
            listView1.CheckBoxes = true;
            listView1.MultiSelect = true;
            listView1.Sorting = SortOrder.Ascending;
            listView1.ColumnClick += new ColumnClickEventHandler(listView1_ColumnClick);
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;


        }


        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.Windows.Forms.ListViewItem listViewItem1 = new System.Windows.Forms.ListViewItem(new string[] {
            "CS2",
            "1/1/2024",
            "30 GB"}, 1);
            System.Windows.Forms.ListViewItem listViewItem2 = new System.Windows.Forms.ListViewItem(new string[] {
            "Valorant",
            "2/2/2024",
            "50 GB"}, 2);
            System.Windows.Forms.ListViewItem listViewItem3 = new System.Windows.Forms.ListViewItem(new string[] {
            "LOL",
            "3/3/2024",
            "60 GB"}, 0);
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.listView1 = new System.Windows.Forms.ListView();
            this.column1 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.column2 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.column3 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.imageList1 = new System.Windows.Forms.ImageList(this.components);
            this.button1 = new System.Windows.Forms.Button();
            this.txtApp = new System.Windows.Forms.TextBox();
            this.button2 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.txtDate = new System.Windows.Forms.TextBox();
            this.txtSize = new System.Windows.Forms.TextBox();
            this.button4 = new System.Windows.Forms.Button();
            this.button5 = new System.Windows.Forms.Button();
            this.txtIndex = new System.Windows.Forms.TextBox();
            this.button6 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // listView1
            // 
            this.listView1.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.column1,
            this.column2,
            this.column3});
            this.listView1.HideSelection = false;
            this.listView1.Items.AddRange(new System.Windows.Forms.ListViewItem[] {
            listViewItem1,
            listViewItem2,
            listViewItem3});
            this.listView1.LargeImageList = this.imageList1;
            this.listView1.Location = new System.Drawing.Point(81, 12);
            this.listView1.Name = "listView1";
            this.listView1.Size = new System.Drawing.Size(890, 333);
            this.listView1.SmallImageList = this.imageList1;
            this.listView1.TabIndex = 0;
            this.listView1.UseCompatibleStateImageBehavior = false;
            this.listView1.View = System.Windows.Forms.View.Details;
            this.listView1.SelectedIndexChanged += new System.EventHandler(this.listView1_SelectedIndexChanged_1);
            // 
            // column1
            // 
            this.column1.Text = "App";
            this.column1.Width = 98;
            // 
            // column2
            // 
            this.column2.Text = "Date";
            this.column2.Width = 100;
            // 
            // column3
            // 
            this.column3.Text = "Size";
            // 
            // imageList1
            // 
            this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
            this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
            this.imageList1.Images.SetKeyName(0, "lol-logo.png");
            this.imageList1.Images.SetKeyName(1, "st,small,507x507-pad,600x600,f8f8f8.jpg");
            this.imageList1.Images.SetKeyName(2, "valorant.png");
            this.imageList1.Images.SetKeyName(3, "excel.png");
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(364, 54);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(90, 23);
            this.button1.TabIndex = 1;
            this.button1.Text = "Them item";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // txtApp
            // 
            this.txtApp.Location = new System.Drawing.Point(364, 174);
            this.txtApp.Name = "txtApp";
            this.txtApp.Size = new System.Drawing.Size(144, 22);
            this.txtApp.TabIndex = 2;
            this.txtApp.TextChanged += new System.EventHandler(this.textBox1_TextChanged);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(542, 54);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(75, 23);
            this.button2.TabIndex = 5;
            this.button2.Text = "Xoa Item";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(364, 104);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(90, 23);
            this.button3.TabIndex = 6;
            this.button3.Text = "Sua Item";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // txtDate
            // 
            this.txtDate.Location = new System.Drawing.Point(364, 218);
            this.txtDate.Name = "txtDate";
            this.txtDate.Size = new System.Drawing.Size(144, 22);
            this.txtDate.TabIndex = 7;
            // 
            // txtSize
            // 
            this.txtSize.Location = new System.Drawing.Point(364, 258);
            this.txtSize.Name = "txtSize";
            this.txtSize.Size = new System.Drawing.Size(144, 22);
            this.txtSize.TabIndex = 8;
            this.txtSize.TextChanged += new System.EventHandler(this.textBox3_TextChanged);
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(532, 93);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(95, 23);
            this.button4.TabIndex = 9;
            this.button4.Text = "Xoa Tat Ca";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // button5
            // 
            this.button5.Location = new System.Drawing.Point(517, 133);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(110, 23);
            this.button5.TabIndex = 10;
            this.button5.Text = "Xoa Tai Vi Tri";
            this.button5.UseVisualStyleBackColor = true;
            this.button5.Click += new System.EventHandler(this.button5_Click);
            // 
            // txtIndex
            // 
            this.txtIndex.Location = new System.Drawing.Point(648, 134);
            this.txtIndex.Name = "txtIndex";
            this.txtIndex.Size = new System.Drawing.Size(41, 22);
            this.txtIndex.TabIndex = 11;
            this.txtIndex.TextChanged += new System.EventHandler(this.txtIndex_TextChanged);
            // 
            // button6
            // 
            this.button6.ImageIndex = 3;
            this.button6.ImageList = this.imageList1;
            this.button6.Location = new System.Drawing.Point(460, 54);
            this.button6.Name = "button6";
            this.button6.Size = new System.Drawing.Size(29, 23);
            this.button6.TabIndex = 12;
            this.button6.UseVisualStyleBackColor = true;
            this.button6.Click += new System.EventHandler(this.button6_Click_1);
            // 
            // Form1
            // 
            this.ClientSize = new System.Drawing.Size(1261, 538);
            this.Controls.Add(this.button6);
            this.Controls.Add(this.txtIndex);
            this.Controls.Add(this.button5);
            this.Controls.Add(this.button4);
            this.Controls.Add(this.txtSize);
            this.Controls.Add(this.txtDate);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.txtApp);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.listView1);
            this.Name = "Form1";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        private ListView listView1;



        private ColumnHeader column1;
        private ColumnHeader column2;
        private ImageList imageList1;
        private System.ComponentModel.IContainer components;
        private ColumnHeader column3;

        private void AddItemAtSelectedIndex(string app, string date, string size)
        {
            DateTime parsedDate;

            // Kiểm tra xem dữ liệu ngày tháng có hợp lệ không
            if (!DateTime.TryParse(date, out parsedDate))
            {
                // Nếu không hợp lệ, hiển thị thông báo và không thêm dữ liệu
                MessageBox.Show("Ngày tháng không hợp lệ! Vui lòng nhập đúng định dạng.", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtDate.Focus(); // Đặt lại focus vào ô nhập ngày
                return; // Thoát khỏi hàm mà không thêm dữ liệu
            }
            ListViewItem item = new ListViewItem(app);
            item.SubItems.Add(parsedDate.ToString("dd/MM/yyyy")); // Định dạng ngày tháng theo chuẩn
            item.SubItems.Add(size);

            if (listView1.SelectedIndices.Count > 0)  // Kiểm tra nếu có mục nào được chọn
            {
                int selectedIndex = listView1.SelectedIndices[0]; // Lấy vị trí mục được chọn
                listView1.Items.Insert(selectedIndex, item); // Thêm vào vị trí đã chọn
            }
            else
            {
                listView1.Items.Add(item); // Nếu không có mục nào được chọn, thêm vào cuối danh sách
            }
        }
        private void ClearTextBoxes()
        {
            txtApp.Clear();  // Nếu bạn có TextBox tên là txtApp
            txtDate.Clear(); // Nếu bạn có TextBox tên là txtDate
            txtSize.Clear(); // Nếu bạn có TextBox tên là txtSize
        }

        private Button button1;
        private void button1_Click(object sender, EventArgs e)
        {
            AddItemAtSelectedIndex(txtApp.Text, txtDate.Text, txtSize.Text);
            ClearTextBoxes();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // Đặt thuộc tính MultiSelect của ListView thành true
            listView1.MultiSelect = true;
        }


        private void btnSelectAll_Click(object sender, EventArgs e)
        {
            foreach (ListViewItem item in listView1.Items)
            {
                item.Checked = true;
            }
        }
        
           
           private void RemoveSelectedItem()
           {
               if (listView1.SelectedItems.Count > 0)
               {
                   // Duyệt qua các mục đã chọn và xóa chúng
                  foreach (ListViewItem item in listView1.SelectedItems)

                   {
                       listView1.Items.Remove(item);
                   }
               } 
               else
               {
                   MessageBox.Show("Không có mục nào được chọn để xóa!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
               }
           } 

        private void RemoveItemByIndex(int index)
        {
            if (index >= 0 && index < listView1.Items.Count)
            {
                listView1.Items.RemoveAt(index); // Xóa item tại vị trí chỉ định
            }


        }
        private void UpdateSelectedItem()
        {
            if (listView1.SelectedItems.Count > 0) // Kiểm tra nếu có mục nào được chọn
            {
                ListViewItem selectedItem = listView1.SelectedItems[0];
                selectedItem.Text = txtApp.Text;              // Cập nhật cột đầu tiên
                selectedItem.SubItems[1].Text = txtDate.Text; // Cập nhật cột thứ hai
                selectedItem.SubItems[2].Text = txtSize.Text; // Cập nhật cột thứ ba
            }
            else
            {
                MessageBox.Show("Vui lòng chọn mục để cập nhật.");
            }
        }

        private TextBox txtApp;

        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {
            // Kiểm tra xem có mục nào được chọn không
            if (listView1.SelectedItems.Count > 0)
            {
                // Lấy mục đang được chọn
                ListViewItem selectedItem = listView1.SelectedItems[0];

                // Cập nhật TextBox với thông tin từ mục được chọn
                txtApp.Text = selectedItem.Text;              // Cột App
                txtApp.Text = selectedItem.SubItems[1].Text; // Cột Date
                txtApp.Text = selectedItem.SubItems[2].Text; // Cột Size
            }
            else
            {
                // Xóa dữ liệu trong TextBox nếu không có mục nào được chọn
                txtApp.Clear();
                txtApp.Clear();
                txtApp.Clear();
            }

        }

        private bool isAscending = true;



        private void listView1_ColumnClick(object sender, ColumnClickEventArgs e)
        {
            // Sắp xếp theo cột được nhấp
            listView1.ListViewItemSorter = new ListViewItemComparer(e.Column, isAscending);
            isAscending = !isAscending; // Đảo ngược thứ tự sắp xếp
            listView1.Sort(); // Sắp xếp ListView
        }

        private void listView1_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            if (listView1.SelectedItems.Count > 0)
            {
                ListViewItem selectedItem = listView1.SelectedItems[0];

                // Hiển thị thông tin của item trong TextBox để chỉnh sửa
                txtApp.Text = selectedItem.Text;
                txtDate.Text = selectedItem.SubItems[1].Text;
                txtSize.Text = selectedItem.SubItems[2].Text;
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private Button button2;
        private Button button3;
        private TextBox txtDate;
        private TextBox txtSize;

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            RemoveSelectedItem();
        }

        private Button button4;

        private void button4_Click(object sender, EventArgs e)
        {
            listView1.Items.Clear();
        }

        private Button button5;

        private void button5_Click(object sender, EventArgs e)
        {
            // Kiểm tra nếu txtIndex chứa giá trị hợp lệ
            if (int.TryParse(txtIndex.Text, out int index))
            {
                RemoveItemByIndex(index);
            }
            else
            {
                MessageBox.Show("Vui lòng nhập chỉ số hợp lệ.");
            }
        }

        private TextBox txtIndex;

        private void txtIndex_TextChanged(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            UpdateSelectedItem();
            ClearTextBoxes();
        }

        private void txtApp2_TextChanged(object sender, EventArgs e)
        {

        }

        private void button6_Click(object sender, EventArgs e)
        {
            if (listView1.SelectedItems.Count > 0)
            {
                // Lấy item đã chọn
                ListViewItem selectedItem = listView1.SelectedItems[0];

                // Cập nhật lại các giá trị của item với thông tin từ TextBox
                selectedItem.Text = txtApp.Text;
                selectedItem.SubItems[1].Text = txtDate.Text;
                selectedItem.SubItems[2].Text = txtSize.Text;

                // Xóa nội dung TextBox sau khi cập nhật
                txtApp.Clear();
                txtDate.Clear();
                txtSize.Clear();
            }
            else
            {
                MessageBox.Show("Vui lòng chọn một mục để sửa.");
            }
        }

        private Button button6;



               private void button6_Click_1(object sender, EventArgs e)
                {
                    OpenFileDialog openFileDialog = new OpenFileDialog
                    {
                        Filter = "Excel Files|*.xlsx;*.xls"
                    };

                    if (openFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        string filePath = openFileDialog.FileName;
                        LoadExcelData(filePath);
                    }
                }

        private void LoadExcelData(string filePath)
        {
            // Kiểm tra xem file có tồn tại không
            if (!File.Exists(filePath))
            {
                MessageBox.Show("File không tồn tại.", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                // Mở file Excel và đọc dữ liệu
                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    // Lấy Sheet đầu tiên trong file Excel
                    var worksheet = package.Workbook.Worksheets[0];

                    // Nếu chưa có cột nào trong ListView, thêm tiêu đề từ dòng đầu tiên của Excel
                    if (listView1.Columns.Count == 0)
                    {
                        for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                        {
                            string columnName = worksheet.Cells[1, col].Text;
                            listView1.Columns.Add(columnName);
                        }
                    }

                    // Đọc từng dòng dữ liệu từ Excel và thêm vào ListView
                    for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                    {
                        ListViewItem item = new ListViewItem(worksheet.Cells[row, 1].Text);
                        for (int col = 2; col <= worksheet.Dimension.End.Column; col++)
                        {
                            item.SubItems.Add(worksheet.Cells[row, col].Text);
                        }
                        listView1.Items.Add(item); // Thêm dữ liệu mới vào cuối danh sách
                    }

                    // Căn chỉnh các cột cho vừa với nội dung
                    listView1.AutoResizeColumns(ColumnHeaderAutoResizeStyle.ColumnContent);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Đã xảy ra lỗi khi tải dữ liệu từ file Excel: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


    }

    // Lớp để so sánh các mục trong ListView
    public class ListViewItemComparer : IComparer
    {
        private int col;
        private bool ascending;

        public ListViewItemComparer(int column, bool isAscending)
        {
            col = column;
            ascending = isAscending;
        }

        public int Compare(object x, object y)
        {
            int returnVal = String.Compare(((ListViewItem)x).SubItems[col].Text, ((ListViewItem)y).SubItems[col].Text);
            return ascending ? returnVal : -returnVal;
        }

}





}




