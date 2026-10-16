using Bt_16_10;
using System.ComponentModel;
using System.Text;
using System.Text.RegularExpressions;
using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace Bt_16_10
{
    public partial class Form1 : Form
    {
        private BindingList<lichHen> danhSach =
            new BindingList<lichHen>();

        private BindingSource bindingSource = new BindingSource();
        private ErrorProvider errorProvider = new ErrorProvider();

        private TableLayoutPanel mainLayout;
        private TableLayoutPanel inputLayout;
        private TableLayoutPanel buttonLayout;

        private TextBox txtMaLich = new TextBox();
        private TextBox txtHoTen = new TextBox();
        private TextBox txtSoDienThoai = new TextBox();
        private TextBox txtLyDo = new TextBox();
        private TextBox txtTimKiem = new TextBox();

        private DateTimePicker dtpNgayHen = new DateTimePicker();
        private DateTimePicker dtpLocNgay = new DateTimePicker();

        private ComboBox cboBacSi = new ComboBox();
        private ComboBox cboTrangThai = new ComboBox();
        private ComboBox cboLocTrangThai = new ComboBox();

        private CheckBox chkLocNgay = new CheckBox();
        private DataGridView dgvLichHen = new DataGridView();

        private Button btnThem = new Button();
        private Button btnSua = new Button();
        private Button btnXoa = new Button();
        private Button btnLamMoi = new Button();

        private ToolStripStatusLabel lblSoLuong =
            new ToolStripStatusLabel();

        private lichHen lichDangChon;
        private bool dangNapDuLieu = false;

        public Form1()
        {
            KhoiTaoGiaoDien();
            KhoiTaoDuLieu();
            GanSuKien();
            ApDungBoLoc();
        }

        private void KhoiTaoGiaoDien()
        {
            Text = "Clinic Appointment Manager";
            StartPosition = FormStartPosition.CenterScreen;
            WindowState = FormWindowState.Maximized;
            MinimumSize = new Size(950, 650);
            Font = new Font("Segoe UI", 9);

            MenuStrip menu = new MenuStrip();
            ToolStripMenuItem menuTep =
                new ToolStripMenuItem("Tệp");
            ToolStripMenuItem menuXuat =
                new ToolStripMenuItem("Xuất danh sách CSV");

            menuXuat.Click += MenuXuat_Click;
            menuTep.DropDownItems.Add(menuXuat);
            menu.Items.Add(menuTep);
            MainMenuStrip = menu;

            StatusStrip status = new StatusStrip();
            lblSoLuong.Text = "Số lịch hẹn: 0";
            status.Items.Add(lblSoLuong);

            mainLayout = new TableLayoutPanel();
            mainLayout.Dock = DockStyle.Fill;
            mainLayout.ColumnCount = 1;
            mainLayout.RowCount = 2;
            mainLayout.Padding = new Padding(8);

            mainLayout.ColumnStyles.Add(
                new ColumnStyle(SizeType.Percent, 100));
            mainLayout.RowStyles.Add(
                new RowStyle(SizeType.Absolute, 245));
            mainLayout.RowStyles.Add(
                new RowStyle(SizeType.Percent, 100));

            Controls.Add(mainLayout);
            Controls.Add(status);
            Controls.Add(menu);

            menu.Dock = DockStyle.Top;
            status.Dock = DockStyle.Bottom;

            // Khu nhập liệu: 5 hàng, 4 cột
            inputLayout = new TableLayoutPanel();
            inputLayout.Dock = DockStyle.Fill;
            inputLayout.ColumnCount = 4;
            inputLayout.RowCount = 5;
            inputLayout.Padding = new Padding(5);

            inputLayout.ColumnStyles.Add(
                new ColumnStyle(SizeType.Percent, 15));
            inputLayout.ColumnStyles.Add(
                new ColumnStyle(SizeType.Percent, 35));
            inputLayout.ColumnStyles.Add(
                new ColumnStyle(SizeType.Percent, 15));
            inputLayout.ColumnStyles.Add(
                new ColumnStyle(SizeType.Percent, 35));

            for (int i = 0; i < 5; i++)
            {
                inputLayout.RowStyles.Add(
                    new RowStyle(SizeType.Percent, 20));
            }

            mainLayout.Controls.Add(inputLayout, 0, 0);

            txtMaLich.Dock = DockStyle.Fill;
            txtHoTen.Dock = DockStyle.Fill;
            txtSoDienThoai.Dock = DockStyle.Fill;
            txtLyDo.Dock = DockStyle.Fill;

            dtpNgayHen.Dock = DockStyle.Fill;
            dtpNgayHen.Format = DateTimePickerFormat.Custom;
            dtpNgayHen.CustomFormat = "dd/MM/yyyy HH:mm";
            dtpNgayHen.ShowUpDown = true;
            dtpNgayHen.Value = DateTime.Now.AddMinutes(30);

            cboBacSi.Dock = DockStyle.Fill;
            cboBacSi.DropDownStyle = ComboBoxStyle.DropDownList;
            cboBacSi.Items.AddRange(new object[]
            {
                "BS Nguyễn Văn An",
                "BS Trần Thị Bình",
                "BS Lê Minh Cường"
            });
            cboBacSi.SelectedIndex = 0;

            cboTrangThai.Dock = DockStyle.Fill;
            cboTrangThai.DropDownStyle = ComboBoxStyle.DropDownList;
            cboTrangThai.Items.AddRange(new object[]
            {
                "Chờ khám",
                "Đang khám",
                "Hoàn tất",
                "Đã hủy"
            });
            cboTrangThai.SelectedIndex = 0;

            ThemTruong("Mã lịch hẹn:", txtMaLich, 0, 0);
            ThemTruong("Họ tên:", txtHoTen, 2, 0);
            ThemTruong("Số điện thoại:", txtSoDienThoai, 0, 1);
            ThemTruong("Ngày giờ hẹn:", dtpNgayHen, 2, 1);
            ThemTruong("Bác sĩ:", cboBacSi, 0, 2);
            ThemTruong("Trạng thái:", cboTrangThai, 2, 2);
            ThemTruong("Lý do khám:", txtLyDo, 0, 3);

            buttonLayout = new TableLayoutPanel();
            buttonLayout.Dock = DockStyle.Fill;
            buttonLayout.ColumnCount = 4;
            buttonLayout.RowCount = 1;

            for (int i = 0; i < 4; i++)
            {
                buttonLayout.ColumnStyles.Add(
                    new ColumnStyle(SizeType.Percent, 25));
            }

            btnThem.Text = "Thêm";
            btnSua.Text = "Cập nhật";
            btnXoa.Text = "Xóa";
            btnLamMoi.Text = "Làm mới";

            Button[] buttons =
            {
                btnThem, btnSua, btnXoa, btnLamMoi
            };

            for (int i = 0; i < buttons.Length; i++)
            {
                buttons[i].Dock = DockStyle.Fill;
                buttons[i].Margin = new Padding(4);
                buttonLayout.Controls.Add(buttons[i], i, 0);
            }

            inputLayout.Controls.Add(buttonLayout, 1, 4);
            inputLayout.SetColumnSpan(buttonLayout, 3);

            // Khu tìm kiếm, lọc và bảng
            TableLayoutPanel tableLayout = new TableLayoutPanel();
            tableLayout.Dock = DockStyle.Fill;
            tableLayout.ColumnCount = 1;
            tableLayout.RowCount = 2;

            tableLayout.RowStyles.Add(
                new RowStyle(SizeType.Absolute, 45));
            tableLayout.RowStyles.Add(
                new RowStyle(SizeType.Percent, 100));

            mainLayout.Controls.Add(tableLayout, 0, 1);

            FlowLayoutPanel filterPanel = new FlowLayoutPanel();
            filterPanel.Dock = DockStyle.Fill;
            filterPanel.WrapContents = false;
            filterPanel.AutoScroll = true;

            Label lblTimKiem = new Label();
            lblTimKiem.Text = "Tìm kiếm:";
            lblTimKiem.AutoSize = true;
            lblTimKiem.Margin = new Padding(4, 10, 3, 0);

            txtTimKiem.Width = 180;
            txtTimKiem.Margin = new Padding(3, 6, 3, 0);

            chkLocNgay.Text = "Lọc ngày:";
            chkLocNgay.AutoSize = true;
            chkLocNgay.Margin = new Padding(8, 9, 3, 0);

            dtpLocNgay.Format = DateTimePickerFormat.Short;
            dtpLocNgay.Width = 110;
            dtpLocNgay.Margin = new Padding(3, 6, 3, 0);

            Label lblTrangThai = new Label();
            lblTrangThai.Text = "Trạng thái:";
            lblTrangThai.AutoSize = true;
            lblTrangThai.Margin = new Padding(8, 10, 3, 0);

            cboLocTrangThai.DropDownStyle =
                ComboBoxStyle.DropDownList;
            cboLocTrangThai.Width = 120;
            cboLocTrangThai.Margin = new Padding(3, 6, 3, 0);

            cboLocTrangThai.Items.Add("Tất cả");
            cboLocTrangThai.Items.AddRange(new object[]
            {
                "Chờ khám",
                "Đang khám",
                "Hoàn tất",
                "Đã hủy"
            });
            cboLocTrangThai.SelectedIndex = 0;

            filterPanel.Controls.Add(lblTimKiem);
            filterPanel.Controls.Add(txtTimKiem);
            filterPanel.Controls.Add(chkLocNgay);
            filterPanel.Controls.Add(dtpLocNgay);
            filterPanel.Controls.Add(lblTrangThai);
            filterPanel.Controls.Add(cboLocTrangThai);

            tableLayout.Controls.Add(filterPanel, 0, 0);

            dgvLichHen.Dock = DockStyle.Fill;
            dgvLichHen.AutoGenerateColumns = false;
            dgvLichHen.SelectionMode =
                DataGridViewSelectionMode.FullRowSelect;
            dgvLichHen.MultiSelect = false;
            dgvLichHen.ReadOnly = true;
            dgvLichHen.AllowUserToAddRows = false;
            dgvLichHen.AllowUserToDeleteRows = false;
            dgvLichHen.RowHeadersVisible = false;
            dgvLichHen.AutoSizeColumnsMode =
                DataGridViewAutoSizeColumnsMode.Fill;

            ThemCot("MaLich", "Mã lịch", 70);
            ThemCot("HoTen", "Họ tên bệnh nhân", 140);
            ThemCot("SoDienThoai", "Số điện thoại", 100);
            ThemCot("NgayHen", "Ngày giờ hẹn", 130);
            ThemCot("BacSi", "Bác sĩ", 120);
            ThemCot("LyDo", "Lý do khám", 150);
            ThemCot("TrangThai", "Trạng thái", 90);

            dgvLichHen.Columns["NgayHen"]
                .DefaultCellStyle.Format = "dd/MM/yyyy HH:mm";

            tableLayout.Controls.Add(dgvLichHen, 0, 1);

            errorProvider.ContainerControl = this;
        }

        private void ThemTruong(
            string ten, Control dieuKhien, int cot, int hang)
        {
            Label lbl = new Label();
            lbl.Text = ten;
            lbl.Dock = DockStyle.Fill;
            lbl.TextAlign = ContentAlignment.MiddleLeft;

            inputLayout.Controls.Add(lbl, cot, hang);
            inputLayout.Controls.Add(dieuKhien, cot + 1, hang);
        }

        private void ThemCot(
            string tenThuocTinh, string tieuDe, float tiLe)
        {
            DataGridViewTextBoxColumn cot =
                new DataGridViewTextBoxColumn();

            cot.Name = tenThuocTinh;
            cot.DataPropertyName = tenThuocTinh;
            cot.HeaderText = tieuDe;
            cot.FillWeight = tiLe;

            dgvLichHen.Columns.Add(cot);
        }

        private void KhoiTaoDuLieu()
        {
            bindingSource.DataSource = danhSach;
            dgvLichHen.DataSource = bindingSource;
        }

        private void GanSuKien()
        {
            btnThem.Click += BtnThem_Click;
            btnSua.Click += BtnSua_Click;
            btnXoa.Click += BtnXoa_Click;
            btnLamMoi.Click += BtnLamMoi_Click;

            dgvLichHen.SelectionChanged +=
                DgvLichHen_SelectionChanged;

            txtTimKiem.TextChanged +=
                delegate { ApDungBoLoc(); };

            chkLocNgay.CheckedChanged +=
                delegate { ApDungBoLoc(); };

            dtpLocNgay.ValueChanged +=
                delegate { ApDungBoLoc(); };

            cboLocTrangThai.SelectedIndexChanged +=
                delegate { ApDungBoLoc(); };
        }

        private lichHen LayDuLieuTuForm()
        {
            return new lichHen
            {
                MaLich = txtMaLich.Text.Trim(),
                HoTen = txtHoTen.Text.Trim(),
                SoDienThoai = txtSoDienThoai.Text.Trim(),
                NgayHen = dtpNgayHen.Value,
                BacSi = cboBacSi.Text,
                LyDo = txtLyDo.Text.Trim(),
                TrangThai = cboTrangThai.Text
            };
        }

        private bool KiemTraDuLieu(bool dangCapNhat)
        {
            errorProvider.Clear();
            bool hopLe = true;

            string ma = txtMaLich.Text.Trim();

            if (ma == "")
            {
                errorProvider.SetError(
                    txtMaLich, "Vui lòng nhập mã lịch hẹn.");
                hopLe = false;
            }
            else if (danhSach.Any(x =>
                x.MaLich.Equals(
                    ma, StringComparison.OrdinalIgnoreCase)
                && (!dangCapNhat || x != lichDangChon)))
            {
                errorProvider.SetError(
                    txtMaLich, "Mã lịch hẹn đã tồn tại.");
                hopLe = false;
            }

            if (txtHoTen.Text.Trim() == "")
            {
                errorProvider.SetError(
                    txtHoTen, "Vui lòng nhập họ tên.");
                hopLe = false;
            }

            if (!Regex.IsMatch(
                txtSoDienThoai.Text.Trim(), @"^0\d{9}$"))
            {
                errorProvider.SetError(
                    txtSoDienThoai,
                    "Số điện thoại phải có 10 chữ số, bắt đầu bằng 0.");
                hopLe = false;
            }

            if (txtLyDo.Text.Trim() == "")
            {
                errorProvider.SetError(
                    txtLyDo, "Vui lòng nhập lý do khám.");
                hopLe = false;
            }

            if (dtpNgayHen.Value < DateTime.Now)
            {
                errorProvider.SetError(
                    dtpNgayHen,
                    "Ngày giờ hẹn không được nằm trong quá khứ.");
                hopLe = false;
            }

            if (cboBacSi.SelectedIndex < 0)
            {
                errorProvider.SetError(
                    cboBacSi, "Vui lòng chọn bác sĩ.");
                hopLe = false;
            }

            if (cboTrangThai.SelectedIndex < 0)
            {
                errorProvider.SetError(
                    cboTrangThai, "Vui lòng chọn trạng thái.");
                hopLe = false;
            }

            return hopLe;
        }

        private void BtnThem_Click(object sender, EventArgs e)
        {
            if (!KiemTraDuLieu(false))
                return;

            if (MessageBox.Show(
                "Bạn có muốn thêm lịch hẹn này không?",
                "Xác nhận",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question) != DialogResult.Yes)
                return;

            danhSach.Add(LayDuLieuTuForm());

            LamMoiForm();
            ApDungBoLoc();
        }

        private void BtnSua_Click(object sender, EventArgs e)
        {
            if (lichDangChon == null ||
                !danhSach.Contains(lichDangChon))
            {
                MessageBox.Show("Bạn hãy chọn lịch hẹn cần cập nhật.");
                return;
            }

            if (!KiemTraDuLieu(true))
                return;

            if (MessageBox.Show(
                "Bạn có muốn cập nhật lịch hẹn này không?",
                "Xác nhận",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question) != DialogResult.Yes)
                return;

            lichHen moi = LayDuLieuTuForm();

            lichDangChon.MaLich = moi.MaLich;
            lichDangChon.HoTen = moi.HoTen;
            lichDangChon.SoDienThoai = moi.SoDienThoai;
            lichDangChon.NgayHen = moi.NgayHen;
            lichDangChon.BacSi = moi.BacSi;
            lichDangChon.LyDo = moi.LyDo;
            lichDangChon.TrangThai = moi.TrangThai;

            LamMoiForm();
            ApDungBoLoc();
        }

        private void BtnXoa_Click(object sender, EventArgs e)
        {
            if (lichDangChon == null ||
                !danhSach.Contains(lichDangChon))
            {
                MessageBox.Show("Bạn hãy chọn lịch hẹn cần xóa.");
                return;
            }

            if (MessageBox.Show(
                "Bạn có chắc chắn muốn xóa lịch hẹn " +
                lichDangChon.MaLich + " không?",
                "Xác nhận xóa",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning) != DialogResult.Yes)
                return;

            danhSach.Remove(lichDangChon);

            LamMoiForm();
            ApDungBoLoc();
        }

        private void DgvLichHen_SelectionChanged(
            object sender, EventArgs e)
        {
            if (dangNapDuLieu)
                return;

            lichHen lh = bindingSource.Current as lichHen;

            if (lh == null)
                return;

            lichDangChon = lh;
            dangNapDuLieu = true;

            txtMaLich.Text = lh.MaLich;
            txtHoTen.Text = lh.HoTen;
            txtSoDienThoai.Text = lh.SoDienThoai;
            dtpNgayHen.Value = lh.NgayHen;
            cboBacSi.SelectedItem = lh.BacSi;
            txtLyDo.Text = lh.LyDo;
            cboTrangThai.SelectedItem = lh.TrangThai;

            dangNapDuLieu = false;
            errorProvider.Clear();
        }

        private void ApDungBoLoc()
        {
            string tuKhoa = txtTimKiem.Text.Trim()
                .ToLowerInvariant();

            string trangThai = cboLocTrangThai.Text;

            var ketQua = danhSach.Where(x =>
                (x.HoTen.ToLowerInvariant().Contains(tuKhoa)
                 || x.SoDienThoai.Contains(tuKhoa))
                &&
                (!chkLocNgay.Checked
                 || x.NgayHen.Date == dtpLocNgay.Value.Date)
                &&
                (trangThai == "Tất cả"
                 || string.IsNullOrEmpty(trangThai)
                 || x.TrangThai == trangThai)
            ).ToList();

            bindingSource.DataSource =
                new BindingList<lichHen>(ketQua);

            dgvLichHen.DataSource = bindingSource;

            lblSoLuong.Text = "Số lịch hẹn: " + ketQua.Count;
        }

        private void BtnLamMoi_Click(object sender, EventArgs e)
        {
            LamMoiForm();
        }

        private void LamMoiForm()
        {
            dangNapDuLieu = true;

            lichDangChon = null;
            errorProvider.Clear();

            txtMaLich.Clear();
            txtHoTen.Clear();
            txtSoDienThoai.Clear();
            txtLyDo.Clear();

            dtpNgayHen.Value = DateTime.Now.AddMinutes(30);

            if (cboBacSi.Items.Count > 0)
                cboBacSi.SelectedIndex = 0;

            if (cboTrangThai.Items.Count > 0)
                cboTrangThai.SelectedIndex = 0;

            dgvLichHen.ClearSelection();

            dangNapDuLieu = false;
        }

        private void MenuXuat_Click(object sender, EventArgs e)
        {
            SaveFileDialog save = new SaveFileDialog();
            save.Title = "Xuất danh sách lịch hẹn";
            save.Filter = "CSV (*.csv)|*.csv";
            save.FileName = "DanhSachLichHen.csv";

            if (save.ShowDialog() != DialogResult.OK)
                return;

            var ketQua = bindingSource.List
                .Cast<lichHen>().ToList();

            using (StreamWriter writer = new StreamWriter(
                save.FileName, false, Encoding.UTF8))
            {
                writer.WriteLine(
                    "MaLich,HoTen,SoDienThoai,NgayHen,BacSi,LyDo,TrangThai");

                foreach (lichHen lh in ketQua)
                {
                    writer.WriteLine(
                        Csv(lh.MaLich) + "," +
                        Csv(lh.HoTen) + "," +
                        Csv(lh.SoDienThoai) + "," +
                        Csv(lh.NgayHen.ToString("dd/MM/yyyy HH:mm")) + "," +
                        Csv(lh.BacSi) + "," +
                        Csv(lh.LyDo) + "," +
                        Csv(lh.TrangThai));
                }
            }

            MessageBox.Show(
                "Xuất CSV thành công!",
                "Thông báo",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }

        private string Csv(string giaTri)
        {
            return "\"" +
                (giaTri ?? "").Replace("\"", "\"\"") +
                "\"";
        }
    }
}
