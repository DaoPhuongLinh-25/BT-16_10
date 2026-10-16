using Bài_16_10__1_;
using System.ComponentModel;
using System.Text;
using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace Bài_16_10__1_
{
    public partial class Form1 : Form
    {
        private BindingList<PhieuMuon> danhSach =
            new BindingList<PhieuMuon>();

        private BindingSource bindingSource = new BindingSource();
        private ErrorProvider errorProvider = new ErrorProvider();

        private TableLayoutPanel mainLayout;
        private TableLayoutPanel inputLayout;

        private TextBox txtMaPhieu = new TextBox();
        private TextBox txtMaSV = new TextBox();
        private TextBox txtHoTen = new TextBox();
        private TextBox txtSoLuong = new TextBox();
        private TextBox txtTimKiem = new TextBox();

        private ComboBox cboThietBi = new ComboBox();
        private ComboBox cboTrangThai = new ComboBox();
        private ComboBox cboLocTrangThai = new ComboBox();

        private DateTimePicker dtpNgayMuon = new DateTimePicker();
        private DateTimePicker dtpHanTra = new DateTimePicker();
        private DateTimePicker dtpNgayTra = new DateTimePicker();

        private CheckBox chkNgayTra = new CheckBox();
        private DataGridView dgvPhieuMuon = new DataGridView();

        private Button btnThem = new Button();
        private Button btnSua = new Button();
        private Button btnXoa = new Button();
        private Button btnLamMoi = new Button();

        private ToolStripStatusLabel lblTong =
            new ToolStripStatusLabel();

        private ToolStripStatusLabel lblQuaHan =
            new ToolStripStatusLabel();

        private PhieuMuon phieuDangChon;
        private bool dangNapDuLieu = false;

        public Form1()
        {
            KhoiTaoGiaoDien();
            GanSuKien();
            CapNhatDanhSach();
        }

        private void KhoiTaoGiaoDien()
        {
            Text = "Lab Equipment Checkout";
            StartPosition = FormStartPosition.CenterScreen;
            WindowState = FormWindowState.Maximized;
            MinimumSize = new Size(1000, 650);
            Font = new Font("Segoe UI", 9);

            // Menu
            MenuStrip menu = new MenuStrip();
            ToolStripMenuItem menuTep =
                new ToolStripMenuItem("Tệp");

            ToolStripMenuItem menuXuat =
                new ToolStripMenuItem("Xuất CSV");

            menuXuat.Click += MenuXuat_Click;
            menuTep.DropDownItems.Add(menuXuat);
            menu.Items.Add(menuTep);
            MainMenuStrip = menu;

            // Thanh trạng thái
            StatusStrip status = new StatusStrip();

            lblTong.Text = "Tổng số phiếu: 0";
            lblQuaHan.Text = "Số phiếu quá hạn: 0";

            status.Items.Add(lblTong);
            status.Items.Add(new ToolStripStatusLabel("     "));
            status.Items.Add(lblQuaHan);

            // Bố cục chính
            mainLayout = new TableLayoutPanel();
            mainLayout.Dock = DockStyle.Fill;
            mainLayout.ColumnCount = 1;
            mainLayout.RowCount = 2;
            mainLayout.Padding = new Padding(8);

            mainLayout.ColumnStyles.Add(
                new ColumnStyle(SizeType.Percent, 100));

            mainLayout.RowStyles.Add(
                new RowStyle(SizeType.Absolute, 260));

            mainLayout.RowStyles.Add(
                new RowStyle(SizeType.Percent, 100));

            Controls.Add(mainLayout);
            Controls.Add(status);
            Controls.Add(menu);

            menu.Dock = DockStyle.Top;
            status.Dock = DockStyle.Bottom;

            // Khu nhập phiếu
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

            txtMaPhieu.Dock = DockStyle.Fill;
            txtMaSV.Dock = DockStyle.Fill;
            txtHoTen.Dock = DockStyle.Fill;
            txtSoLuong.Dock = DockStyle.Fill;

            dtpNgayMuon.Dock = DockStyle.Fill;
            dtpNgayMuon.Format = DateTimePickerFormat.Short;
            dtpNgayMuon.Value = DateTime.Today;

            dtpHanTra.Dock = DockStyle.Fill;
            dtpHanTra.Format = DateTimePickerFormat.Short;
            dtpHanTra.Value = DateTime.Today.AddDays(7);

            dtpNgayTra.Dock = DockStyle.Fill;
            dtpNgayTra.Format = DateTimePickerFormat.Short;
            dtpNgayTra.Value = DateTime.Today;
            dtpNgayTra.ShowCheckBox = true;
            dtpNgayTra.Checked = false;

            cboThietBi.Dock = DockStyle.Fill;
            cboThietBi.DropDownStyle = ComboBoxStyle.DropDownList;
            cboThietBi.Items.AddRange(new object[]
            {
                "Máy tính",
                "Laptop",
                "Chuột máy tính",
                "Bàn phím",
                "Màn hình",
                "Máy chiếu",
                "Cáp HDMI",
                "Bộ dụng cụ thực hành"
            });
            cboThietBi.SelectedIndex = 0;

            cboTrangThai.Dock = DockStyle.Fill;
            cboTrangThai.DropDownStyle = ComboBoxStyle.DropDownList;
            cboTrangThai.Items.AddRange(new object[]
            {
                "Đang mượn",
                "Đã trả",
                "Quá hạn"
            });
            cboTrangThai.SelectedIndex = 0;

            ThemTruong("Mã phiếu:", txtMaPhieu, 0, 0);
            ThemTruong("Mã sinh viên:", txtMaSV, 2, 0);

            ThemTruong("Họ tên sinh viên:", txtHoTen, 0, 1);
            ThemTruong("Thiết bị:", cboThietBi, 2, 1);

            ThemTruong("Số lượng:", txtSoLuong, 0, 2);
            ThemTruong("Ngày mượn:", dtpNgayMuon, 2, 2);

            ThemTruong("Hạn trả:", dtpHanTra, 0, 3);

            // Ngày trả thực tế có checkbox để thể hiện chưa trả
            Label lblNgayTra = new Label();
            lblNgayTra.Text = "Ngày trả thực tế:";
            lblNgayTra.Dock = DockStyle.Fill;
            lblNgayTra.TextAlign = ContentAlignment.MiddleLeft;

            inputLayout.Controls.Add(lblNgayTra, 2, 3);

            FlowLayoutPanel panelNgayTra = new FlowLayoutPanel();
            panelNgayTra.Dock = DockStyle.Fill;
            panelNgayTra.WrapContents = false;

            chkNgayTra.Text = "Đã trả";
            chkNgayTra.AutoSize = true;
            chkNgayTra.Checked = false;

            dtpNgayTra.Width = 125;
            dtpNgayTra.Enabled = false;

            panelNgayTra.Controls.Add(chkNgayTra);
            panelNgayTra.Controls.Add(dtpNgayTra);

            inputLayout.Controls.Add(panelNgayTra, 3, 3);

            ThemTruong("Trạng thái:", cboTrangThai, 0, 4);

            // Các nút chức năng
            TableLayoutPanel buttonLayout = new TableLayoutPanel();
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
                new RowStyle(SizeType.Absolute, 42));

            tableLayout.RowStyles.Add(
                new RowStyle(SizeType.Percent, 100));

            mainLayout.Controls.Add(tableLayout, 0, 1);

            FlowLayoutPanel filterPanel = new FlowLayoutPanel();
            filterPanel.Dock = DockStyle.Fill;
            filterPanel.WrapContents = false;
            filterPanel.AutoScroll = true;

            Label lblTim = new Label();
            lblTim.Text = "Tìm mã SV/thiết bị:";
            lblTim.AutoSize = true;
            lblTim.Margin = new Padding(4, 10, 3, 0);

            txtTimKiem.Width = 190;
            txtTimKiem.Margin = new Padding(3, 6, 3, 0);

            Label lblLoc = new Label();
            lblLoc.Text = "Lọc trạng thái:";
            lblLoc.AutoSize = true;
            lblLoc.Margin = new Padding(8, 10, 3, 0);

            cboLocTrangThai.DropDownStyle =
                ComboBoxStyle.DropDownList;

            cboLocTrangThai.Width = 130;
            cboLocTrangThai.Margin = new Padding(3, 6, 3, 0);

            cboLocTrangThai.Items.AddRange(new object[]
            {
                "Tất cả",
                "Đang mượn",
                "Quá hạn",
                "Đã trả"
            });

            cboLocTrangThai.SelectedIndex = 0;

            filterPanel.Controls.Add(lblTim);
            filterPanel.Controls.Add(txtTimKiem);
            filterPanel.Controls.Add(lblLoc);
            filterPanel.Controls.Add(cboLocTrangThai);

            tableLayout.Controls.Add(filterPanel, 0, 0);

            // DataGridView
            dgvPhieuMuon.Dock = DockStyle.Fill;
            dgvPhieuMuon.AutoGenerateColumns = false;
            dgvPhieuMuon.SelectionMode =
                DataGridViewSelectionMode.FullRowSelect;

            dgvPhieuMuon.MultiSelect = false;
            dgvPhieuMuon.ReadOnly = true;
            dgvPhieuMuon.AllowUserToAddRows = false;
            dgvPhieuMuon.AllowUserToDeleteRows = false;
            dgvPhieuMuon.RowHeadersVisible = false;

            dgvPhieuMuon.AutoSizeColumnsMode =
                DataGridViewAutoSizeColumnsMode.Fill;

            ThemCot("MaPhieu", "Mã phiếu", 70);
            ThemCot("MaSinhVien", "Mã sinh viên", 90);
            ThemCot("HoTen", "Họ tên sinh viên", 120);
            ThemCot("ThietBi", "Thiết bị", 120);
            ThemCot("SoLuong", "Số lượng", 55);
            ThemCot("NgayMuon", "Ngày mượn", 90);
            ThemCot("HanTra", "Hạn trả", 90);
            ThemCot("NgayTraThucTe", "Ngày trả thực tế", 100);
            ThemCot("TrangThai", "Trạng thái", 85);

            dgvPhieuMuon.Columns["NgayMuon"]
                .DefaultCellStyle.Format = "dd/MM/yyyy";

            dgvPhieuMuon.Columns["HanTra"]
                .DefaultCellStyle.Format = "dd/MM/yyyy";

            dgvPhieuMuon.Columns["NgayTraThucTe"]
                .DefaultCellStyle.Format = "dd/MM/yyyy";

            tableLayout.Controls.Add(dgvPhieuMuon, 0, 1);

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

            dgvPhieuMuon.Columns.Add(cot);
        }

        private void GanSuKien()
        {
            bindingSource.DataSource = danhSach;
            dgvPhieuMuon.DataSource = bindingSource;

            btnThem.Click += BtnThem_Click;
            btnSua.Click += BtnSua_Click;
            btnXoa.Click += BtnXoa_Click;
            btnLamMoi.Click += BtnLamMoi_Click;

            dgvPhieuMuon.SelectionChanged +=
                DgvPhieuMuon_SelectionChanged;

            txtTimKiem.TextChanged +=
                delegate { CapNhatDanhSach(); };

            cboLocTrangThai.SelectedIndexChanged +=
                delegate { CapNhatDanhSach(); };

            chkNgayTra.CheckedChanged += delegate
            {
                dtpNgayTra.Enabled = chkNgayTra.Checked;

                if (chkNgayTra.Checked)
                {
                    cboTrangThai.SelectedItem = "Đã trả";

                    if (dtpNgayTra.Value.Date < dtpNgayMuon.Value.Date)
                        dtpNgayTra.Value = dtpNgayMuon.Value.Date;
                }
                else
                {
                    if (cboTrangThai.Text == "Đã trả")
                        cboTrangThai.SelectedItem = "Đang mượn";
                }
            };

            cboTrangThai.SelectedIndexChanged += delegate
            {
                if (dangNapDuLieu)
                    return;

                if (cboTrangThai.Text == "Đã trả")
                {
                    chkNgayTra.Checked = true;

                    if (dtpNgayTra.Value.Date < dtpNgayMuon.Value.Date)
                        dtpNgayTra.Value = dtpNgayMuon.Value.Date;
                }
                else
                {
                    chkNgayTra.Checked = false;
                }
            };
        }

        private string TinhTrang(PhieuMuon p)
        {
            if (p.NgayTraThucTe != DateTime.MinValue)
                return "Đã trả";

            if (p.HanTra.Date < DateTime.Today)
                return "Quá hạn";

            return "Đang mượn";
        }

        private void CapNhatDanhSach()
        {
            string tuKhoa = txtTimKiem.Text.Trim()
                .ToLowerInvariant();

            string trangThaiLoc = cboLocTrangThai.Text;

            var ketQua = danhSach.Where(p =>
                (p.MaSinhVien.ToLowerInvariant().Contains(tuKhoa)
                 || p.ThietBi.ToLowerInvariant().Contains(tuKhoa))
                &&
                (trangThaiLoc == "Tất cả"
                 || string.IsNullOrEmpty(trangThaiLoc)
                 || TinhTrang(p) == trangThaiLoc)
            ).ToList();

            bindingSource.DataSource =
                new BindingList<PhieuMuon>(ketQua);

            dgvPhieuMuon.DataSource = bindingSource;

            lblTong.Text = "Tổng số phiếu: " + danhSach.Count;

            lblQuaHan.Text = "Số phiếu quá hạn: " +
                danhSach.Count(p => TinhTrang(p) == "Quá hạn");
        }

        private PhieuMuon LayDuLieuTuForm()
        {
            int soLuong;

            int.TryParse(txtSoLuong.Text.Trim(), out soLuong);

            PhieuMuon p = new PhieuMuon();

            p.MaPhieu = txtMaPhieu.Text.Trim();
            p.MaSinhVien = txtMaSV.Text.Trim();
            p.HoTen = txtHoTen.Text.Trim();
            p.ThietBi = cboThietBi.Text;
            p.SoLuong = soLuong;
            p.NgayMuon = dtpNgayMuon.Value.Date;
            p.HanTra = dtpHanTra.Value.Date;

            if (chkNgayTra.Checked)
                p.NgayTraThucTe = dtpNgayTra.Value.Date;
            else
                p.NgayTraThucTe = DateTime.MinValue;

            p.TrangThai = TinhTrang(p);

            return p;
        }

        private bool KiemTraDuLieu(bool dangCapNhat)
        {
            errorProvider.Clear();
            bool hopLe = true;

            string ma = txtMaPhieu.Text.Trim();

            if (ma == "")
            {
                errorProvider.SetError(
                    txtMaPhieu, "Vui lòng nhập mã phiếu.");
                hopLe = false;
            }
            else if (danhSach.Any(p =>
                p.MaPhieu.Equals(
                    ma, StringComparison.OrdinalIgnoreCase)
                && (!dangCapNhat || p != phieuDangChon)))
            {
                errorProvider.SetError(
                    txtMaPhieu, "Mã phiếu đã tồn tại.");
                hopLe = false;
            }

            if (txtMaSV.Text.Trim() == "")
            {
                errorProvider.SetError(
                    txtMaSV, "Vui lòng nhập mã sinh viên.");
                hopLe = false;
            }

            if (txtHoTen.Text.Trim() == "")
            {
                errorProvider.SetError(
                    txtHoTen, "Vui lòng nhập họ tên sinh viên.");
                hopLe = false;
            }

            if (cboThietBi.SelectedIndex < 0)
            {
                errorProvider.SetError(
                    cboThietBi, "Vui lòng chọn thiết bị.");
                hopLe = false;
            }

            int soLuong;

            if (!int.TryParse(txtSoLuong.Text.Trim(), out soLuong)
                || soLuong <= 0)
            {
                errorProvider.SetError(
                    txtSoLuong, "Số lượng phải lớn hơn 0.");
                hopLe = false;
            }

            if (dtpHanTra.Value.Date < dtpNgayMuon.Value.Date)
            {
                errorProvider.SetError(
                    dtpHanTra, "Hạn trả không được trước ngày mượn.");
                hopLe = false;
            }

            if (chkNgayTra.Checked &&
                dtpNgayTra.Value.Date < dtpNgayMuon.Value.Date)
            {
                errorProvider.SetError(
                    dtpNgayTra,
                    "Ngày trả thực tế không được trước ngày mượn.");
                hopLe = false;
            }

            if (chkNgayTra.Checked &&
                dtpNgayTra.Value.Date > DateTime.Today)
            {
                errorProvider.SetError(
                    dtpNgayTra,
                    "Ngày trả thực tế không được ở tương lai.");
                hopLe = false;
            }

            if (cboTrangThai.Text == "Quá hạn" &&
                chkNgayTra.Checked)
            {
                errorProvider.SetError(
                    cboTrangThai,
                    "Phiếu đã trả không thể ở trạng thái quá hạn.");
                hopLe = false;
            }

            return hopLe;
        }

        private void BtnThem_Click(object sender, EventArgs e)
        {
            if (!KiemTraDuLieu(false))
                return;

            if (MessageBox.Show(
                "Bạn có muốn thêm phiếu mượn này không?",
                "Xác nhận",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question) != DialogResult.Yes)
                return;

            danhSach.Add(LayDuLieuTuForm());

            LamMoiForm();
            CapNhatDanhSach();
        }

        private void BtnSua_Click(object sender, EventArgs e)
        {
            if (phieuDangChon == null ||
                !danhSach.Contains(phieuDangChon))
            {
                MessageBox.Show("Bạn hãy chọn phiếu cần cập nhật.");
                return;
            }

            if (!KiemTraDuLieu(true))
                return;

            if (MessageBox.Show(
                "Bạn có muốn cập nhật phiếu mượn này không?",
                "Xác nhận",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question) != DialogResult.Yes)
                return;

            PhieuMuon moi = LayDuLieuTuForm();

            phieuDangChon.MaPhieu = moi.MaPhieu;
            phieuDangChon.MaSinhVien = moi.MaSinhVien;
            phieuDangChon.HoTen = moi.HoTen;
            phieuDangChon.ThietBi = moi.ThietBi;
            phieuDangChon.SoLuong = moi.SoLuong;
            phieuDangChon.NgayMuon = moi.NgayMuon;
            phieuDangChon.HanTra = moi.HanTra;
            phieuDangChon.NgayTraThucTe = moi.NgayTraThucTe;
            phieuDangChon.TrangThai = moi.TrangThai;

            danhSach.ResetBindings();
            LamMoiForm();
            CapNhatDanhSach();
        }

        private void BtnXoa_Click(object sender, EventArgs e)
        {
            if (phieuDangChon == null ||
                !danhSach.Contains(phieuDangChon))
            {
                MessageBox.Show("Bạn hãy chọn phiếu cần xóa.");
                return;
            }

            if (MessageBox.Show(
                "Bạn có chắc chắn muốn xóa phiếu " +
                phieuDangChon.MaPhieu + " không?",
                "Xác nhận xóa",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning) != DialogResult.Yes)
                return;

            danhSach.Remove(phieuDangChon);

            LamMoiForm();
            CapNhatDanhSach();
        }

        private void DgvPhieuMuon_SelectionChanged(
            object sender, EventArgs e)
        {
            if (dangNapDuLieu)
                return;

            PhieuMuon p = bindingSource.Current as PhieuMuon;

            if (p == null)
                return;

            phieuDangChon = p;
            dangNapDuLieu = true;

            txtMaPhieu.Text = p.MaPhieu;
            txtMaSV.Text = p.MaSinhVien;
            txtHoTen.Text = p.HoTen;
            cboThietBi.SelectedItem = p.ThietBi;
            txtSoLuong.Text = p.SoLuong.ToString();

            dtpNgayMuon.Value = p.NgayMuon;
            dtpHanTra.Value = p.HanTra;

            if (p.NgayTraThucTe != DateTime.MinValue)
            {
                chkNgayTra.Checked = true;
                dtpNgayTra.Value = p.NgayTraThucTe;
            }
            else
            {
                chkNgayTra.Checked = false;
                dtpNgayTra.Value = DateTime.Today;
            }

            cboTrangThai.SelectedItem = TinhTrang(p);

            dangNapDuLieu = false;
            errorProvider.Clear();
        }

        private void LamMoiForm()
        {
            dangNapDuLieu = true;

            phieuDangChon = null;
            errorProvider.Clear();

            txtMaPhieu.Clear();
            txtMaSV.Clear();
            txtHoTen.Clear();
            txtSoLuong.Clear();

            if (cboThietBi.Items.Count > 0)
                cboThietBi.SelectedIndex = 0;

            dtpNgayMuon.Value = DateTime.Today;
            dtpHanTra.Value = DateTime.Today.AddDays(7);
            dtpNgayTra.Value = DateTime.Today;

            chkNgayTra.Checked = false;
            dtpNgayTra.Enabled = false;

            cboTrangThai.SelectedItem = "Đang mượn";

            dgvPhieuMuon.ClearSelection();

            dangNapDuLieu = false;
        }

        private void BtnLamMoi_Click(object sender, EventArgs e)
        {
            LamMoiForm();
        }

        private void MenuXuat_Click(object sender, EventArgs e)
        {
            SaveFileDialog save = new SaveFileDialog();

            save.Title = "Xuất danh sách phiếu mượn";
            save.Filter = "CSV (*.csv)|*.csv";
            save.FileName = "DanhSachPhieuMuon.csv";

            if (save.ShowDialog() != DialogResult.OK)
                return;

            var ketQua = bindingSource.List
                .Cast<PhieuMuon>().ToList();

            using (StreamWriter writer = new StreamWriter(
                save.FileName, false, Encoding.UTF8))
            {
                writer.WriteLine(
                    "MaPhieu,MaSinhVien,HoTen,ThietBi,SoLuong," +
                    "NgayMuon,HanTra,NgayTraThucTe,TrangThai");

                foreach (PhieuMuon p in ketQua)
                {
                    writer.WriteLine(
                        Csv(p.MaPhieu) + "," +
                        Csv(p.MaSinhVien) + "," +
                        Csv(p.HoTen) + "," +
                        Csv(p.ThietBi) + "," +
                        p.SoLuong + "," +
                        Csv(p.NgayMuon.ToString("dd/MM/yyyy")) + "," +
                        Csv(p.HanTra.ToString("dd/MM/yyyy")) + "," +
                        Csv(p.NgayTraThucTe == DateTime.MinValue
                            ? ""
                            : p.NgayTraThucTe.ToString("dd/MM/yyyy")) + "," +
                        Csv(TinhTrang(p)));
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