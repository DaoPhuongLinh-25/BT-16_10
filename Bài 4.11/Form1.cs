

using System;
using System.ComponentModel;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace WinFormsApp11
{
    public partial class Form1 : Form
    {
        private BindingList<RepairOrder> danhSach =
            new BindingList<RepairOrder>();

        private BindingList<RepairOrder> danhSachHienThi =
            new BindingList<RepairOrder>();

        private BindingSource bindingSource = new BindingSource();
        private ErrorProvider errorProvider = new ErrorProvider();

        private TextBox txtMa = new TextBox();
        private TextBox txtTen = new TextBox();
        private TextBox txtSDT = new TextBox();
        private TextBox txtMoTa = new TextBox();
        private TextBox txtChiPhi = new TextBox();
        private TextBox txtTimKiem = new TextBox();

        private ComboBox cboLoai = new ComboBox();
        private ComboBox cboTrangThai = new ComboBox();
        private ComboBox cboLoc = new ComboBox();

        private DateTimePicker dtpNgay = new DateTimePicker();
        private PictureBox picAnh = new PictureBox();
        private DataGridView dgv = new DataGridView();

        private ToolStripStatusLabel lblThongKe =
            new ToolStripStatusLabel();
        private string duongDanAnh = "";
        private string maDangSua = "";
        private bool dangNapDuLieu = false;

        public Form1()
        {
            BuildUI();
            CauHinhBang();
            GanSuKien();
            CapNhatDanhSach();
        }

        private void BuildUI()
        {
            Text = "Repair Service Tracker - Quản lý sửa chữa";
            StartPosition = FormStartPosition.CenterScreen;
            Size = new Size(1200, 750);
            MinimumSize = new Size(900, 600);

            errorProvider.ContainerControl = this;

            MenuStrip menu = new MenuStrip();

            ToolStripMenuItem menuFile =
                new ToolStripMenuItem("Tệp");

            ToolStripMenuItem menuXuat =
                new ToolStripMenuItem("Xuất danh sách CSV");

            ToolStripMenuItem menuThoat =
                new ToolStripMenuItem("Thoát");

            menuXuat.Click += (s, e) => XuatCSV();
            menuThoat.Click += (s, e) => Close();

            menuFile.DropDownItems.Add(menuXuat);
            menuFile.DropDownItems.Add(menuThoat);
            menu.Items.Add(menuFile);

            MainMenuStrip = menu;
            Controls.Add(menu);

            StatusStrip status = new StatusStrip();
            lblThongKe.Text = "Số đơn: 0 | Tổng chi phí chưa hoàn tất: 0";
            status.Items.Add(lblThongKe);
            Controls.Add(status);

            
            

            SplitContainer split = new SplitContainer();
            split.Dock = DockStyle.Fill;
            split.Orientation = Orientation.Vertical;
            split.Panel1MinSize = 0;
            split.Panel2MinSize = 0;
            split.SplitterDistance = 390;


            Controls.Add(split);
            split.BringToFront();

            // Khu vực nhập thông tin
            Panel panelTrai = split.Panel1;

            TableLayoutPanel form = new TableLayoutPanel();
            form.Dock = DockStyle.Fill;
            form.AutoScroll = true;
            form.Padding = new Padding(10);
            form.ColumnCount = 2;
            form.ColumnStyles.Add(
                new ColumnStyle(SizeType.Absolute, 125));
            form.ColumnStyles.Add(
                new ColumnStyle(SizeType.Percent, 100));

            panelTrai.Controls.Add(form);

            Label tieuDe = new Label();
            tieuDe.Text = "THÔNG TIN ĐƠN SỬA CHỮA";
            tieuDe.Font = new Font("Arial", 13, FontStyle.Bold);
            tieuDe.AutoSize = true;
            ThemDong(form, tieuDe, 2);

            ThemTruong(form, "Mã đơn:", txtMa);
            ThemTruong(form, "Khách hàng:", txtTen);
            ThemTruong(form, "Số điện thoại:", txtSDT);

            cboLoai.DropDownStyle = ComboBoxStyle.DropDownList;
            cboLoai.Items.AddRange(new object[]
            {
                "Điện thoại", "Laptop", "Máy tính",
                "Máy tính bảng", "Thiết bị khác"
            });
            cboLoai.SelectedIndex = 0;
            ThemTruong(form, "Loại thiết bị:", cboLoai);

            ThemTruong(form, "Mô tả lỗi:", txtMoTa);
            txtMoTa.Multiline = true;
            txtMoTa.Height = 55;
            txtMoTa.ScrollBars = ScrollBars.Vertical;

            dtpNgay.Format = DateTimePickerFormat.Short;
            ThemTruong(form, "Ngày tiếp nhận:", dtpNgay);

            ThemTruong(form, "Chi phí (VNĐ):", txtChiPhi);

            cboTrangThai.DropDownStyle = ComboBoxStyle.DropDownList;
            cboTrangThai.Items.AddRange(new object[]
            {
                "Mới tiếp nhận", "Đang sửa", "Chờ linh kiện",
                "Hoàn tất", "Đã hủy"
            });
            cboTrangThai.SelectedIndex = 0;
            ThemTruong(form, "Trạng thái:", cboTrangThai);

            picAnh.Height = 140;
            picAnh.Dock = DockStyle.Fill;
            picAnh.BorderStyle = BorderStyle.FixedSingle;
            picAnh.SizeMode = PictureBoxSizeMode.Zoom;
            ThemDong(form, picAnh, 2);
            ThemDong(form, TaoNut("Chọn ảnh thiết bị", ChonAnh), 2);

            FlowLayoutPanel nut = new FlowLayoutPanel();
            nut.AutoSize = true;
            nut.WrapContents = true;
            nut.Controls.Add(TaoNut("Thêm", ThemDon));
            nut.Controls.Add(TaoNut("Cập nhật", CapNhatDon));
            nut.Controls.Add(TaoNut("Xóa", XoaDon));
            nut.Controls.Add(TaoNut("Làm mới", LamMoi));
            ThemDong(form, nut, 2);

            // Khu vực tìm kiếm và danh sách
            Panel panelPhai = split.Panel2;

            TableLayoutPanel khuVuc = new TableLayoutPanel();
            khuVuc.Dock = DockStyle.Fill;
            khuVuc.Padding = new Padding(8);
            khuVuc.ColumnCount = 1;
            khuVuc.RowCount = 3;
            khuVuc.RowStyles.Add(
                new RowStyle(SizeType.Absolute, 42));
            khuVuc.RowStyles.Add(
                new RowStyle(SizeType.Absolute, 42));
            khuVuc.RowStyles.Add(
                new RowStyle(SizeType.Percent, 100));

            panelPhai.Controls.Add(khuVuc);

            Label tieuDeDS = new Label();
            tieuDeDS.Text = "DANH SÁCH ĐƠN SỬA CHỮA";
            tieuDeDS.Font = new Font("Arial", 12, FontStyle.Bold);
            tieuDeDS.Dock = DockStyle.Fill;
            tieuDeDS.TextAlign = ContentAlignment.MiddleLeft;
            khuVuc.Controls.Add(tieuDeDS, 0, 0);

            FlowLayoutPanel timLoc = new FlowLayoutPanel();
            timLoc.Dock = DockStyle.Fill;
            timLoc.WrapContents = false;
            timLoc.AutoScroll = true;

            Label lblTim = new Label();
            lblTim.Text = "Tìm:";
            lblTim.AutoSize = true;
            lblTim.Margin = new Padding(3, 8, 3, 3);

            txtTimKiem.Width = 150;

            Label lblLoc = new Label();
            lblLoc.Text = "Trạng thái:";
            lblLoc.AutoSize = true;
            lblLoc.Margin = new Padding(5, 8, 3, 3);

            cboLoc.DropDownStyle = ComboBoxStyle.DropDownList;
            cboLoc.Width = 130;
            cboLoc.Items.AddRange(new object[]
            {
                "Tất cả", "Mới tiếp nhận", "Đang sửa",
                "Chờ linh kiện", "Hoàn tất", "Đã hủy"
            });
            cboLoc.SelectedIndex = 0;

            timLoc.Controls.Add(lblTim);
            timLoc.Controls.Add(txtTimKiem);
            timLoc.Controls.Add(lblLoc);
            timLoc.Controls.Add(cboLoc);

            khuVuc.Controls.Add(timLoc, 0, 1);

            dgv.Dock = DockStyle.Fill;
            dgv.ReadOnly = true;
            dgv.AllowUserToAddRows = false;
            dgv.AllowUserToDeleteRows = false;
            dgv.MultiSelect = false;
            dgv.SelectionMode =
                DataGridViewSelectionMode.FullRowSelect;
            dgv.AutoGenerateColumns = false;
            dgv.AutoSizeColumnsMode =
                DataGridViewAutoSizeColumnsMode.Fill;
            dgv.RowHeadersVisible = false;

            khuVuc.Controls.Add(dgv, 0, 2);
        }

        private void ThemDong(
            TableLayoutPanel panel, Control control, int soCot)
        {
            int row = panel.RowCount;
            panel.RowCount++;
            panel.RowStyles.Add(
                new RowStyle(SizeType.AutoSize));

            panel.Controls.Add(control, 0, row);

            if (soCot == 2)
            {
                panel.SetColumnSpan(control, 2);
            }
        }

        private void ThemTruong(
            TableLayoutPanel panel, string ten, Control control)
        {
            int row = panel.RowCount;
            panel.RowCount++;
            panel.RowStyles.Add(
                new RowStyle(SizeType.AutoSize));

            Label label = new Label();
            label.Text = ten;
            label.AutoSize = true;
            label.Anchor = AnchorStyles.Left;
            label.Margin = new Padding(3, 7, 3, 3);

            control.Dock = DockStyle.Fill;
            control.Margin = new Padding(3, 4, 3, 4);

            panel.Controls.Add(label, 0, row);
            panel.Controls.Add(control, 1, row);
        }

        private Button TaoNut(string ten, Action hanhDong)
        {
            Button button = new Button();
            button.Text = ten;
            button.AutoSize = true;
            button.Height = 32;
            button.Click += (s, e) => hanhDong();
            return button;
        }

        private void CauHinhBang()
        {
            ThemCot("Mã đơn", "MaDon", 80);
            ThemCot("Khách hàng", "TenKhachHang", 130);
            ThemCot("Số điện thoại", "SoDienThoai", 100);
            ThemCot("Thiết bị", "LoaiThietBi", 100);
            ThemCot("Ngày nhận", "NgayTiepNhan", 95);
            ThemCot("Chi phí (VNĐ)", "ChiPhiDuKien", 110);
            ThemCot("Trạng thái", "TrangThai", 110);

            dgv.Columns["NgayTiepNhan"].DefaultCellStyle.Format =
                "dd/MM/yyyy";

            dgv.Columns["ChiPhiDuKien"].DefaultCellStyle.Format =
                "N0";

            dgv.Columns["ChiPhiDuKien"].DefaultCellStyle.Alignment =
                DataGridViewContentAlignment.MiddleRight;

            bindingSource.DataSource = danhSachHienThi;
            dgv.DataSource = bindingSource;
        }

        private void ThemCot(
            string tieuDe, string tenThuocTinh, float doRong)
        {
            DataGridViewTextBoxColumn cot =
                new DataGridViewTextBoxColumn();

            cot.HeaderText = tieuDe;
            cot.DataPropertyName = tenThuocTinh;
            cot.Name = tenThuocTinh;
            cot.MinimumWidth = 70;
            cot.FillWeight = doRong;

            dgv.Columns.Add(cot);
        }

        private void GanSuKien()
        {
            txtTimKiem.TextChanged += (s, e) => LocDuLieu();

            cboLoc.SelectedIndexChanged +=
                (s, e) => LocDuLieu();

            dgv.SelectionChanged += (s, e) =>
            {
                if (dgv.CurrentRow != null &&
                    dgv.CurrentRow.DataBoundItem is RepairOrder don)
                {
                    NapLenForm(don);
                }
            };
        }

        private void ThemDon()
        {
            if (!KiemTraDuLieu())
                return;

            foreach (RepairOrder don in danhSach)
            {
                if (don.MaDon.Equals(
                    txtMa.Text.Trim(),
                    StringComparison.OrdinalIgnoreCase))
                {
                    errorProvider.SetError(
                        txtMa, "Mã đơn đã tồn tại!");
                    MessageBox.Show("Mã đơn không được trùng.");
                    return;
                }
            }

            RepairOrder moi = DocTuForm();
            danhSach.Add(moi);

            CapNhatDanhSach();
            ChonDonTheoMa(moi.MaDon);

            MessageBox.Show("Thêm đơn sửa chữa thành công!");
        }

        private void CapNhatDon()
        {
            if (string.IsNullOrEmpty(maDangSua))
            {
                MessageBox.Show("Bạn hãy chọn đơn cần cập nhật.");
                return;
            }

            if (!KiemTraDuLieu())
                return;

            foreach (RepairOrder don in danhSach)
            {
                if (don != TimDonTheoMa(maDangSua) &&
                    don.MaDon.Equals(
                        txtMa.Text.Trim(),
                        StringComparison.OrdinalIgnoreCase))
                {
                    errorProvider.SetError(
                        txtMa, "Mã đơn đã tồn tại!");
                    MessageBox.Show("Mã đơn không được trùng.");
                    return;
                }
            }

            RepairOrder cu = TimDonTheoMa(maDangSua);

            if (cu == null)
                return;

            RepairOrder moi = DocTuForm();

            cu.MaDon = moi.MaDon;
            cu.TenKhachHang = moi.TenKhachHang;
            cu.SoDienThoai = moi.SoDienThoai;
            cu.LoaiThietBi = moi.LoaiThietBi;
            cu.MoTaLoi = moi.MoTaLoi;
            cu.NgayTiepNhan = moi.NgayTiepNhan;
            cu.ChiPhiDuKien = moi.ChiPhiDuKien;
            cu.TrangThai = moi.TrangThai;
            cu.DuongDanAnh = moi.DuongDanAnh;

            maDangSua = cu.MaDon;

            CapNhatDanhSach();
            ChonDonTheoMa(cu.MaDon);

            MessageBox.Show("Cập nhật đơn thành công!");
        }

        private void XoaDon()
        {
            if (string.IsNullOrEmpty(maDangSua))
            {
                MessageBox.Show("Bạn hãy chọn đơn cần xóa.");
                return;
            }

            RepairOrder don = TimDonTheoMa(maDangSua);

            if (don == null)
                return;

            DialogResult kq = MessageBox.Show(
                "Bạn có chắc muốn xóa đơn " + don.MaDon + "?",
                "Xác nhận xóa",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (kq == DialogResult.Yes)
            {
                danhSach.Remove(don);
                LamMoi();
                CapNhatDanhSach();
            }
        }

        private bool KiemTraDuLieu()
        {
            errorProvider.Clear();
            bool hopLe = true;

            if (string.IsNullOrWhiteSpace(txtMa.Text))
            {
                errorProvider.SetError(txtMa, "Hãy nhập mã đơn.");
                hopLe = false;
            }

            if (string.IsNullOrWhiteSpace(txtTen.Text))
            {
                errorProvider.SetError(
                    txtTen, "Tên khách hàng không được trống.");
                hopLe = false;
            }

            if (string.IsNullOrWhiteSpace(txtSDT.Text))
            {
                errorProvider.SetError(
                    txtSDT, "Số điện thoại không được trống.");
                hopLe = false;
            }

            if (string.IsNullOrWhiteSpace(txtMoTa.Text))
            {
                errorProvider.SetError(
                    txtMoTa, "Mô tả lỗi là bắt buộc.");
                hopLe = false;
            }

            decimal chiPhi;

            if (string.IsNullOrWhiteSpace(txtChiPhi.Text) ||
                !decimal.TryParse(
                    txtChiPhi.Text.Trim(),
                    NumberStyles.Number,
                    CultureInfo.CurrentCulture,
                    out chiPhi) ||
                chiPhi < 0)
            {
                errorProvider.SetError(
                    txtChiPhi,
                    "Chi phí phải là số không âm.");
                hopLe = false;
            }

            if (!hopLe)
            {
                MessageBox.Show(
                    "Vui lòng sửa các thông tin không hợp lệ.",
                    "Lỗi dữ liệu",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }

            return hopLe;
        }

        private RepairOrder DocTuForm()
        {
            decimal.TryParse(
                txtChiPhi.Text.Trim(),
                NumberStyles.Number,
                CultureInfo.CurrentCulture,
                out decimal chiPhi);

            return new RepairOrder
            {
                MaDon = txtMa.Text.Trim(),
                TenKhachHang = txtTen.Text.Trim(),
                SoDienThoai = txtSDT.Text.Trim(),
                LoaiThietBi = cboLoai.Text,
                MoTaLoi = txtMoTa.Text.Trim(),
                NgayTiepNhan = dtpNgay.Value.Date,
                ChiPhiDuKien = chiPhi,
                TrangThai = cboTrangThai.Text,
                DuongDanAnh = duongDanAnh
            };
        }

        private RepairOrder TimDonTheoMa(string ma)
        {
            foreach (RepairOrder don in danhSach)
            {
                if (don.MaDon.Equals(
                    ma, StringComparison.OrdinalIgnoreCase))
                    return don;
            }

            return null;
        }

        private void NapLenForm(RepairOrder don)
        {
            if (don == null)
                return;

            dangNapDuLieu = true;

            txtMa.Text = don.MaDon;
            txtTen.Text = don.TenKhachHang;
            txtSDT.Text = don.SoDienThoai;
            txtMoTa.Text = don.MoTaLoi;
            txtChiPhi.Text = don.ChiPhiDuKien.ToString("0.##");
            dtpNgay.Value = don.NgayTiepNhan;

            cboLoai.SelectedItem = don.LoaiThietBi;
            cboTrangThai.SelectedItem = don.TrangThai;

            if (cboLoai.SelectedIndex < 0)
                cboLoai.SelectedIndex = 0;

            if (cboTrangThai.SelectedIndex < 0)
                cboTrangThai.SelectedIndex = 0;

            maDangSua = don.MaDon;
            duongDanAnh = don.DuongDanAnh;

            HienThiAnh(duongDanAnh);

            dangNapDuLieu = false;
        }

        private void ChonDonTheoMa(string ma)
        {
            foreach (DataGridViewRow row in dgv.Rows)
            {
                if (row.DataBoundItem is RepairOrder don &&
                    don.MaDon.Equals(
                        ma, StringComparison.OrdinalIgnoreCase))
                {
                    row.Selected = true;
                    dgv.CurrentCell = row.Cells[0];
                    return;
                }
            }
        }

        private void LocDuLieu()
        {
            string tuKhoa = txtTimKiem.Text.Trim();
            string trangThai = cboLoc.Text;

            danhSachHienThi.RaiseListChangedEvents = false;
            danhSachHienThi.Clear();

            foreach (RepairOrder don in danhSach)
            {
                bool dungTuKhoa =
                    don.MaDon.IndexOf(
                        tuKhoa, StringComparison.OrdinalIgnoreCase) >= 0 ||
                    don.TenKhachHang.IndexOf(
                        tuKhoa, StringComparison.OrdinalIgnoreCase) >= 0 ||
                    don.SoDienThoai.IndexOf(
                        tuKhoa, StringComparison.OrdinalIgnoreCase) >= 0;

                bool dungTrangThai =
                    trangThai == "Tất cả" ||
                    don.TrangThai == trangThai;

                if (dungTuKhoa && dungTrangThai)
                    danhSachHienThi.Add(don);
            }

            danhSachHienThi.RaiseListChangedEvents = true;
            bindingSource.ResetBindings(false);

            CapNhatThongKe();
        }

        private void CapNhatDanhSach()
        {
            LocDuLieu();
            CapNhatThongKe();
        }

        private void CapNhatThongKe()
        {
            decimal tong = 0;

            foreach (RepairOrder don in danhSach)
            {
                if (don.TrangThai != "Hoàn tất")
                    tong += don.ChiPhiDuKien;
            }

            lblThongKe.Text =
                "Số đơn theo bộ lọc: " + danhSachHienThi.Count +
                " | Tổng chi phí đơn chưa hoàn tất: " +
                tong.ToString("N0") + " VNĐ";
        }

        private void ChonAnh()
        {
            using (OpenFileDialog open = new OpenFileDialog())
            {
                open.Title = "Chọn ảnh thiết bị";
                open.Filter =
                    "Ảnh (*.jpg;*.jpeg;*.png;*.bmp)|*.jpg;*.jpeg;*.png;*.bmp";

                if (open.ShowDialog() == DialogResult.OK)
                {
                    duongDanAnh = open.FileName;
                    HienThiAnh(duongDanAnh);
                }
            }
        }

        private void HienThiAnh(string duongDan)
        {
            Image anhCu = picAnh.Image;
            picAnh.Image = null;

            if (anhCu != null)
                anhCu.Dispose();

            if (!string.IsNullOrWhiteSpace(duongDan) &&
                File.Exists(duongDan))
            {
                using (Image anhGoc = Image.FromFile(duongDan))
                {
                    picAnh.Image = new Bitmap(anhGoc);
                }
            }

            picAnh.SizeMode = PictureBoxSizeMode.Zoom;
        }

        private void LamMoi()
        {
            dangNapDuLieu = true;

            txtMa.Clear();
            txtTen.Clear();
            txtSDT.Clear();
            txtMoTa.Clear();
            txtChiPhi.Clear();

            cboLoai.SelectedIndex = 0;
            cboTrangThai.SelectedIndex = 0;
            dtpNgay.Value = DateTime.Today;

            duongDanAnh = "";
            maDangSua = "";

            HienThiAnh("");
            errorProvider.Clear();

            dgv.ClearSelection();
            dangNapDuLieu = false;
        }

        private void XuatCSV()
        {
            using (SaveFileDialog save = new SaveFileDialog())
            {
                save.Title = "Xuất danh sách đơn sửa chữa";
                save.Filter = "CSV (*.csv)|*.csv";
                save.FileName = "DanhSachSuaChua.csv";

                if (save.ShowDialog() != DialogResult.OK)
                    return;

                StringBuilder sb = new StringBuilder();

                sb.AppendLine(
                    "MaDon,TenKhachHang,SoDienThoai,LoaiThietBi," +
                    "MoTaLoi,NgayTiepNhan,ChiPhiDuKien,TrangThai");

                foreach (RepairOrder don in danhSachHienThi)
                {
                    sb.AppendLine(
                        CSV(don.MaDon) + "," +
                        CSV(don.TenKhachHang) + "," +
                        CSV(don.SoDienThoai) + "," +
                        CSV(don.LoaiThietBi) + "," +
                        CSV(don.MoTaLoi) + "," +
                        CSV(don.NgayTiepNhan.ToString("dd/MM/yyyy")) + "," +
                        CSV(don.ChiPhiDuKien.ToString(
                            CultureInfo.InvariantCulture)) + "," +
                        CSV(don.TrangThai));
                }

                File.WriteAllText(
                    save.FileName, sb.ToString(),
                    new UTF8Encoding(true));

                MessageBox.Show("Xuất CSV thành công!");
            }
        }

        private string CSV(string giaTri)
        {
            if (giaTri == null)
                giaTri = "";

            return "\"" + giaTri.Replace("\"", "\"\"") + "\"";
        }
    }
}

