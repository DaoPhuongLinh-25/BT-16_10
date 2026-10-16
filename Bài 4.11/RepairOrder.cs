using System;
using System.Collections.Generic;
using System.Text;

namespace WinFormsApp11
{
    public class RepairOrder
    {
        public string MaDon { get; set; }
        public string TenKhachHang { get; set; }
        public string SoDienThoai { get; set; }
        public string LoaiThietBi { get; set; }
        public string MoTaLoi { get; set; }
        public DateTime NgayTiepNhan { get; set; }
        public decimal ChiPhiDuKien { get; set; }
        public string TrangThai { get; set; }
        public string DuongDanAnh { get; set; }

        public RepairOrder()
        {
            MaDon = "";
            TenKhachHang = "";
            SoDienThoai = "";
            LoaiThietBi = "Điện thoại";
            MoTaLoi = "";
            NgayTiepNhan = DateTime.Today;
            ChiPhiDuKien = 0;
            TrangThai = "Mới tiếp nhận";
            DuongDanAnh = "";
        }
    }
}
