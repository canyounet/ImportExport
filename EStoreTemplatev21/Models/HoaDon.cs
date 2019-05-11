using System;
using System.Collections.Generic;

namespace EStoreTemplatev21.Models
{
    public partial class HoaDon
    {
        public HoaDon()
        {
            ChiTietHd = new HashSet<ChiTietHd>();
        }

        public int MaHd { get; set; }
        public string MaKh { get; set; }
        public DateTime NgayDat { get; set; }
        public DateTime? NgayCan { get; set; }
        public DateTime? NgayGiao { get; set; }
        public string HoTen { get; set; }
        public string DiaChi { get; set; }
        public string CachThanhToan { get; set; }
        public string CachVanChuyen { get; set; }
        public double PhiVanChuyen { get; set; }
        public int MaTrangThai { get; set; }
        public string MaNv { get; set; }
        public string GhiChu { get; set; }

        public KhachHang MaKhNavigation { get; set; }
        public NhanVien MaNvNavigation { get; set; }
        public TrangThai MaTrangThaiNavigation { get; set; }
        public ICollection<ChiTietHd> ChiTietHd { get; set; }
    }
}
