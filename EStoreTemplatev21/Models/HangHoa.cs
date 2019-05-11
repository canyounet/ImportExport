using System;
using System.Collections.Generic;

namespace EStoreTemplatev21.Models
{
    public partial class HangHoa
    {
        public HangHoa()
        {
            BanBe = new HashSet<BanBe>();
            ChiTietHd = new HashSet<ChiTietHd>();
            YeuThich = new HashSet<YeuThich>();
        }

        public int MaHh { get; set; }
        public string TenHh { get; set; }
        public int MaLoai { get; set; }
        public string MoTaDonVi { get; set; }
        public double? DonGia { get; set; }
        public string Hinh { get; set; }
        public DateTime NgaySx { get; set; }
        public double GiamGia { get; set; }
        public int SoLanXem { get; set; }
        public string MoTa { get; set; }
        public string MaNcc { get; set; }

        public Loai MaLoaiNavigation { get; set; }
        public NhaCungCap MaNccNavigation { get; set; }
        public ICollection<BanBe> BanBe { get; set; }
        public ICollection<ChiTietHd> ChiTietHd { get; set; }
        public ICollection<YeuThich> YeuThich { get; set; }
    }
}
