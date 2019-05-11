using System;
using System.Collections.Generic;

namespace EStoreTemplatev21.Models
{
    public partial class YeuThich
    {
        public int MaYt { get; set; }
        public int? MaHh { get; set; }
        public string MaKh { get; set; }
        public DateTime? NgayChon { get; set; }
        public string MoTa { get; set; }

        public HangHoa MaHhNavigation { get; set; }
        public KhachHang MaKhNavigation { get; set; }
    }
}
