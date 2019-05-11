using System;
using System.Collections.Generic;

namespace EStoreTemplatev21.Models
{
    public partial class PhanQuyen
    {
        public int MaPq { get; set; }
        public string MaPb { get; set; }
        public int? MaTrang { get; set; }
        public bool Them { get; set; }
        public bool Sua { get; set; }
        public bool Xoa { get; set; }
        public bool Xem { get; set; }

        public PhongBan MaPbNavigation { get; set; }
        public TrangWeb MaTrangNavigation { get; set; }
    }
}
