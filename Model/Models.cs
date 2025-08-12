using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PartsCounter.Model
{
    public class Models
    {
        public class Summary
        {
            public DateTime log_datetime { get; set; }
            public string log_order_no { get; set; }
            public string log_item_code { get; set; }
            public string log_batch_no { get; set; }
            public string log_sublot_no { get; set; }
            public int log_blocks_count { get; set; }
            public int log_actual_count { get; set; }
            public int log_ng_mark { get; set; }
            public int log_unacc { get; set; }
            public string log_reason { get; set; }
            public string log_high_unacc_reason { get; set; }
            public int log_part_counter_no { get; set; }
        }

        public class Breakdown
        {
            public DateTime log_datetime { get; set; }
            public string log_order_no { get; set; }
            public string log_item_code { get; set; }
            public string log_batch_no { get; set; }
            public string log_sublot_no { get; set; }
            public int log_pallet_no { get; set; }
            public int log_actual_count { get; set; }
            public string log_op_number { get; set; }
            public int log_parts_counter_no { get; set; }
            public int summaryID { get; set; }
        }

    }
}
