using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace icloud
{
    internal class ConfigInfo
    {
        // Cau hinh thong tin lien quan ve giao dien chrome tai day
        public static int chrome_width = 390;
        public static int chrome_height = 480;
        public static int chrome_distance_x = 500;
        public static int chrome_distance_y = 600;

        // variable of data
        public static string[] text_comments { get; set; }
        public static string[] image_comments { get; set; }
        public static string[] entity_ids { get; set; }
        public static string[] post_ids { get; set; }



    }

    class ThreadState
    {
        public static bool all_thread_together_running { get; set; }
        public static string proxy { get; set; }
        public static bool allow_running { get; set; }
    }
}
