using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Management;
using System.Net.Http;
using System.Net.Sockets;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using DebtRaven;
using MongoDB.Bson;
using MongoDB.Bson.IO;
using MongoDB.Bson.Serialization.Attributes;
using MongoDB.Driver;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OpenQA.Selenium;
using OpenQA.Selenium.BiDi.Network;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;
using xNet;
using static System.Net.WebRequestMethods;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ListView;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Rebar;
using static icloud.Form1;
using static icloud.Form1.DebtReminderManager;
using JsonConvert = Newtonsoft.Json.JsonConvert;
using LicenseContext = OfficeOpenXml.LicenseContext;

namespace icloud
{
    public partial class Form1 : Form
    {
        private Random random = new Random();
        private string filePath;
        private static bool shouldMerge = false;
        private static bool checkQuaHan = false;
        bool bRunning = false; // dùng biến này để kiểm tra xem chương trình có chạy k
        private static object locker = new object(); // tránh xung đột luồng
        // Dictionary để map mã đơn hàng với index thực tế trong originalData
        private Dictionary<string, int> originalIndexMap = new Dictionary<string, int>();
        // Biến toàn cục để lưu trữ dữ liệu gốc
        private static List<string> originalData = new List<string>();
        private Dictionary<string, DebtReminderInfo> originalDataNew = new Dictionary<string, DebtReminderInfo>();
        private Dictionary<string, string> apiKeys = new Dictionary<string, string>();
        // Dictionary để lưu trạng thái của từng dòng theo mã đơn hàng
        private Dictionary<string, string> statusData = new Dictionary<string, string>();
        private Dictionary<string, string> namePhoneData = new Dictionary<string, string>();
        private AutoUpdater autoUpdater;
        private System.Windows.Forms.Timer updateTimer;
        private Dictionary<int, DateTime> chromeStartTimes = new Dictionary<int, DateTime>();
        private const string LICENSE_API_URL = "https://banhmichaothuongnho.com/bmcthuongnho/";
        private Dictionary<string, string> savedAccounts;
        private Dictionary<string, string> savedAccounts1;
        private DebtReminderManager reminderManager;
        private string accountsFilePath = "accounts.json";
        private string accountsFilePath1 = "accounts1.json";
        Queue<int> qu_position_pro5 = new Queue<int>();
        HttpClient httpClient = new HttpClient();
        public Form1()
        {
            InitializeComponent();
            LoadSavedAccounts();
            LoadSavedAccounts1();
            LoadUsernamesIntoComboBox();
            LoadUsernamesIntoComboBox1();
            reminderManager = new DebtReminderManager();
            SetupDataGridView();
            CheckAndValidateLicense();
            InitializeAutoUpdater();
            comboBox1.SelectedIndex = 0;

        }
        #region thêm cột ngày tháng trong datagridview
        private DateTimePicker dtp = new DateTimePicker();
        private Button btnClear = new Button();

        private void Form1_Load(object sender, EventArgs e)
        {
            // Đăng ký events
            dataGridView1.CellBeginEdit += dataGridView1_CellBeginEdit;
            dataGridView1.CellEndEdit += dataGridView1_CellEndEdit;
            dataGridView1.Scroll += dataGridView1_Scroll;
            dataGridView1.KeyDown += dataGridView1_KeyDown;

            // ===== THÊM EVENTS QUAN TRỌNG ĐỂ COMMIT DỮ LIỆU =====
            dataGridView1.CellValueChanged += dataGridView1_CellValueChanged;
            dataGridView1.CurrentCellDirtyStateChanged += dataGridView1_CurrentCellDirtyStateChanged;
            dataGridView1.RowLeave += dataGridView1_RowLeave;

            // Format cột ngày tháng
            if (dataGridView1.Columns["NgayThang"] != null)
            {
                var column = dataGridView1.Columns["NgayThang"];
                column.ValueType = typeof(DateTime);
                column.DefaultCellStyle.Format = "dd/MM/yyyy";
                column.DefaultCellStyle.NullValue = null;

                // Đảm bảo không có time component
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    if (row.Cells["NgayThang"].Value != null && row.Cells["NgayThang"].Value != DBNull.Value)
                    {
                        if (DateTime.TryParse(row.Cells["NgayThang"].Value.ToString(), out DateTime existingDate))
                        {
                            row.Cells["NgayThang"].Value = existingDate.Date;
                        }
                    }
                }
            }

            // Đảm bảo chỉ chọn từng cell
            dataGridView1.SelectionMode = DataGridViewSelectionMode.CellSelect;
            dataGridView1.MultiSelect = false;
        }

        // ===== THÊM CÁC EVENT HANDLERS ĐỂ COMMIT DỮ LIỆU =====
        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            // Commit ngay khi có thay đổi
            if (sender is DataGridView dgv && e.RowIndex >= 0 && e.ColumnIndex >= 0)
            {
                dgv.CommitEdit(DataGridViewDataErrorContexts.Commit);
                dgv.NotifyCurrentCellDirty(false);
            }
        }

        private void dataGridView1_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            // Commit khi cell dirty
            if (dataGridView1.IsCurrentCellDirty)
            {
                dataGridView1.CommitEdit(DataGridViewDataErrorContexts.Commit);
            }
        }

        private void dataGridView1_RowLeave(object sender, DataGridViewCellEventArgs e)
        {
            // Commit khi rời khỏi row
            dataGridView1.EndEdit();
            if (dataGridView1.DataSource is BindingSource bs)
            {
                bs.EndEdit();
            }
        }

        private void dataGridView1_KeyDown(object sender, KeyEventArgs e)
        {
            if (dataGridView1.CurrentCell != null &&
                dataGridView1.Columns[dataGridView1.CurrentCell.ColumnIndex].Name == "NgayThang" &&
                (e.KeyCode == System.Windows.Forms.Keys.Delete || e.KeyCode == System.Windows.Forms.Keys.Back))
            {
                // Chỉ xóa cell hiện tại, không xóa toàn bộ cột
                dataGridView1[dataGridView1.CurrentCell.ColumnIndex, dataGridView1.CurrentCell.RowIndex].Value = DBNull.Value;

                // ===== QUAN TRỌNG: COMMIT SAU KHI XÓA =====
                dataGridView1.CommitEdit(DataGridViewDataErrorContexts.Commit);
                dataGridView1.NotifyCurrentCellDirty(false);

                e.Handled = true;
            }
        }

        private void dataGridView1_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            if (dataGridView1.Columns[e.ColumnIndex].Name == "NgayThang")
            {
                Rectangle rect = dataGridView1.GetCellDisplayRectangle(e.ColumnIndex, e.RowIndex, false);

                // Thiết lập DateTimePicker
                dtp.Size = new Size(rect.Width - 25, rect.Height);
                dtp.Location = rect.Location;
                dtp.Format = DateTimePickerFormat.Short;

                // Thiết lập nút Clear
                btnClear.Size = new Size(25, rect.Height);
                btnClear.Location = new Point(rect.X + rect.Width - 25, rect.Y);
                btnClear.Text = "X";
                btnClear.Font = new Font("Arial", 8, FontStyle.Bold);
                btnClear.BackColor = Color.LightCoral;
                btnClear.ForeColor = Color.White;
                btnClear.FlatStyle = FlatStyle.Flat;

                // Gán giá trị hiện tại
                var cellValue = dataGridView1[e.ColumnIndex, e.RowIndex].Value;
                if (cellValue != null && cellValue != DBNull.Value && !string.IsNullOrEmpty(cellValue.ToString()))
                {
                    DateTime dateValue;
                    if (DateTime.TryParse(cellValue.ToString(), out dateValue))
                    {
                        dtp.Value = dateValue;
                    }
                    else
                    {
                        dtp.Value = DateTime.Now;
                    }
                }
                else
                {
                    dtp.Value = DateTime.Now;
                }

                // Gỡ bỏ event cũ trước khi thêm mới (tránh duplicate events)
                btnClear.Click -= BtnClear_Click;
                btnClear.Click += BtnClear_Click;

                // ===== THÊM EVENT CHO DATETIMEPICKER =====
                dtp.ValueChanged -= Dtp_ValueChanged;
                dtp.ValueChanged += Dtp_ValueChanged;

                dataGridView1.Controls.Add(dtp);
                dataGridView1.Controls.Add(btnClear);
                dtp.BringToFront();
                btnClear.BringToFront();

                // SỬA: Sử dụng Invoke thay vì BeginInvoke để đồng bộ hơn
                this.Invoke(new Action(() => dtp.Focus()));
            }
        }

        // ===== THÊM EVENT CHO DATETIMEPICKER =====
        private void Dtp_ValueChanged(object sender, EventArgs e)
        {
            if (dataGridView1.CurrentCell != null)
            {
                int currentRow = dataGridView1.CurrentCell.RowIndex;
                int currentCol = dataGridView1.CurrentCell.ColumnIndex;

                // Cập nhật giá trị ngay
                dataGridView1[currentCol, currentRow].Value = dtp.Value.Date;
                dataGridView1.CommitEdit(DataGridViewDataErrorContexts.Commit);
                dataGridView1.NotifyCurrentCellDirty(false);
            }
        }

        // Event handler cho nút Clear
        private void BtnClear_Click(object sender, EventArgs e)
        {
            if (dataGridView1.CurrentCell != null)
            {
                int currentRow = dataGridView1.CurrentCell.RowIndex;
                int currentCol = dataGridView1.CurrentCell.ColumnIndex;

                dataGridView1[currentCol, currentRow].Value = DBNull.Value;

                // ===== QUAN TRỌNG: COMMIT SAU KHI CLEAR =====
                dataGridView1.CommitEdit(DataGridViewDataErrorContexts.Commit);
                dataGridView1.NotifyCurrentCellDirty(false);

                dataGridView1.Controls.Remove(dtp);
                dataGridView1.Controls.Remove(btnClear);
                dataGridView1.EndEdit();
            }
        }

        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            // ===== XỬ LÝ TẤT CẢ CÁC CỘT, KHÔNG CHỈ NgayThang =====
            if (dataGridView1.Columns[e.ColumnIndex].Name == "NgayThang")
            {
                if (dataGridView1.Controls.Contains(dtp))
                {
                    // Chỉ lấy phần ngày tháng năm, bỏ giờ phút giây
                    dataGridView1[e.ColumnIndex, e.RowIndex].Value = dtp.Value.Date;
                    dataGridView1.Controls.Remove(dtp);
                }
                if (dataGridView1.Controls.Contains(btnClear))
                {
                    dataGridView1.Controls.Remove(btnClear);
                }
            }

            // ===== QUAN TRỌNG: COMMIT CHO TẤT CẢ CÁC CỘT =====
            dataGridView1.CommitEdit(DataGridViewDataErrorContexts.Commit);
            dataGridView1.NotifyCurrentCellDirty(false);

            // Nếu có DataSource, commit luôn
            if (dataGridView1.DataSource is BindingSource bs)
            {
                bs.EndEdit();
            }
            else if (dataGridView1.DataSource is DataTable dt)
            {
                dt.AcceptChanges();
            }
        }

        private void dataGridView1_Scroll(object sender, ScrollEventArgs e)
        {
            // ===== COMMIT TRƯỚC KHI REMOVE CONTROLS =====
            if (dataGridView1.Controls.Contains(dtp))
            {
                if (dataGridView1.CurrentCell != null)
                {
                    int currentRow = dataGridView1.CurrentCell.RowIndex;
                    int currentCol = dataGridView1.CurrentCell.ColumnIndex;
                    dataGridView1[currentCol, currentRow].Value = dtp.Value.Date;
                    dataGridView1.CommitEdit(DataGridViewDataErrorContexts.Commit);
                }
                dataGridView1.Controls.Remove(dtp);
            }
            if (dataGridView1.Controls.Contains(btnClear))
                dataGridView1.Controls.Remove(btnClear);
        }
        #endregion

        private static readonly Dictionary<string, string> ModelMappings = new Dictionary<string, string>
    {
        // iPhone models
        {"ip3", "IPHONE 3"},
        {"ip11", "IPHONE 11"},
        {"ip12", "IPHONE 12"},
        {"ip13", "IPHONE 13"},
        {"ip14", "IPHONE 14"},
        {"ip15", "IPHONE 15"},
        {"ip16", "IPHONE 16"},
        
        // Pro variants
        {"ip11pro", "IPHONE 11 PRO"},
        {"ip12pro", "IPHONE 12 PRO"},
        {"ip13pro", "IPHONE 13 PRO"},
        {"ip14pro", "IPHONE 14 PRO"},
        {"ip15pro", "IPHONE 15 PRO"},
        {"ip16pro", "IPHONE 16 PRO"},
        
        // Pro Max variants
        {"ip11prm", "IPHONE 11 PRO MAX"},
        {"ip11promax", "IPHONE 11 PRO MAX"},
        {"ip12prm", "IPHONE 12 PRO MAX"},
        {"ip12promax", "IPHONE 12 PRO MAX"},
        {"ip13prm", "IPHONE 13 PRO MAX"},
        {"ip13promax", "IPHONE 13 PRO MAX"},
        {"ip14prm", "IPHONE 14 PRO MAX"},
        {"ip14promax", "IPHONE 14 PRO MAX"},
        {"ip15prm", "IPHONE 15 PRO MAX"},
        {"ip15promax", "IPHONE 15 PRO MAX"},
        {"ip16prm", "IPHONE 16 PRO MAX"},
        {"ip16promax", "IPHONE 16 PRO MAX"},
        
        // Plus variants
        {"ip12plus", "IPHONE 12 PLUS"},
        {"ip13plus", "IPHONE 13 PLUS"},
        {"ip14plus", "IPHONE 14 PLUS"},
        {"ip15plus", "IPHONE 15 PLUS"},
        {"ip16plus", "IPHONE 16 PLUS"},
        
        // Mini variants
        {"ip12mini", "IPHONE 12 MINI"},
        {"ip13mini", "IPHONE 13 MINI"},
        
        // Special cases
        {"ip16e", "IPHONE 16E"},
        
        // Pro variants with different notation
        {"ip11p", "IPHONE 11 PRO"},
        {"ip12p", "IPHONE 12 PRO"},
        {"ip13p", "IPHONE 13 PRO"},
        {"ip14p", "IPHONE 14 PRO"},
        {"ip15p", "IPHONE 15 PRO"},
        {"ip16p", "IPHONE 16 PRO"},
        
        // Pro Max with p notation
        {"ip11pm", "IPHONE 11 PRO MAX"},
        {"ip12pm", "IPHONE 12 PRO MAX"},
        {"ip13pm", "IPHONE 13 PRO MAX"},
        {"ip14pm", "IPHONE 14 PRO MAX"},
        {"ip15pm", "IPHONE 15 PRO MAX"},
        {"ip16pm", "IPHONE 16 PRO MAX"}
    };
        public static string GetIPhoneName(string customerInfo)
        {
            if (string.IsNullOrWhiteSpace(customerInfo))
                return "UNKNOWN";

            // Làm sạch và chuẩn hóa chuỗi đầu vào
            string cleanInput = customerInfo.ToLower().Trim();

            // Regex để tìm pattern iPhone trong chuỗi
            // Tìm "ip" theo sau bởi số và có thể có thêm các từ như pro, max, mini, plus
            string pattern = @"ip\s*(\d+)\s*([a-z]*)\s*([a-z]*)";
            Match match = Regex.Match(cleanInput, pattern);

            if (!match.Success)
            {
                // Nếu không match được pattern chính, thử tìm trực tiếp trong dictionary
                foreach (var mapping in ModelMappings)
                {
                    if (cleanInput.Contains(mapping.Key))
                    {
                        return mapping.Value;
                    }
                }
                return "UNKNOWN";
            }

            // Xây dựng key để tìm trong dictionary
            string modelNumber = match.Groups[1].Value;
            string variant1 = match.Groups[2].Value.Trim();
            string variant2 = match.Groups[3].Value.Trim();

            // Thử các combination khác nhau
            List<string> possibleKeys = new List<string>();

            // Thêm các khả năng với variant
            if (!string.IsNullOrEmpty(variant1) && !string.IsNullOrEmpty(variant2))
            {
                possibleKeys.Add($"ip{modelNumber}{variant1}{variant2}");
                possibleKeys.Add($"ip{modelNumber}{variant2}{variant1}");
            }

            if (!string.IsNullOrEmpty(variant1))
            {
                possibleKeys.Add($"ip{modelNumber}{variant1}");
            }

            // Thêm model cơ bản
            possibleKeys.Add($"ip{modelNumber}");

            // Tìm trong dictionary
            foreach (string key in possibleKeys)
            {
                if (ModelMappings.ContainsKey(key))
                {
                    return ModelMappings[key];
                }
            }

            // Nếu không tìm thấy, trả về model cơ bản
            return $"IPHONE {modelNumber}";
        }

        private async Task CheckAndValidateLicense()
        {
            string pathKey = "key.txt";

            // Kiểm tra xem file key.txt có tồn tại không
            if (System.IO.File.Exists(pathKey))
            {
                string storedKey = System.IO.File.ReadAllText(pathKey).Trim();

                if (!string.IsNullOrEmpty(storedKey))
                {
                    var licenseStatus = await CheckLicenseAsync();

                    if (licenseStatus.IsActive && licenseStatus.key == storedKey)
                    {
                        return; // License hợp lệ
                    }
                }
            }

            // Kiểm tra license hoặc yêu cầu mới
            var newLicenseStatus = await CheckLicenseAsync();

            if (newLicenseStatus.IsActive && !string.IsNullOrEmpty(newLicenseStatus.key))
            {
                System.IO.File.WriteAllText(pathKey, newLicenseStatus.key);
            }
            else if (newLicenseStatus.Status == "PENDING")
            {
                MessageBox.Show("⏳ License đang chờ duyệt", "Thông Báo");
                DisableFeatures();
            }
            else if (newLicenseStatus.Status == "REJECTED")
            {
                MessageBox.Show("❌ Yêu cầu license đã bị từ chối", "Thông Báo");
                DisableFeatures();
            }
            else
            {
                var requestResult = await RequestLicenseAsync();
                if (requestResult.Success)
                {
                    MessageBox.Show($"⏳ {requestResult.Message}", "Thông Báo");
                    DisableFeatures();
                }
            }
        }
        // Lấy mã máy tính duy nhất
        private static string _cachedMachineId = null;
        private static readonly object _lockObject = new object();

        public string GetMachineId()
        {
            lock (_lockObject)
            {
                if (!string.IsNullOrEmpty(_cachedMachineId))
                    return _cachedMachineId;

                try
                {
                    // CHỈ SỬ DỤNG CÁC THÔNG SỐ CỰC KỲ ỔN ĐỊNH
                    List<string> stableIdentifiers = new List<string>();

                    // 1. System UUID - Ưu tiên số 1 (gần như không bao giờ đổi)
                    string systemUuid = GetSystemUuid();
                    if (IsValidStableIdentifier(systemUuid))
                    {
                        stableIdentifiers.Add($"UUID:{systemUuid}");
                        Console.WriteLine($"✓ System UUID found: {systemUuid}");
                    }

                    // 2. BIOS Serial - Ưu tiên số 2 (rất ổn định)
                    string biosSerial = GetBiosSerial();
                    if (IsValidStableIdentifier(biosSerial))
                    {
                        stableIdentifiers.Add($"BIOS:{biosSerial}");
                        Console.WriteLine($"✓ BIOS Serial found: {biosSerial}");
                    }

                    // 3. CPU ProcessorID - Ưu tiên số 3 (ổn định trên hầu hết máy)
                    string cpuId = GetCpuIdSafe();
                    if (IsValidStableIdentifier(cpuId))
                    {
                        stableIdentifiers.Add($"CPU:{cpuId}");
                        Console.WriteLine($"✓ CPU ID found: {cpuId}");
                    }

                    // Kiểm tra xem có đủ identifiers ổn định không
                    if (stableIdentifiers.Count == 0)
                    {
                        throw new InvalidOperationException("Không tìm thấy identifier ổn định nào. Máy này có thể là máy ảo hoặc có vấn đề với WMI.");
                    }

                    // Nếu chỉ có 1 identifier, cảnh báo nhưng vẫn tiếp tục
                    if (stableIdentifiers.Count == 1)
                    {
                        Console.WriteLine("⚠️ WARNING: Chỉ tìm thấy 1 stable identifier. Machine ID có thể không hoàn toàn unique.");
                    }

                    // Sắp xếp để đảm bảo thứ tự nhất quán
                    stableIdentifiers.Sort();
                    string combined = string.Join("|", stableIdentifiers);

                    _cachedMachineId = GenerateStableHash(combined);

                    Console.WriteLine($"Machine ID generated from: {combined}");
                    Console.WriteLine($"Final Machine ID: {_cachedMachineId}");
                    Console.WriteLine($"Sử dụng {stableIdentifiers.Count} stable identifiers");

                    return _cachedMachineId;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"❌ Error generating machine ID: {ex.Message}");

                    // FALLBACK CUỐI CÙNG - chỉ sử dụng khi hoàn toàn không có gì khác
                    // Sử dụng thông tin ít thay đổi nhất có thể
                    var fallbackData = $"FALLBACK:{Environment.OSVersion.Platform}:{Environment.ProcessorCount}:{Environment.SystemDirectory}";
                    _cachedMachineId = GenerateStableHash(fallbackData);

                    Console.WriteLine($"⚠️ Using emergency fallback Machine ID");
                    return _cachedMachineId;
                }
            }
        }

        // Kiểm tra tính hợp lệ CỰC KỲ NGHIÊM NGẶT
        private bool IsValidStableIdentifier(string identifier)
        {
            if (string.IsNullOrWhiteSpace(identifier) || identifier.Length < 4)
                return false;

            // Danh sách đen các giá trị không hợp lệ (mở rộng)
            var invalidValues = new[]
            {
        // Generic invalid values
        "UNKNOWN_CPU", "UNKNOWN_MB", "UNKNOWN_DISK", "UNKNOWN", "NOT_FOUND",
        
        // OEM placeholders
        "To be filled by O.E.M.", "System Serial Number", "Default string",
        "Not Specified", "None", "N/A", "NULL", "TBD",
        
        // Common placeholder patterns
        "0000000000000000", "FFFFFFFFFFFFFFFF", "123456789", "ABCDEFGH",
        "00000000-0000-0000-0000-000000000000",
        
        // Virtual machine indicators
        "VMware", "VBOX", "VirtualBox", "QEMU", "Microsoft Corporation"
    };

            string upperIdentifier = identifier.ToUpperInvariant();

            // Kiểm tra exact match
            if (invalidValues.Any(invalid => upperIdentifier.Equals(invalid.ToUpperInvariant())))
                return false;

            // Kiểm tra contains (cho các pattern)
            if (invalidValues.Any(invalid => upperIdentifier.Contains(invalid.ToUpperInvariant()) && invalid.Length > 5))
                return false;

            // Kiểm tra pattern toàn số 0 hoặc F
            if (identifier.All(c => c == '0' || c == 'F' || c == '-'))
                return false;

            // Kiểm tra độ dài tối thiểu cho từng loại
            if (identifier.Length < 6)
                return false;

            return true;
        }

        private string GetSystemUuid()
        {
            try
            {
                using (var mc = new ManagementClass("Win32_ComputerSystemProduct"))
                using (var moc = mc.GetInstances())
                {
                    foreach (ManagementObject mo in moc)
                    {
                        try
                        {
                            string uuid = mo.Properties["UUID"]?.Value?.ToString();
                            if (!string.IsNullOrWhiteSpace(uuid))
                            {
                                // Kiểm tra UUID có định dạng chuẩn không
                                if (Guid.TryParse(uuid, out Guid guidResult))
                                {
                                    // Loại bỏ nil UUID
                                    if (guidResult != Guid.Empty)
                                    {
                                        return uuid.Trim().ToUpperInvariant();
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Error processing UUID: {ex.Message}");
                        }
                        finally
                        {
                            mo?.Dispose();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error getting System UUID: {ex.Message}");
            }
            return null;
        }

        private string GetBiosSerial()
        {
            try
            {
                using (var mc = new ManagementClass("Win32_BIOS"))
                using (var moc = mc.GetInstances())
                {
                    foreach (ManagementObject mo in moc)
                    {
                        try
                        {
                            string serialNumber = mo.Properties["SerialNumber"]?.Value?.ToString();
                            if (!string.IsNullOrWhiteSpace(serialNumber))
                            {
                                return serialNumber.Trim().ToUpperInvariant();
                            }
                        }
                        finally
                        {
                            mo?.Dispose();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error getting BIOS Serial: {ex.Message}");
            }
            return null;
        }

        private string GetCpuIdSafe()
        {
            try
            {
                using (var mc = new ManagementClass("Win32_Processor"))
                using (var moc = mc.GetInstances())
                {
                    foreach (ManagementObject mo in moc)
                    {
                        try
                        {
                            string processorId = mo.Properties["ProcessorID"]?.Value?.ToString();
                            if (!string.IsNullOrWhiteSpace(processorId))
                            {
                                return processorId.Trim().ToUpperInvariant();
                            }
                        }
                        finally
                        {
                            mo?.Dispose();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error getting CPU ID: {ex.Message}");
            }
            return null;
        }

        private string GenerateStableHash(string input)
        {
            using (var sha256Hash = SHA256.Create())
            {
                // Thêm salt để tăng security
                string saltedInput = $"MachineID_v2.0_{input}_Salt2024";
                byte[] bytes = sha256Hash.ComputeHash(Encoding.UTF8.GetBytes(saltedInput));

                // Sử dụng Base32 để tránh ký tự đặc biệt
                return ConvertToBase32(bytes).Substring(0, 20);
            }
        }

        private string ConvertToBase32(byte[] bytes)
        {
            const string alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ234567";
            StringBuilder result = new StringBuilder();

            for (int i = 0; i < bytes.Length; i += 5)
            {
                int byteCount = Math.Min(5, bytes.Length - i);
                ulong buffer = 0;

                for (int j = 0; j < byteCount; j++)
                {
                    buffer = (buffer << 8) | bytes[i + j];
                }

                int bitCount = byteCount * 8;
                while (bitCount > 0)
                {
                    int index = (int)((buffer >> Math.Max(bitCount - 5, 0)) & 0x1F);
                    result.Append(alphabet[index]);
                    bitCount -= 5;
                }
            }

            return result.ToString();
        }
        public async Task<ApiResponse> RequestLicenseAsync()
        {
            try
            {
                string machineId = GetMachineId();
                if (string.IsNullOrEmpty(machineId))
                    return new ApiResponse { Success = false, Message = "Không thể lấy mã máy tính" };

                var requestData = new
                {
                    action = "request",
                    MachineId = machineId,
                    ComputerName = Environment.MachineName,
                    UserName = Environment.UserName,
                    RequestTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
                };

                string jsonData = JsonConvert.SerializeObject(requestData);
                var content = new System.Net.Http.StringContent(jsonData, Encoding.UTF8, "application/json");

                HttpResponseMessage response = await httpClient.PostAsync($"{LICENSE_API_URL}/license2.php", content);
                string responseContent = await response.Content.ReadAsStringAsync();

                // In nội dung phản hồi để debug
                Console.WriteLine($"Server response: {responseContent}");

                if (response.IsSuccessStatusCode)
                {
                    var result = JsonConvert.DeserializeObject<LicenseApiResponse>(responseContent);
                    return new ApiResponse
                    {
                        Success = result.Success,
                        Message = result.Message,
                        Data = result.Data
                    };
                }
                else
                {
                    // Ghi log lỗi chi tiết
                    Console.WriteLine($"HTTP Error: {response.StatusCode}, Content: {responseContent}");
                    return new ApiResponse
                    {
                        Success = false,
                        Message = $"Lỗi server: {response.StatusCode} - {responseContent}"
                    };
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Exception: {ex.Message}");
                return new ApiResponse
                {
                    Success = false,
                    Message = $"Lỗi kết nối: {ex.Message}"
                };
            }
        }
        public class ApiResponse
        {
            public bool Success { get; set; }
            public string Message { get; set; }
            public object Data { get; set; }
        }

        public class LicenseStatus
        {
            public bool IsActive { get; set; }
            public string Status { get; set; }
            public string key { get; set; }
            public DateTime? ExpiryDate { get; set; }
        }


        public async Task<LicenseStatus> CheckLicenseAsync()
        {
            try
            {
                string machineId = GetMachineId();
                if (string.IsNullOrEmpty(machineId))
                    return new LicenseStatus { IsActive = false, Status = "ERROR", key = "Không thể lấy mã máy tính" };

                HttpResponseMessage response = await httpClient.GetAsync($"{LICENSE_API_URL}/license2.php?action=check_license&machineId={machineId}");

                if (response.IsSuccessStatusCode)
                {
                    string responseContent = await response.Content.ReadAsStringAsync();
                    var licenseResponse = JsonConvert.DeserializeObject<LicenseApiResponse>(responseContent);

                    return new LicenseStatus
                    {
                        IsActive = licenseResponse.Success && licenseResponse.Status.ToUpper() == "APPROVED",
                        Status = licenseResponse.Status.ToUpper(),
                        key = licenseResponse.Data?.LicenseKey ?? ""
                    };
                }

                return new LicenseStatus
                {
                    IsActive = false,
                    Status = "ERROR",
                    key = "Không thể kết nối đến server license"
                };
            }
            catch (Exception ex)
            {
                return new LicenseStatus
                {
                    IsActive = false,
                    Status = "ERROR",
                    key = $"Lỗi kiểm tra license: {ex.Message}"
                };
            }
        }
        public class LicenseApiResponse
        {
            public bool Success { get; set; }
            public string Message { get; set; }
            public string Status { get; set; }
            public LicenseData Data { get; set; }
        }

        public class LicenseData
        {
            public string LicenseKey { get; set; }
            public bool IsActive { get; set; }
            public string ExpiryDate { get; set; }
            public int ActivationCount { get; set; }
            public int ActivationLimit { get; set; }
            public int DaysRemaining { get; set; }
            public string ComputerName { get; set; }
            public string UserName { get; set; }
        }

        #region update
        private void InitializeAutoUpdater()
        {
            // Khởi tạo SimpleAutoUpdater với URL server của bạn
            autoUpdater = new AutoUpdater("https://banhmichaothuongnho.com/updates");

            // Thiết lập timer để kiểm tra cập nhật định kỳ (mỗi 30 phút)
            updateTimer = new System.Windows.Forms.Timer();
            updateTimer.Interval = 30 * 60 * 1000; // 30 phút
            updateTimer.Tick += async (s, e) => await autoUpdater.CheckAndUpdateAsync();
            updateTimer.Start();

            // Kiểm tra cập nhật ngay khi khởi động (sau 10 giây)
            var startupTimer = new System.Windows.Forms.Timer();
            startupTimer.Interval = 3000; // 10 giây
            startupTimer.Tick += async (s, e) =>
            {
                startupTimer.Stop();
                startupTimer.Dispose();
                await autoUpdater.CheckAndUpdateAsync();
            };
            startupTimer.Start();
        }
        public class AutoUpdater
        {
            private readonly string updateUrl;
            private readonly string currentVersion;
            private readonly string appDirectory;
            private readonly string currentExePath;

            public AutoUpdater(string updateUrl)
            {
                this.updateUrl = updateUrl;
                this.currentVersion = Assembly.GetExecutingAssembly().GetName().Version.ToString();
                this.currentExePath = Assembly.GetExecutingAssembly().Location;
                this.appDirectory = Path.GetDirectoryName(currentExePath);
            }

            // Kiểm tra và thực hiện cập nhật tự động
            public async Task CheckAndUpdateAsync()
            {
                try
                {
                    var updateInfo = await CheckForUpdatesAsync();
                    if (updateInfo != null && IsNewerVersion(updateInfo.Version))
                    {
                        await DownloadAndReplaceAsync(updateInfo);
                    }
                }
                catch (Exception ex)
                {
                    // Log error nhưng không hiện thông báo
                    try
                    {
                        System.IO.File.WriteAllText(
                            Path.Combine(appDirectory, "update_error.log"),
                            $"{DateTime.Now}: {ex.Message}\n"
                        );
                    }
                    catch { }
                }
            }

            // Kiểm tra phiên bản mới
            private async Task<UpdateInfo> CheckForUpdatesAsync()
            {
                using (var client = new HttpClient())
                {
                    client.Timeout = TimeSpan.FromSeconds(10);
                    var response = await client.GetStringAsync($"{updateUrl}/version.json");
                    return System.Text.Json.JsonSerializer.Deserialize<UpdateInfo>(response);
                }
            }

            // So sánh phiên bản
            private bool IsNewerVersion(string newVersion)
            {
                try
                {
                    var current = new Version(currentVersion);
                    var newer = new Version(newVersion);
                    return newer > current;
                }
                catch
                {
                    return false;
                }
            }

            // Tải và thay thế file exe
            private async Task DownloadAndReplaceAsync(UpdateInfo updateInfo)
            {
                var tempUpdateFile = Path.Combine(appDirectory, "DebtRaven_update.exe");

                try
                {
                    // Tải file mới về
                    using (var client = new HttpClient())
                    {
                        var fileBytes = await client.GetByteArrayAsync(updateInfo.DownloadUrl);
                        System.IO.File.WriteAllBytes(tempUpdateFile, fileBytes);
                    }

                    // Tạo file updater script
                    var updaterScript = CreateUpdaterScript(tempUpdateFile);

                    // Chạy updater và thoát
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = "cmd.exe",
                        Arguments = $"/c \"{updaterScript}\"",
                        WindowStyle = ProcessWindowStyle.Hidden,
                        CreateNoWindow = true
                    });

                    Application.Exit();
                }
                catch
                {
                    // Dọn dẹp nếu có lỗi
                    try
                    {
                        if (System.IO.File.Exists(tempUpdateFile))
                            System.IO.File.Delete(tempUpdateFile);
                    }
                    catch { }
                }
            }
            private string CreateUpdaterScript(string tempFile)
            {
                var currentExeName = Path.GetFileName(currentExePath);
                var scriptPath = Path.Combine(appDirectory, "updater.bat");

                var scriptContent = $@"
@echo off
cd /d ""{appDirectory}""
timeout /t 2 /nobreak > nul

REM Dừng process nếu đang chạy
taskkill /f /im ""{currentExeName}"" > nul 2>&1
timeout /t 1 /nobreak > nul

REM Xóa file cũ và copy file mới
del ""{currentExeName}"" > nul 2>&1
move ""{Path.GetFileName(tempFile)}"" ""{currentExeName}"" > nul 2>&1

REM Khởi động lại ứng dụng
timeout /t 1 /nobreak > nul
start """" ""{currentExeName}""

REM Dọn dẹp script
timeout /t 3 /nobreak > nul
del ""{Path.GetFileName(scriptPath)}"" > nul 2>&1
";

                System.IO.File.WriteAllText(scriptPath, scriptContent);
                return scriptPath;
            }
        }

        
        public class UpdateInfo
        {
            public string Version { get; set; }
            public string DownloadUrl { get; set; }
            
        }


        #endregion

        private void SetupDataGridView()
        {
            // Thiết lập DataGridView
            dataGridView1.EditMode = DataGridViewEditMode.EditOnEnter;
            dataGridView1.SelectionMode = DataGridViewSelectionMode.CellSelect;
            dataGridView1.ReadOnly = false;
        }
        public class DebtReminderManager
        {
            private string filePath;
            private Dictionary<string, DebtReminderInfo> debtData;
            public class DebtReminderInfo
            {
                public string MaHD { get; set; }
                public string TenKh { get; set; }
                public string userIcloud { get; set; }
                public string PassIcloud { get; set; }
                public string NamePhoneChange { get; set; }
                public string Note { get; set; }
                public string Line { get; set; }
                public string TenSo { get; set; }

                public string date { get; set; }
            }

            public class dataRungIcloud
            {
                [BsonId]
                public ObjectId Id { get; set; } = ObjectId.GenerateNewId();
                public string Name { get; set; }
                public string Data { get; set; }
            }
            public DebtReminderManager(string jsonFilePath = "DebtReminderData.json")
            {
                filePath = jsonFilePath;
                LoadDataFromJson();
            }

            // Load dữ liệu từ file JSON
            private void LoadDataFromJson()
            {
                try
                {
                    if (System.IO.File.Exists(filePath))
                    {
                        
                        string jsonContent = System.IO.File.ReadAllText(filePath);
                        debtData = JsonConvert.DeserializeObject<Dictionary<string, DebtReminderInfo>>(jsonContent)
                                  ?? new Dictionary<string, DebtReminderInfo>();
                    }
                    else
                    {
                        debtData = new Dictionary<string, DebtReminderInfo>();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Lỗi khi load dữ liệu: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    debtData = new Dictionary<string, DebtReminderInfo>();
                }
            }

            // Lưu dữ liệu vào file JSON
            public bool SaveDataToJson()
            {
                try
                {

                    string jsonContent = JsonConvert.SerializeObject(debtData, Formatting.Indented);

                    System.IO.File.WriteAllText(filePath, jsonContent);
                    return true;
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Lỗi khi lưu dữ liệu: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
            }
            // Xóa file JSON cũ
            public bool DeleteOldJsonFile()
            {
                try
                {
                    if (System.IO.File.Exists(filePath))
                    {
                        System.IO.File.Delete(filePath);
                        return true;
                    }
                    return true; // File không tồn tại cũng coi như thành công
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Lỗi khi xóa file cũ: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
            }
           
            
            // Load thông tin nhắc nợ vào DataGridView
            public void LoadReminderData(DataGridView dataGridView)
            {
                try
                {
                    foreach (object obj in ((IEnumerable)dataGridView.Rows))
                    {
                        DataGridViewRow dataGridViewRow = (DataGridViewRow)obj;
                        bool isNewRow = dataGridViewRow.IsNewRow;
                        if (!isNewRow)
                        {
                            DataGridViewCell dataGridViewCell = dataGridViewRow.Cells["MaHD"];
                            string text;
                            if (dataGridViewCell == null)
                            {
                                text = null;
                            }
                            else
                            {
                                object value = dataGridViewCell.Value;
                                text = ((value != null) ? value.ToString() : null);
                            }
                            string text2 = text;
                            bool flag = string.IsNullOrEmpty(text2);
                            if (!flag)
                            {
                                bool flag2 = this.debtData.ContainsKey(text2);
                                if (flag2)
                                {
                                    Form1.DebtReminderManager.DebtReminderInfo debtReminderInfo = this.debtData[text2];
                                    bool flag3 = dataGridViewRow.Cells[1] != null;
                                    if (flag3)
                                    {
                                        dataGridViewRow.Cells[1].Value = debtReminderInfo.TenKh;
                                    }
                                    bool flag4 = dataGridViewRow.Cells[10] != null;
                                    if (flag4)
                                    {
                                        dataGridViewRow.Cells[10].Value = debtReminderInfo.userIcloud;
                                    }
                                    bool flag5 = dataGridViewRow.Cells[11] != null;
                                    if (flag5)
                                    {
                                        dataGridViewRow.Cells[11].Value = debtReminderInfo.PassIcloud;
                                    }
                                    bool flag6 = dataGridViewRow.Cells[9] != null;
                                    if (flag6)
                                    {
                                        dataGridViewRow.Cells[9].Value = debtReminderInfo.NamePhoneChange;
                                    }
                                    bool flag7 = dataGridViewRow.Cells[3] != null;
                                    if (flag7)
                                    {
                                        dataGridViewRow.Cells[3].Value = debtReminderInfo.Note;
                                    }
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Lỗi khi load dữ liệu nhắc nợ: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                }
            }
            // Lấy thông tin nhắc theo Mã HĐ
            public DebtReminderInfo GetReminderByMaHD(string maHD)
            {
                return debtData.ContainsKey(maHD) ? debtData[maHD] : null;
            }

            // Xóa thông tin nhắc theo Mã HĐ
            public bool DeleteReminder(string maHD)
            {
                if (debtData.ContainsKey(maHD))
                {
                    debtData.Remove(maHD);
                    return SaveDataToJson();
                }
                return false;
            }

            // Lấy tất cả dữ liệu
            public Dictionary<string, DebtReminderInfo> GetAllData()
            {
                return new Dictionary<string, DebtReminderInfo>(debtData);
            }

            public bool SaveReminderData(DataGridView dataGridView, string line, bool deleteOldFile = true)
            {
                bool result;
                try
                {
                    // Commit các thay đổi trong DataGridView
                    dataGridView.CommitEdit(DataGridViewDataErrorContexts.Commit);
                    dataGridView.NotifyCurrentCellDirty(false);
                    Application.DoEvents();

                    // Xóa file JSON cũ nếu cần
                    if (deleteOldFile)
                    {
                        bool deleteSuccess = this.DeleteOldJsonFile();
                        if (!deleteSuccess)
                        {
                            return false;
                        }
                        this.debtData.Clear();
                    }

                    // Duyệt qua từng dòng trong DataGridView
                    foreach (object rowObj in ((IEnumerable)dataGridView.Rows))
                    {
                        DataGridViewRow row = (DataGridViewRow)rowObj;

                        // Bỏ qua dòng mới (dòng trống để thêm dữ liệu)
                        if (row.IsNewRow)
                        {
                            continue;
                        }

                        // Lấy mã hợp đồng từ cột "MaHD"
                        DataGridViewCell contractCodeCell = row.Cells["MaHD"];
                        string contractCode = contractCodeCell?.Value?.ToString();

                        // Bỏ qua nếu không có mã hợp đồng
                        if (string.IsNullOrEmpty(contractCode))
                        {
                            continue;
                        }

                        // Lấy tên khách hàng từ cột 1
                        DataGridViewCell customerNameCell = row.Cells[1];
                        string customerName = customerNameCell?.Value?.ToString() ?? "";

                        // Tạo key duy nhất từ mã hợp đồng và tên khách hàng
                        string uniqueKey = contractCode + "|" + customerName;

                        // Xử lý ngày tháng từ cột 14
                        string formattedDate = "";
                        DataGridViewCell dateCell = row.Cells[14];
                        object dateValue = dateCell?.Value;

                        if (dateValue != null && dateValue != DBNull.Value)
                        {
                            // Nếu giá trị là kiểu DateTime
                            if (dateValue is DateTime)
                            {
                                DateTime dateTime = (DateTime)dateValue;
                                formattedDate = dateTime.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture);
                            }
                            else
                            {
                                // Nếu giá trị là chuỗi, cần parse
                                string dateString = dateValue.ToString().Trim();

                                if (!string.IsNullOrEmpty(dateString))
                                {
                                    // Các định dạng ngày có thể có
                                    string[] dateFormats = new string[]
                                    {
                            "dd/MM/yyyy", "d/MM/yyyy", "dd/M/yyyy", "d/M/yyyy",
                            "dd-MM-yyyy", "d-MM-yyyy", "dd-M-yyyy", "d-M-yyyy",
                            "yyyy-MM-dd",
                            "dd/MM/yy", "d/MM/yy", "dd/M/yy", "d/M/yy"
                                    };

                                    bool dateParseSuccess = false;

                                    // Thử parse với các định dạng có sẵn
                                    foreach (string format in dateFormats)
                                    {
                                        DateTime parsedDate;
                                        if (DateTime.TryParseExact(dateString, format, CultureInfo.InvariantCulture, DateTimeStyles.None, out parsedDate))
                                        {
                                            formattedDate = parsedDate.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture);
                                            dateParseSuccess = true;
                                            break;
                                        }
                                    }

                                    // Nếu chưa parse được, thử với culture Việt Nam
                                    if (!dateParseSuccess)
                                    {
                                        CultureInfo vietnameseCulture = new CultureInfo("vi-VN");
                                        DateTime parsedDate;

                                        if (DateTime.TryParse(dateString, vietnameseCulture, DateTimeStyles.None, out parsedDate))
                                        {
                                            formattedDate = parsedDate.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture);
                                            dateParseSuccess = true;
                                        }
                                    }

                                    // Nếu vẫn không parse được, giữ nguyên chuỗi gốc
                                    if (!dateParseSuccess)
                                    {
                                        formattedDate = dateString;
                                    }
                                }
                            }
                        }

                        // Xác định tên thay đổi điện thoại
                        string phoneChangeName;
                        if (line == "3")
                        {
                            phoneChangeName = "NT LẦN 1";
                        }else if (line == "5")
                        {
                            phoneChangeName = "Đã Nhắc QH";
                        }
                        else
                        {
                            DataGridViewCell phoneChangeCell = row.Cells[9];
                            phoneChangeName = phoneChangeCell?.Value?.ToString() ?? "";
                        }

                        // Tạo đối tượng thông tin nhắc nợ
                        Form1.DebtReminderManager.DebtReminderInfo reminderInfo = new Form1.DebtReminderManager.DebtReminderInfo();
                        reminderInfo.MaHD = contractCode;
                        reminderInfo.TenKh = customerName;

                        // Lấy thông tin iCloud từ cột 10
                        DataGridViewCell icloudUserCell = row.Cells[10];
                        reminderInfo.userIcloud = icloudUserCell?.Value?.ToString() ?? "";

                        // Lấy mật khẩu iCloud từ cột 11
                        DataGridViewCell icloudPassCell = row.Cells[11];
                        reminderInfo.PassIcloud = icloudPassCell?.Value?.ToString() ?? "";

                        reminderInfo.NamePhoneChange = phoneChangeName;

                        // Lấy ghi chú từ cột 3
                        DataGridViewCell noteCell = row.Cells[3];
                        reminderInfo.Note = noteCell?.Value?.ToString() ?? "";

                        // Lấy Line từ cột 2
                        DataGridViewCell lineCell = row.Cells[2];
                        reminderInfo.Line = lineCell?.Value?.ToString() ?? "";

                        reminderInfo.date = formattedDate;

                        // Lấy tên số từ cột 8
                        DataGridViewCell phoneNameCell = row.Cells[8];
                        reminderInfo.TenSo = phoneNameCell?.Value?.ToString() ?? "";

                        // Kiểm tra xem có ít nhất một trường có dữ liệu không
                        bool hasAnyData = !string.IsNullOrEmpty(reminderInfo.TenKh) ||
                                         !string.IsNullOrEmpty(reminderInfo.userIcloud) ||
                                         !string.IsNullOrEmpty(reminderInfo.PassIcloud) ||
                                         !string.IsNullOrEmpty(reminderInfo.Note) ||
                                         !string.IsNullOrEmpty(reminderInfo.Line) ||
                                         !string.IsNullOrEmpty(reminderInfo.date) ||
                                         !string.IsNullOrEmpty(reminderInfo.NamePhoneChange);

                        if (hasAnyData)
                        {
                            this.debtData[uniqueKey] = reminderInfo;
                        }
                    }

                    // Chuẩn bị dữ liệu gốc để lưu
                    int rowIndex = 0;
                    Form1.originalData.Clear();

                    foreach (KeyValuePair<string, Form1.DebtReminderManager.DebtReminderInfo> debtEntry in this.debtData)
                    {
                        rowIndex++;

                        string formattedRow = string.Format("{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}|{8}", new object[]
                        {
                debtEntry.Key,
                debtEntry.Value.Note,
                debtEntry.Value.userIcloud,
                debtEntry.Value.PassIcloud,
                debtEntry.Value.NamePhoneChange,
                rowIndex - 1,
                debtEntry.Value.Line,
                debtEntry.Value.date,
                debtEntry.Value.TenSo
                        });

                        Form1.originalData.Add(formattedRow);
                    }

                    // Lưu dữ liệu vào MongoDB dựa trên loại line
                    if (line == "1")
                    {
                        result = this.SaveDataMongoDB1();
                    }
                    else if (line == "2")
                    {
                        result = this.SaveDataMongoDB2();
                    }
                    else if (line == "3")
                    {
                        result = this.SaveDataMongoDB(Form1.shouldMerge);
                    }else if(line == "5")
                    {
                        result = SaveDataMongoQuaHan(checkQuaHan);
                    }
                    else
                    {
                        string jsonContent = JsonConvert.SerializeObject(this.debtData, Formatting.Indented);
                        result = this.SaveDataDeleteMongoDB(jsonContent);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Lỗi khi lưu dữ liệu nhắc nợ: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                    result = false;
                }

                return result;
            }

            public bool SaveDataDeleteMongoDB(string jsonContent)
            {
                bool result;
                try
                {
                    string text = "mongodb+srv://banhmichaothuongnhoicloud:8390813asd@cluster0.srtd7rc.mongodb.net/?retryWrites=true&w=majority&appName=Cluster0";
                    MongoClientSettings mongoClientSettings = MongoClientSettings.FromConnectionString(text);
                    mongoClientSettings.ServerApi = new ServerApi(ServerApiVersion.V1, default(Optional<bool?>), default(Optional<bool?>));
                    MongoClient mongoClient = new MongoClient(mongoClientSettings);
                    IMongoDatabase database = mongoClient.GetDatabase("duLieuAPP", null);
                    IMongoCollection<Form1.DebtReminderManager.dataRungIcloud> collection = database.GetCollection<Form1.DebtReminderManager.dataRungIcloud>("dataRung", null);
                    FilterDefinition<Form1.DebtReminderManager.dataRungIcloud> filterDefinition = Builders<Form1.DebtReminderManager.dataRungIcloud>.Filter.Eq<string>((Form1.DebtReminderManager.dataRungIcloud x) => x.Name, "dulieuXoa");
                    Form1.DebtReminderManager.dataRungIcloud dataRungIcloud = IFindFluentExtensions.FirstOrDefault<Form1.DebtReminderManager.dataRungIcloud, Form1.DebtReminderManager.dataRungIcloud>(IMongoCollectionExtensions.Find<Form1.DebtReminderManager.dataRungIcloud>(collection, filterDefinition, null), default(CancellationToken));
                    bool flag = dataRungIcloud != null;
                    if (flag)
                    {
                        UpdateDefinition<Form1.DebtReminderManager.dataRungIcloud> updateDefinition = Builders<Form1.DebtReminderManager.dataRungIcloud>.Update.Set<string>((Form1.DebtReminderManager.dataRungIcloud x) => x.Data, jsonContent);
                        collection.UpdateOne(filterDefinition, updateDefinition, null, default(CancellationToken));
                    }
                    else
                    {
                        Form1.DebtReminderManager.dataRungIcloud dataRungIcloud2 = new Form1.DebtReminderManager.dataRungIcloud
                        {
                            Name = "dulieuXoa",
                            Data = jsonContent
                        };
                        collection.InsertOne(dataRungIcloud2, null, default(CancellationToken));
                    }
                    IMongoCollection<Form1.DebtReminderManager.dataRungIcloud> collection2 = database.GetCollection<Form1.DebtReminderManager.dataRungIcloud>("DataBackup", null);
                    string str = DateTime.Now.ToString("_dd_MM_yyyy_HH_mm_ss");
                    FilterDefinition<Form1.DebtReminderManager.dataRungIcloud> filterDefinition2 = Builders<Form1.DebtReminderManager.dataRungIcloud>.Filter.Eq<string>((Form1.DebtReminderManager.dataRungIcloud x) => x.Name, "dulieuXoa" + str);
                    Form1.DebtReminderManager.dataRungIcloud dataRungIcloud3 = IFindFluentExtensions.FirstOrDefault<Form1.DebtReminderManager.dataRungIcloud, Form1.DebtReminderManager.dataRungIcloud>(IMongoCollectionExtensions.Find<Form1.DebtReminderManager.dataRungIcloud>(collection2, filterDefinition2, null), default(CancellationToken));
                    bool flag2 = dataRungIcloud3 != null;
                    if (flag2)
                    {
                        UpdateDefinition<Form1.DebtReminderManager.dataRungIcloud> updateDefinition2 = Builders<Form1.DebtReminderManager.dataRungIcloud>.Update.Set<string>((Form1.DebtReminderManager.dataRungIcloud x) => x.Data, jsonContent);
                        collection2.UpdateOne(filterDefinition2, updateDefinition2, null, default(CancellationToken));
                    }
                    else
                    {
                        Form1.DebtReminderManager.dataRungIcloud dataRungIcloud4 = new Form1.DebtReminderManager.dataRungIcloud
                        {
                            Name = "dulieuXoa" + str,
                            Data = jsonContent
                        };
                        collection2.InsertOne(dataRungIcloud4, null, default(CancellationToken));
                    }
                    this.CleanupOldBackupData(collection2, "3");
                    result = true;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Lỗi khi lưu dữ liệu: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                    result = false;
                }
                return result;
            }

            //public bool SaveDataMongoDB(bool shouldMerge)
            //{
            //    bool result;
            //    try
            //    {
            //        string text = JsonConvert.SerializeObject(this.debtData, Formatting.Indented);
            //        string text2 = "mongodb+srv://banhmichaothuongnhoicloud:8390813asd@cluster0.srtd7rc.mongodb.net/?retryWrites=true&w=majority&appName=Cluster0";
            //        MongoClientSettings mongoClientSettings = MongoClientSettings.FromConnectionString(text2);
            //        mongoClientSettings.ServerApi = new ServerApi(ServerApiVersion.V1, default(Optional<bool?>), default(Optional<bool?>));
            //        MongoClient mongoClient = new MongoClient(mongoClientSettings);
            //        IMongoDatabase database = mongoClient.GetDatabase("duLieuAPP", null);
            //        IMongoCollection<Form1.DebtReminderManager.dataRungIcloud> collection = database.GetCollection<Form1.DebtReminderManager.dataRungIcloud>("dataRung", null);
            //        FilterDefinition<Form1.DebtReminderManager.dataRungIcloud> filterDefinition = Builders<Form1.DebtReminderManager.dataRungIcloud>.Filter.Eq<string>((Form1.DebtReminderManager.dataRungIcloud x) => x.Name, "dulieuhangngay");
            //        Form1.DebtReminderManager.dataRungIcloud dataRungIcloud = IFindFluentExtensions.FirstOrDefault<Form1.DebtReminderManager.dataRungIcloud, Form1.DebtReminderManager.dataRungIcloud>(IMongoCollectionExtensions.Find<Form1.DebtReminderManager.dataRungIcloud>(collection, filterDefinition, null), default(CancellationToken));
            //        bool flag = dataRungIcloud != null;
            //        if (flag)
            //        {
            //            if (shouldMerge)
            //            {
            //                JObject jobject = JObject.Parse(dataRungIcloud.Data);
            //                JObject jobject2 = JObject.Parse(text);
            //                jobject.Merge(jobject2, new JsonMergeSettings
            //                {
            //                    MergeArrayHandling = MergeArrayHandling.Union  // ← Sửa ở đây
            //                });
            //                string text3 = jobject.ToString(Formatting.None);
            //                UpdateDefinition<Form1.DebtReminderManager.dataRungIcloud> updateDefinition = Builders<Form1.DebtReminderManager.dataRungIcloud>.Update.Set<string>((Form1.DebtReminderManager.dataRungIcloud x) => x.Data, text3);
            //                collection.UpdateOne(filterDefinition, updateDefinition, null, default(CancellationToken));
            //                Console.WriteLine("Đã merge dữ liệu mới vào dữ liệu cũ");
            //            }
            //            else
            //            {
            //                UpdateDefinition<Form1.DebtReminderManager.dataRungIcloud> updateDefinition2 = Builders<Form1.DebtReminderManager.dataRungIcloud>.Update.Set<string>((Form1.DebtReminderManager.dataRungIcloud x) => x.Data, text);
            //                collection.UpdateOne(filterDefinition, updateDefinition2, null, default(CancellationToken));
            //                Console.WriteLine("Đã thay thế hoàn toàn dữ liệu cũ");
            //            }
            //        }
            //        else
            //        {
            //            Form1.DebtReminderManager.dataRungIcloud dataRungIcloud2 = new Form1.DebtReminderManager.dataRungIcloud
            //            {
            //                Name = "dulieuhangngay",
            //                Data = text
            //            };
            //            collection.InsertOne(dataRungIcloud2, null, default(CancellationToken));
            //        }
            //        string str = DateTime.Now.ToString("_dd_MM_yyyy_HH_mm_ss");
            //        IMongoCollection<Form1.DebtReminderManager.dataRungIcloud> collection2 = database.GetCollection<Form1.DebtReminderManager.dataRungIcloud>("DataBackup", null);
            //        FilterDefinition<Form1.DebtReminderManager.dataRungIcloud> filterDefinition2 = Builders<Form1.DebtReminderManager.dataRungIcloud>.Filter.Eq<string>((Form1.DebtReminderManager.dataRungIcloud x) => x.Name, "dulieuhangngay" + str);
            //        Form1.DebtReminderManager.dataRungIcloud dataRungIcloud3 = IFindFluentExtensions.FirstOrDefault<Form1.DebtReminderManager.dataRungIcloud, Form1.DebtReminderManager.dataRungIcloud>(IMongoCollectionExtensions.Find<Form1.DebtReminderManager.dataRungIcloud>(collection, filterDefinition2, null), default(CancellationToken));
            //        bool flag2 = dataRungIcloud3 != null;
            //        if (flag2)
            //        {
            //            if (shouldMerge)
            //            {
            //                JObject jobject3 = JObject.Parse(dataRungIcloud.Data);
            //                JObject jobject4 = JObject.Parse(text);
            //                jobject3.Merge(jobject4, new JsonMergeSettings
            //                {
            //                    MergeArrayHandling = MergeArrayHandling.Union  // ← Sửa ở đây
            //                });
            //                string text4 = jobject3.ToString(0, Array.Empty<JsonConverter>());
            //                UpdateDefinition<Form1.DebtReminderManager.dataRungIcloud> updateDefinition3 = Builders<Form1.DebtReminderManager.dataRungIcloud>.Update.Set<string>((Form1.DebtReminderManager.dataRungIcloud x) => x.Data, text4);
            //                collection2.UpdateOne(filterDefinition2, updateDefinition3, null, default(CancellationToken));
            //                Console.WriteLine("Đã merge dữ liệu mới vào dữ liệu cũ");
            //            }
            //            else
            //            {
            //                UpdateDefinition<Form1.DebtReminderManager.dataRungIcloud> updateDefinition4 = Builders<Form1.DebtReminderManager.dataRungIcloud>.Update.Set<string>((Form1.DebtReminderManager.dataRungIcloud x) => x.Data, text);
            //                collection2.UpdateOne(filterDefinition2, updateDefinition4, null, default(CancellationToken));
            //                Console.WriteLine("Đã thay thế hoàn toàn dữ liệu cũ");
            //            }
            //        }
            //        else
            //        {
            //            Form1.DebtReminderManager.dataRungIcloud dataRungIcloud4 = new Form1.DebtReminderManager.dataRungIcloud
            //            {
            //                Name = "dulieuhangngay" + str,
            //                Data = text
            //            };
            //            collection2.InsertOne(dataRungIcloud4, null, default(CancellationToken));
            //        }
            //        this.CleanupOldBackupData(collection2, "4");
            //        result = true;
            //    }
            //    catch (Exception ex)
            //    {
            //        MessageBox.Show("Lỗi khi lưu dữ liệu: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Hand);
            //        result = false;
            //    }
            //    return result;
            //}

            public bool SaveDataMongoDB(bool shouldMerge)
            {
                try
                {
                    // Serialize dữ liệu
                    string jsonData = JsonConvert.SerializeObject(this.debtData, Formatting.Indented);

                    // Khởi tạo MongoDB client
                    string connectionString = "mongodb+srv://banhmichaothuongnhoicloud:8390813asd@cluster0.srtd7rc.mongodb.net/?retryWrites=true&w=majority&appName=Cluster0";
                    var settings = MongoClientSettings.FromConnectionString(connectionString);
                    settings.ServerApi = new ServerApi(ServerApiVersion.V1);
                    var client = new MongoClient(settings);
                    var database = client.GetDatabase("duLieuAPP");

                    // === Lưu vào collection chính (dataRung) ===
                    var mainCollection = database.GetCollection<Form1.DebtReminderManager.dataRungIcloud>("dataRung");
                    var mainFilter = Builders<Form1.DebtReminderManager.dataRungIcloud>.Filter.Eq(x => x.Name, "dulieuhangngay");
                    var existingMain = mainCollection.Find(mainFilter).FirstOrDefault();

                    if (existingMain != null)
                    {
                        string dataToSave = shouldMerge ? MergeJsonData(existingMain.Data, jsonData) : jsonData;
                        var update = Builders<Form1.DebtReminderManager.dataRungIcloud>.Update.Set(x => x.Data, dataToSave);
                        mainCollection.UpdateOne(mainFilter, update);
                        Console.WriteLine(shouldMerge ? "Đã merge dữ liệu mới vào dữ liệu cũ" : "Đã thay thế hoàn toàn dữ liệu cũ");
                    }
                    else
                    {
                        mainCollection.InsertOne(new Form1.DebtReminderManager.dataRungIcloud
                        {
                            Name = "dulieuhangngay",
                            Data = jsonData
                        });
                    }

                    // === Lưu vào collection backup (DataBackup) ===
                    var backupCollection = database.GetCollection<Form1.DebtReminderManager.dataRungIcloud>("DataBackup");
                    string backupName = "dulieuhangngay" + DateTime.Now.ToString("_dd_MM_yyyy_HH_mm_ss");
                    var backupFilter = Builders<Form1.DebtReminderManager.dataRungIcloud>.Filter.Eq(x => x.Name, backupName);
                    var existingBackup = backupCollection.Find(backupFilter).FirstOrDefault();

                    if (existingBackup != null)
                    {
                        string dataToSave = shouldMerge ? MergeJsonData(existingBackup.Data, jsonData) : jsonData;
                        var update = Builders<Form1.DebtReminderManager.dataRungIcloud>.Update.Set(x => x.Data, dataToSave);
                        backupCollection.UpdateOne(backupFilter, update);
                    }
                    else
                    {
                        backupCollection.InsertOne(new Form1.DebtReminderManager.dataRungIcloud
                        {
                            Name = backupName,
                            Data = jsonData
                        });
                    }

                    // Dọn dẹp backup cũ
                    this.CleanupOldBackupData(backupCollection, "4");

                    return true;
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Lỗi khi lưu dữ liệu: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                    return false;
                }
            }
            public bool SaveDataMongoQuaHan(bool shouldMerge)
            {
                try
                {
                    // Serialize dữ liệu
                    string jsonData = JsonConvert.SerializeObject(this.debtData, Formatting.Indented);

                    // Khởi tạo MongoDB client
                    string connectionString = "mongodb+srv://banhmichaothuongnhoicloud:8390813asd@cluster0.srtd7rc.mongodb.net/?retryWrites=true&w=majority&appName=Cluster0";
                    var settings = MongoClientSettings.FromConnectionString(connectionString);
                    settings.ServerApi = new ServerApi(ServerApiVersion.V1);
                    var client = new MongoClient(settings);
                    var database = client.GetDatabase("duLieuAPP");

                    // === Lưu vào collection chính (dataRung) ===
                    var mainCollection = database.GetCollection<Form1.DebtReminderManager.dataRungIcloud>("dataRung");
                    var mainFilter = Builders<Form1.DebtReminderManager.dataRungIcloud>.Filter.Eq(x => x.Name, "dulieuquahan");
                    var existingMain = mainCollection.Find(mainFilter).FirstOrDefault();

                    if (existingMain != null)
                    {
                        string dataToSave = shouldMerge ? MergeJsonData(existingMain.Data, jsonData) : jsonData;
                        var update = Builders<Form1.DebtReminderManager.dataRungIcloud>.Update.Set(x => x.Data, dataToSave);
                        mainCollection.UpdateOne(mainFilter, update);
                        Console.WriteLine(shouldMerge ? "Đã merge dữ liệu mới vào dữ liệu cũ" : "Đã thay thế hoàn toàn dữ liệu cũ");
                    }
                    else
                    {
                        mainCollection.InsertOne(new Form1.DebtReminderManager.dataRungIcloud
                        {
                            Name = "dulieuquahan",
                            Data = jsonData
                        });
                    }

                    // === Lưu vào collection backup (DataBackup) ===
                    var backupCollection = database.GetCollection<Form1.DebtReminderManager.dataRungIcloud>("DataBackup");
                    string backupName = "dulieuquahan" + DateTime.Now.ToString("_dd_MM_yyyy_HH_mm_ss");
                    var backupFilter = Builders<Form1.DebtReminderManager.dataRungIcloud>.Filter.Eq(x => x.Name, backupName);
                    var existingBackup = backupCollection.Find(backupFilter).FirstOrDefault();

                    if (existingBackup != null)
                    {
                        string dataToSave = shouldMerge ? MergeJsonData(existingBackup.Data, jsonData) : jsonData;
                        var update = Builders<Form1.DebtReminderManager.dataRungIcloud>.Update.Set(x => x.Data, dataToSave);
                        backupCollection.UpdateOne(backupFilter, update);
                    }
                    else
                    {
                        backupCollection.InsertOne(new Form1.DebtReminderManager.dataRungIcloud
                        {
                            Name = backupName,
                            Data = jsonData
                        });
                    }

                    // Dọn dẹp backup cũ
                    this.CleanupOldBackupData(backupCollection, "4");

                    return true;
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Lỗi khi lưu dữ liệu: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                    return false;
                }
            }
            // Helper method để merge JSON - Giữ nguyên dữ liệu cũ, chỉ thêm key mới
            private string MergeJsonData(string existingJson, string newJson)
            {
                var existingDict = JsonConvert.DeserializeObject<Dictionary<string, DebtReminderInfo>>(existingJson);
                var newDict = JsonConvert.DeserializeObject<Dictionary<string, DebtReminderInfo>>(newJson);

                // Chỉ thêm những key CHƯA có trong dữ liệu cũ
                foreach (var kvp in newDict)
                {
                    if (!existingDict.ContainsKey(kvp.Key))
                    {
                        existingDict[kvp.Key] = kvp.Value;
                    }
                    // Nếu key đã tồn tại thì BỎ QUA (giữ nguyên dữ liệu cũ)
                }

                return JsonConvert.SerializeObject(existingDict, Formatting.None);
            }
            public bool SaveDataMongoDB2()
            {
                bool result;
                try
                {
                    string text = JsonConvert.SerializeObject(this.debtData, Formatting.Indented);
                    string text2 = "mongodb+srv://banhmichaothuongnhoicloud:8390813asd@cluster0.srtd7rc.mongodb.net/?retryWrites=true&w=majority&appName=Cluster0";
                    MongoClientSettings mongoClientSettings = MongoClientSettings.FromConnectionString(text2);
                    mongoClientSettings.ServerApi = new ServerApi(ServerApiVersion.V1, default(Optional<bool?>), default(Optional<bool?>));
                    MongoClient mongoClient = new MongoClient(mongoClientSettings);
                    IMongoDatabase database = mongoClient.GetDatabase("duLieuAPP", null);
                    IMongoCollection<Form1.DebtReminderManager.dataRungIcloud> collection = database.GetCollection<Form1.DebtReminderManager.dataRungIcloud>("dataRung", null);
                    FilterDefinition<Form1.DebtReminderManager.dataRungIcloud> filterDefinition = Builders<Form1.DebtReminderManager.dataRungIcloud>.Filter.Eq<string>((Form1.DebtReminderManager.dataRungIcloud x) => x.Name, "dulieurungLine2");
                    Form1.DebtReminderManager.dataRungIcloud dataRungIcloud = IFindFluentExtensions.FirstOrDefault<Form1.DebtReminderManager.dataRungIcloud, Form1.DebtReminderManager.dataRungIcloud>(IMongoCollectionExtensions.Find<Form1.DebtReminderManager.dataRungIcloud>(collection, filterDefinition, null), default(CancellationToken));
                    bool flag = dataRungIcloud != null;
                    if (flag)
                    {
                        UpdateDefinition<Form1.DebtReminderManager.dataRungIcloud> updateDefinition = Builders<Form1.DebtReminderManager.dataRungIcloud>.Update.Set<string>((Form1.DebtReminderManager.dataRungIcloud x) => x.Data, text);
                        collection.UpdateOne(filterDefinition, updateDefinition, null, default(CancellationToken));
                    }
                    else
                    {
                        Form1.DebtReminderManager.dataRungIcloud dataRungIcloud2 = new Form1.DebtReminderManager.dataRungIcloud
                        {
                            Name = "dulieurungLine2",
                            Data = text
                        };
                        collection.InsertOne(dataRungIcloud2, null, default(CancellationToken));
                    }
                    IMongoCollection<Form1.DebtReminderManager.dataRungIcloud> collection2 = database.GetCollection<Form1.DebtReminderManager.dataRungIcloud>("DataBackup", null);
                    string str = DateTime.Now.ToString("_dd_MM_yyyy_HH_mm_ss");
                    FilterDefinition<Form1.DebtReminderManager.dataRungIcloud> filterDefinition2 = Builders<Form1.DebtReminderManager.dataRungIcloud>.Filter.Eq<string>((Form1.DebtReminderManager.dataRungIcloud x) => x.Name, "Line2" + str);
                    Form1.DebtReminderManager.dataRungIcloud dataRungIcloud3 = IFindFluentExtensions.FirstOrDefault<Form1.DebtReminderManager.dataRungIcloud, Form1.DebtReminderManager.dataRungIcloud>(IMongoCollectionExtensions.Find<Form1.DebtReminderManager.dataRungIcloud>(collection2, filterDefinition2, null), default(CancellationToken));
                    bool flag2 = dataRungIcloud3 != null;
                    if (flag2)
                    {
                        UpdateDefinition<Form1.DebtReminderManager.dataRungIcloud> updateDefinition2 = Builders<Form1.DebtReminderManager.dataRungIcloud>.Update.Set<string>((Form1.DebtReminderManager.dataRungIcloud x) => x.Data, text);
                        collection2.UpdateOne(filterDefinition2, updateDefinition2, null, default(CancellationToken));
                    }
                    else
                    {
                        Form1.DebtReminderManager.dataRungIcloud dataRungIcloud4 = new Form1.DebtReminderManager.dataRungIcloud
                        {
                            Name = "Line2" + str,
                            Data = text
                        };
                        collection2.InsertOne(dataRungIcloud4, null, default(CancellationToken));
                    }
                    this.CleanupOldBackupData(collection2, "2");
                    result = true;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Lỗi khi lưu dữ liệu: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                    result = false;
                }
                return result;
            }
            private void CleanupOldBackupData(IMongoCollection<Form1.DebtReminderManager.dataRungIcloud> collection, string line)
            {
                try
                {
                    string text = "";
                    bool flag = line == "1";
                    if (flag)
                    {
                        text = "Line1_";
                    }
                    else
                    {
                        bool flag2 = line == "2";
                        if (flag2)
                        {
                            text = "Line2_";
                        }
                        else
                        {
                            bool flag3 = line == "3";
                            if (flag3)
                            {
                                text = "dulieuXoa_";
                            }
                            else
                            {
                                bool flag4 = line == "4";
                                if (flag4)
                                {
                                    text = "dulieuhangngay_";
                                }
                            }
                        }
                    }
                    DateTime t = DateTime.Now.AddDays(-2.0);
                    List<Form1.DebtReminderManager.dataRungIcloud> list = IAsyncCursorSourceExtensions.ToList<Form1.DebtReminderManager.dataRungIcloud>(IMongoCollectionExtensions.Find<Form1.DebtReminderManager.dataRungIcloud>(collection, Builders<Form1.DebtReminderManager.dataRungIcloud>.Filter.Regex((Form1.DebtReminderManager.dataRungIcloud x) => x.Name, "^" + text), null), default(CancellationToken));
                    foreach (Form1.DebtReminderManager.dataRungIcloud dataRungIcloud in list)
                    {
                        string text2 = dataRungIcloud.Name.Replace(text, "");
                        string[] array = text2.Split(new char[]
                        {
                            '_'
                        });
                        bool flag5 = array.Length >= 6;
                        if (flag5)
                        {
                            try
                            {
                                int day = int.Parse(array[0]);
                                int month = int.Parse(array[1]);
                                int year = int.Parse(array[2]);
                                int hour = int.Parse(array[3]);
                                int minute = int.Parse(array[4]);
                                int second = int.Parse(array[5]);
                                DateTime dateTime = new DateTime(year, month, day, hour, minute, second);
                                bool flag6 = dateTime < t;
                                if (flag6)
                                {
                                    FilterDefinition<Form1.DebtReminderManager.dataRungIcloud> filterDefinition = Builders<Form1.DebtReminderManager.dataRungIcloud>.Filter.Eq<string>((Form1.DebtReminderManager.dataRungIcloud x) => x.Name, dataRungIcloud.Name);
                                    collection.DeleteOne(filterDefinition, default(CancellationToken));
                                    Console.WriteLine(string.Format("Đã xóa backup cũ: {0} - Ngày: {1:dd/MM/yyyy HH:mm:ss}", dataRungIcloud.Name, dateTime));
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("Không thể parse ngày từ tên backup: " + dataRungIcloud.Name + " - Lỗi: " + ex.Message);
                            }
                        }
                    }
                }
                catch (Exception ex2)
                {
                    Console.WriteLine("Lỗi khi dọn dẹp backup cũ: " + ex2.Message);
                }
            }
            public bool SaveDataMongoDB1()
            {
                bool result;
                try
                {
                    string text = JsonConvert.SerializeObject(this.debtData, Formatting.Indented);
                    string text2 = "mongodb+srv://banhmichaothuongnhoicloud:8390813asd@cluster0.srtd7rc.mongodb.net/?retryWrites=true&w=majority&appName=Cluster0";
                    MongoClientSettings mongoClientSettings = MongoClientSettings.FromConnectionString(text2);
                    mongoClientSettings.ServerApi = new ServerApi(ServerApiVersion.V1, default(Optional<bool?>), default(Optional<bool?>));
                    MongoClient mongoClient = new MongoClient(mongoClientSettings);
                    IMongoDatabase database = mongoClient.GetDatabase("duLieuAPP", null);
                    IMongoCollection<Form1.DebtReminderManager.dataRungIcloud> collection = database.GetCollection<Form1.DebtReminderManager.dataRungIcloud>("dataRung", null);
                    FilterDefinition<Form1.DebtReminderManager.dataRungIcloud> filterDefinition = Builders<Form1.DebtReminderManager.dataRungIcloud>.Filter.Eq<string>((Form1.DebtReminderManager.dataRungIcloud x) => x.Name, "Line1");
                    Form1.DebtReminderManager.dataRungIcloud dataRungIcloud = IFindFluentExtensions.FirstOrDefault<Form1.DebtReminderManager.dataRungIcloud, Form1.DebtReminderManager.dataRungIcloud>(IMongoCollectionExtensions.Find<Form1.DebtReminderManager.dataRungIcloud>(collection, filterDefinition, null), default(CancellationToken));
                    bool flag = dataRungIcloud != null;
                    if (flag)
                    {
                        UpdateDefinition<Form1.DebtReminderManager.dataRungIcloud> updateDefinition = Builders<Form1.DebtReminderManager.dataRungIcloud>.Update.Set<string>((Form1.DebtReminderManager.dataRungIcloud x) => x.Data, text);
                        collection.UpdateOne(filterDefinition, updateDefinition, null, default(CancellationToken));
                    }
                    else
                    {
                        Form1.DebtReminderManager.dataRungIcloud dataRungIcloud2 = new Form1.DebtReminderManager.dataRungIcloud
                        {
                            Name = "Line1",
                            Data = text
                        };
                        collection.InsertOne(dataRungIcloud2, null, default(CancellationToken));
                    }
                    IMongoCollection<Form1.DebtReminderManager.dataRungIcloud> collection2 = database.GetCollection<Form1.DebtReminderManager.dataRungIcloud>("DataBackup", null);
                    string str = DateTime.Now.ToString("_dd_MM_yyyy_HH_mm_ss");
                    FilterDefinition<Form1.DebtReminderManager.dataRungIcloud> filterDefinition2 = Builders<Form1.DebtReminderManager.dataRungIcloud>.Filter.Eq<string>((Form1.DebtReminderManager.dataRungIcloud x) => x.Name, "Line1" + str);
                    Form1.DebtReminderManager.dataRungIcloud dataRungIcloud3 = IFindFluentExtensions.FirstOrDefault<Form1.DebtReminderManager.dataRungIcloud, Form1.DebtReminderManager.dataRungIcloud>(IMongoCollectionExtensions.Find<Form1.DebtReminderManager.dataRungIcloud>(collection2, filterDefinition2, null), default(CancellationToken));
                    bool flag2 = dataRungIcloud3 != null;
                    if (flag2)
                    {
                        UpdateDefinition<Form1.DebtReminderManager.dataRungIcloud> updateDefinition2 = Builders<Form1.DebtReminderManager.dataRungIcloud>.Update.Set<string>((Form1.DebtReminderManager.dataRungIcloud x) => x.Data, text);
                        collection2.UpdateOne(filterDefinition2, updateDefinition2, null, default(CancellationToken));
                    }
                    else
                    {
                        Form1.DebtReminderManager.dataRungIcloud dataRungIcloud4 = new Form1.DebtReminderManager.dataRungIcloud
                        {
                            Name = "Line1" + str,
                            Data = text
                        };
                        collection2.InsertOne(dataRungIcloud4, null, default(CancellationToken));
                    }
                    this.CleanupOldBackupData(collection2, "1");
                    result = true;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Lỗi khi lưu dữ liệu: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                    result = false;
                }
                return result;
            }

            internal bool SaveDataBackMongoDB1(string jsonContent)
            {
                bool result;
                try
                {
                    string text = "mongodb+srv://banhmichaothuongnhoicloud:8390813asd@cluster0.srtd7rc.mongodb.net/?retryWrites=true&w=majority&appName=Cluster0";
                    MongoClientSettings mongoClientSettings = MongoClientSettings.FromConnectionString(text);
                    mongoClientSettings.ServerApi = new ServerApi(ServerApiVersion.V1, default(Optional<bool?>), default(Optional<bool?>));
                    MongoClient mongoClient = new MongoClient(mongoClientSettings);
                    IMongoDatabase database = mongoClient.GetDatabase("duLieuAPP", null);
                    IMongoCollection<Form1.DebtReminderManager.dataRungIcloud> collection = database.GetCollection<Form1.DebtReminderManager.dataRungIcloud>("dataRung", null);
                    FilterDefinition<Form1.DebtReminderManager.dataRungIcloud> filterDefinition = Builders<Form1.DebtReminderManager.dataRungIcloud>.Filter.Eq<string>((Form1.DebtReminderManager.dataRungIcloud x) => x.Name, "Line1");
                    Form1.DebtReminderManager.dataRungIcloud dataRungIcloud = IFindFluentExtensions.FirstOrDefault<Form1.DebtReminderManager.dataRungIcloud, Form1.DebtReminderManager.dataRungIcloud>(IMongoCollectionExtensions.Find<Form1.DebtReminderManager.dataRungIcloud>(collection, filterDefinition, null), default(CancellationToken));
                    bool flag = dataRungIcloud != null;
                    if (flag)
                    {
                        UpdateDefinition<Form1.DebtReminderManager.dataRungIcloud> updateDefinition = Builders<Form1.DebtReminderManager.dataRungIcloud>.Update.Set<string>((Form1.DebtReminderManager.dataRungIcloud x) => x.Data, jsonContent);
                        collection.UpdateOne(filterDefinition, updateDefinition, null, default(CancellationToken));
                    }
                    else
                    {
                        Form1.DebtReminderManager.dataRungIcloud dataRungIcloud2 = new Form1.DebtReminderManager.dataRungIcloud
                        {
                            Name = "Line1",
                            Data = jsonContent
                        };
                        collection.InsertOne(dataRungIcloud2, null, default(CancellationToken));
                    }
                    IMongoCollection<Form1.DebtReminderManager.dataRungIcloud> collection2 = database.GetCollection<Form1.DebtReminderManager.dataRungIcloud>("DataBackup", null);
                    string str = DateTime.Now.ToString("_dd_MM_yyyy_HH_mm_ss");
                    FilterDefinition<Form1.DebtReminderManager.dataRungIcloud> filterDefinition2 = Builders<Form1.DebtReminderManager.dataRungIcloud>.Filter.Eq<string>((Form1.DebtReminderManager.dataRungIcloud x) => x.Name, "Line1" + str);
                    Form1.DebtReminderManager.dataRungIcloud dataRungIcloud3 = IFindFluentExtensions.FirstOrDefault<Form1.DebtReminderManager.dataRungIcloud, Form1.DebtReminderManager.dataRungIcloud>(IMongoCollectionExtensions.Find<Form1.DebtReminderManager.dataRungIcloud>(collection2, filterDefinition2, null), default(CancellationToken));
                    bool flag2 = dataRungIcloud3 != null;
                    if (flag2)
                    {
                        UpdateDefinition<Form1.DebtReminderManager.dataRungIcloud> updateDefinition2 = Builders<Form1.DebtReminderManager.dataRungIcloud>.Update.Set<string>((Form1.DebtReminderManager.dataRungIcloud x) => x.Data, jsonContent);
                        collection2.UpdateOne(filterDefinition2, updateDefinition2, null, default(CancellationToken));
                    }
                    else
                    {
                        Form1.DebtReminderManager.dataRungIcloud dataRungIcloud4 = new Form1.DebtReminderManager.dataRungIcloud
                        {
                            Name = "Line1" + str,
                            Data = jsonContent
                        };
                        collection2.InsertOne(dataRungIcloud4, null, default(CancellationToken));
                    }
                    result = true;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Lỗi khi lưu dữ liệu: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                    result = false;
                }
                return result;
            }

            public bool SaveDataBackMongoDB2(string jsonContent2)
            {
                bool result;
                try
                {
                    string text = "mongodb+srv://banhmichaothuongnhoicloud:8390813asd@cluster0.srtd7rc.mongodb.net/?retryWrites=true&w=majority&appName=Cluster0";
                    MongoClientSettings mongoClientSettings = MongoClientSettings.FromConnectionString(text);
                    mongoClientSettings.ServerApi = new ServerApi(ServerApiVersion.V1, default(Optional<bool?>), default(Optional<bool?>));
                    MongoClient mongoClient = new MongoClient(mongoClientSettings);
                    IMongoDatabase database = mongoClient.GetDatabase("duLieuAPP", null);
                    IMongoCollection<Form1.DebtReminderManager.dataRungIcloud> collection = database.GetCollection<Form1.DebtReminderManager.dataRungIcloud>("dataRung", null);
                    FilterDefinition<Form1.DebtReminderManager.dataRungIcloud> filterDefinition = Builders<Form1.DebtReminderManager.dataRungIcloud>.Filter.Eq<string>((Form1.DebtReminderManager.dataRungIcloud x) => x.Name, "dulieurungLine2");
                    Form1.DebtReminderManager.dataRungIcloud dataRungIcloud = IFindFluentExtensions.FirstOrDefault<Form1.DebtReminderManager.dataRungIcloud, Form1.DebtReminderManager.dataRungIcloud>(IMongoCollectionExtensions.Find<Form1.DebtReminderManager.dataRungIcloud>(collection, filterDefinition, null), default(CancellationToken));
                    bool flag = dataRungIcloud != null;
                    if (flag)
                    {
                        UpdateDefinition<Form1.DebtReminderManager.dataRungIcloud> updateDefinition = Builders<Form1.DebtReminderManager.dataRungIcloud>.Update.Set<string>((Form1.DebtReminderManager.dataRungIcloud x) => x.Data, jsonContent2);
                        collection.UpdateOne(filterDefinition, updateDefinition, null, default(CancellationToken));
                    }
                    else
                    {
                        Form1.DebtReminderManager.dataRungIcloud dataRungIcloud2 = new Form1.DebtReminderManager.dataRungIcloud
                        {
                            Name = "dulieurungLine2",
                            Data = jsonContent2
                        };
                        collection.InsertOne(dataRungIcloud2, null, default(CancellationToken));
                    }
                    IMongoCollection<Form1.DebtReminderManager.dataRungIcloud> collection2 = database.GetCollection<Form1.DebtReminderManager.dataRungIcloud>("DataBackup", null);
                    string str = DateTime.Now.ToString("_dd_MM_yyyy_HH_mm_ss");
                    FilterDefinition<Form1.DebtReminderManager.dataRungIcloud> filterDefinition2 = Builders<Form1.DebtReminderManager.dataRungIcloud>.Filter.Eq<string>((Form1.DebtReminderManager.dataRungIcloud x) => x.Name, "Line2" + str);
                    Form1.DebtReminderManager.dataRungIcloud dataRungIcloud3 = IFindFluentExtensions.FirstOrDefault<Form1.DebtReminderManager.dataRungIcloud, Form1.DebtReminderManager.dataRungIcloud>(IMongoCollectionExtensions.Find<Form1.DebtReminderManager.dataRungIcloud>(collection2, filterDefinition2, null), default(CancellationToken));
                    bool flag2 = dataRungIcloud3 != null;
                    if (flag2)
                    {
                        UpdateDefinition<Form1.DebtReminderManager.dataRungIcloud> updateDefinition2 = Builders<Form1.DebtReminderManager.dataRungIcloud>.Update.Set<string>((Form1.DebtReminderManager.dataRungIcloud x) => x.Data, jsonContent2);
                        collection2.UpdateOne(filterDefinition2, updateDefinition2, null, default(CancellationToken));
                    }
                    else
                    {
                        Form1.DebtReminderManager.dataRungIcloud dataRungIcloud4 = new Form1.DebtReminderManager.dataRungIcloud
                        {
                            Name = "Line2" + str,
                            Data = jsonContent2
                        };
                        collection2.InsertOne(dataRungIcloud4, null, default(CancellationToken));
                    }
                    result = true;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Lỗi khi lưu dữ liệu: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                    result = false;
                }
                return result;
            }
        }
        // Biến static để quản lý trạng thái
        private static bool isProcessing = false;
        private static List<IWebDriver> activeDrivers = new List<IWebDriver>();
        // Thêm biến static để quản lý vị trí toàn cục
        public class ChromePositionManager
        {
            private static int currentPosition = 0;
            private static readonly object positionLock = new object();

            public static int GetNextPosition()
            {
                lock (positionLock)
                {
                    currentPosition++;
                    return currentPosition;
                }
            }

            public static void ResetPosition()
            {
                lock (positionLock)
                {
                    currentPosition = 0;
                }
            }
        }


        //private void HandleLineButtonClick(int rowIndex)
        //{
        //    int displayRowIndex = FindDisplayRowIndexByRealIndex(rowIndex);
        //    if (displayRowIndex == -1)
        //    {

        //        return;
        //    }
        //    // Lấy thông tin từ dòng được click
        //    string searchText = dataGridView1.Rows[displayRowIndex].Cells[1].Value?.ToString();
        //    // Mở LINE và search luôn
        //    if (LineWindowManager.OpenLineAndSearch(searchText))
        //    {
        //        // Thành công - có thể minimize form hiện tại
        //        this.WindowState = FormWindowState.Minimized;
        //    }
        //    else
        //    {
        //        MessageBox.Show("Lỗi khi mở LINE!");
        //    }
        //    if (!string.IsNullOrEmpty(searchText))
        //    {
        //        Clipboard.SetText(searchText);
        //        richTextBox1.AppendText("Đã Mở Line và Copy Tên Khách Hàng \r\n");
        //    }

        //}
        private void RadioButton_CheckedChanged(object sender, EventArgs e)
        {
            richTextBox2.Text = GetMessageText();
        }
        
        // Thêm phương thức này vào class của bạn
        private string GetMessageText()
        {
            if (radioButton1.Checked)
            {
                return "เงินกู้ของคุณเกินกำหนดชำระแล้ว กรุณาชำระเพื่อหลีกเลี่ยงปัญหา คำเตือนครั้งที่ 2";
            }
            else if (radioButton2.Checked)
            {
                return "เงินกู้ของคุณเลยกำหนดชำระแล้ว กรุณาชำระเงินเพื่อหลีกเลี่ยงปัญหาที่อาจเกิดขึ้น คำเตือนครั้งที่ 3 หลังจากคำเตือนนี้ หากเรายังไม่ได้รับการชำระเงิน เราจะส่งข้อมูลของคุณให้กับบริษัททวงหนี้";
            }
            else
            {
                return richTextBox2.Text;
            }
        }
        
        private void HandleLineButtonClick(int rowIndex)
        {
            int num = this.FindDisplayRowIndexByRealIndex(rowIndex);
            bool flag = num == -1;
            if (!flag)
            {
                object value = this.dataGridView1.Rows[num].Cells[1].Value;
                string searchText = (value != null) ? value.ToString() : null;

                // Lấy text từ phương thức chung
                string text = GetMessageText();
                richTextBox2.Text = text; // Cập nhật hiển thị

                string text2 = (text != null) ? text.ToString() : null;
                bool flag2 = text2 == "" || text2 == null;
                if (flag2)
                {
                    object value2 = this.dataGridView1.Rows[num].Cells[1].Value;
                    text2 = ((value2 != null) ? value2.ToString() : null);
                }

                bool flag3 = LineWindowManager.OpenLineAndSearch(searchText);
                if (flag3)
                {
                    base.WindowState = FormWindowState.Minimized;
                }
                else
                {
                    MessageBox.Show("Lỗi khi mở LINE!");
                }

                bool flag4 = !string.IsNullOrEmpty(text2);
                if (flag4)
                {
                    bool flag5 = Form1.shouldMerge;
                    if (flag5)
                    {
                        this.HandleLoiNhacButtonClick(rowIndex);
                    }
                    else if (checkQuaHan)
                    {
                        HandleLoiNhacButtonClick1(rowIndex);
                    }
                    else
                    {
                        Clipboard.SetText(text2);
                    }
                    this.richTextBox1.AppendText("Đã Mở Line và Copy Lời Nhắc \r\n");
                }
            }
        }
        //private void HandleXoaButtonClick(int rowIndex)
        //{
        //    int displayRowIndex = FindDisplayRowIndexByRealIndex(rowIndex);
        //    if (displayRowIndex == -1)
        //    {

        //        return;
        //    }
        //    try
        //    {
        //        if (displayRowIndex >= 0 && displayRowIndex < dataGridView1.Rows.Count)
        //        {
        //            // Kiểm tra xem có phải dòng mới không (new row)
        //            if (!dataGridView1.Rows[displayRowIndex].IsNewRow)
        //            {
        //                dataGridView1.Rows.RemoveAt(displayRowIndex);
        //                MessageBox.Show($"Đã xóa dòng {displayRowIndex}", "Thông báo",
        //                    MessageBoxButtons.OK, MessageBoxIcon.Information);
        //            }
        //            else
        //            {
        //                MessageBox.Show("Không thể xóa dòng mới", "Thông báo",
        //                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
        //            }
        //        }
        //        else
        //        {
        //            MessageBox.Show("Index không hợp lệ", "Lỗi",
        //                MessageBoxButtons.OK, MessageBoxIcon.Error);
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show($"Lỗi khi xóa dòng: {ex.Message}", "Lỗi",
        //            MessageBoxButtons.OK, MessageBoxIcon.Error);
        //    }

        //}
        //private void HandleLoiNhacButtonClick(int rowIndex)
        //{

        //    try
        //    {
        //        int displayRowIndex = FindDisplayRowIndexByRealIndex(rowIndex);
        //        if (displayRowIndex == -1)
        //        {

        //            return;
        //        }
        //        // Lấy thông tin từ dòng được click
        //        string phoneNumber = dataGridView1.Rows[displayRowIndex].Cells[2].Value?.ToString();
        //        string nameGuest = dataGridView1.Rows[displayRowIndex].Cells[1].Value?.ToString();
        //        string loinhac = "สวัสดีค่ะ เงินกู้ของคุณจะครบกำหนดในวันพรุ่งนี้ ดิฉันขอแจ้งเตือนให้คุณชำระหนี้ตามสัญญาเลขที่: "+nameGuest+" จำนวนเงินที่คุณต้องชำระคือ "+phoneNumber+" บาท - กรุณาชำระเงินตามข้อมูลด้านล่างนี้: ชำระเข้าบัญชีเลขที่: 0334421279 ชื่อบัญชี: นิตยา ธนาคาร: ไทยพาณิชย์ -- SCB - หรือส่งไปยัง QR Code ที่ดิฉันได้ส่งให้คุณด้านล่าง - หากชำระเงินแล้ว กรุณาส่งภาพสลิปการโอนเงิน => หากคุณไม่ชำระเงินตรงเวลา บริษัทของเราจะแจ้งและล็อคโทรศัพท์ของคุณ ขอบคุณค่ะ";
        //        if (!string.IsNullOrEmpty(loinhac))
        //        {
        //            // Cách 1: Copy cả text và hình ảnh cùng lúc bằng DataObject
        //            DataObject dataObject = new DataObject();

        //            // Thêm text vào clipboard
        //            dataObject.SetData(DataFormats.Text, loinhac);
        //            dataObject.SetData(DataFormats.UnicodeText, loinhac);

        //            // Thêm hình ảnh vào clipboard
        //            try
        //            {
        //                // Nếu bạn có đường dẫn file hình ảnh
        //                Image image = Image.FromFile("S3350530.jpg");
        //                dataObject.SetData(DataFormats.Bitmap, image);

        //                // Hoặc nếu bạn có hình ảnh từ PictureBox
        //                // Image image = pictureBox1.Image;
        //                // dataObject.SetData(DataFormats.Bitmap, image);

        //                // Set DataObject vào clipboard
        //                Clipboard.SetDataObject(dataObject, true);

        //                richTextBox1.AppendText("Đã Copy Lời Nhắc và Hình Ảnh\r\n");
        //            }
        //            catch (Exception ex)
        //            {
        //                MessageBox.Show("Lỗi khi load hình ảnh: " + ex.Message);
        //            }
        //        }

        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show($"Lỗi khi copy: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //    }
        //}
        private void HandleLoiNhacButtonClick(int rowIndex)
        {
            try
            {
                int num = this.FindDisplayRowIndexByRealIndex(rowIndex);
                bool flag = num == -1;
                if (!flag)
                {
                    object value = this.dataGridView1.Rows[num].Cells[3].Value;
                    string input = (value != null) ? value.ToString() : null;
                    object value2 = this.dataGridView1.Rows[num].Cells[1].Value;
                    string text = (value2 != null) ? value2.ToString() : null;
                    string pattern = "[\\d,]+";
                    MatchCollection matchCollection = Regex.Matches(input, pattern);
                    string text2 = "";
                    foreach (object obj in matchCollection)
                    {
                        Match match = (Match)obj;
                        text2 = match.Value;
                    }
                    string text3 = string.Concat(new string[]
                    {
                "สวัสดีค่ะ เงินกู้ของคุณจะครบกำหนดในวันพรุ่งนี้ ดิฉันขอแจ้งเตือนให้คุณชำระหนี้ตามสัญญาเลขที่: ",
                text,
                " จำนวนเงินที่คุณต้องชำระคือ ",
                text2,
                " บาท - กรุณาชำระเงินตามข้อมูลด้านล่างนี้: ชำระเข้าบัญชีเลขที่: 0334421279 ชื่อบัญชี: นิตยา ธนาคาร: ไทยพาณิชย์ -- SCB - หรือส่งไปยัง QR Code ที่ดิฉันได้ส่งให้คุณด้านล่าง - หากชำระเงินแล้ว กรุณาส่งภาพสลิปการโอนเงิน => หากคุณไม่ชำระเงินตรงเวลา บริษัทของเราจะแจ้งและล็อคโทรศัพท์ของคุณ ขอบคุณค่ะ"
                    });
                    bool flag2 = !string.IsNullOrEmpty(text3);
                    if (flag2)
                    {
                        DataObject dataObject = new DataObject();
                        dataObject.SetData(DataFormats.Text, text3);
                        dataObject.SetData(DataFormats.UnicodeText, text3);
                        try
                        {
                            Image data = Image.FromFile("S3350530.jpg");
                            dataObject.SetData(DataFormats.Bitmap, data);
                            Clipboard.SetDataObject(dataObject, true);
                            this.richTextBox1.AppendText("Đã Copy Lời Nhắc và Hình Ảnh\r\n");
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Lỗi khi load hình ảnh: " + ex.Message);
                        }
                    }
                }
            }
            catch (Exception ex2)
            {
                MessageBox.Show("Lỗi khi copy: " + ex2.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Hand);
            }
        }
        private void HandleLoiNhacButtonClick1(int rowIndex)
        {
            try
            {
                int displayRowIndex = this.FindDisplayRowIndexByRealIndex(rowIndex);
                bool rowNotFound = displayRowIndex == -1;

                if (!rowNotFound)
                {
                    // Tạo tin nhắn nhắc nợ bằng tiếng Thái
                    string reminderMessage = string.Concat(new string[]
                    {
                "เราเพิ่งได้รับการแจ้งเตือน! โทรศัพท์ของคุณจะถูกล็อคภายใน 24 ชั่วโมงข้างหน้า สาเหตุคือคุณติดหนี้และไม่ยอมชำระเงิน ทีมงานของเราได้ติดต่อคุณหลายครั้งแล้วแต่คุณไม่ตั้งใจจะชำระเงิน เรากำลังแจ้งให้คุณทราบก่อนที่โทรศัพท์ของคุณจะถูกล็อค หากเราไม่ได้รับการตอบกลับจากคุณ โทรศัพท์ของคุณจะถูกล็อคภายใน 24 ชั่วโมงข้างหน้าตามค่าเริ่มต้น เราจะทำการล็อคโดยไม่แจ้งให้ทราบล่วงหน้า ดังนั้นคุณกรุณาบันทึกข้อมูลให้ครบถ้วนด้วยนะคะ\n\n",
                "หลังจากล็อคโทรศัพท์ของคุณแล้ว โปรดติดต่อที่อยู่ด้านล่าง:\n",
                "Line: MiuMiu-NamThanIcloud, ID line: annammedia112000\n",
                "Line: Annam-Nathamicloud, ID line: annamicloud\n",
                "FB: https://www.facebook.com/100083517735368"
                    });

                    bool hasValidMessage = !string.IsNullOrEmpty(reminderMessage);

                    if (hasValidMessage)
                    {
                        try
                        {
                            Clipboard.SetText(reminderMessage);
                            this.richTextBox1.AppendText("Đã Copy Lời Nhắc\r\n");
                        }
                        catch (Exception clipboardException)
                        {
                            MessageBox.Show("Lỗi khi copy: " + clipboardException.Message);
                        }
                    }
                }
            }
            catch (Exception copyException)
            {
                MessageBox.Show("Lỗi khi copy: " + copyException.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Hand);
            }
        }
        private void HandleXoaButtonClick(int rowIndex)
        {
            int num = this.FindDisplayRowIndexByRealIndex(rowIndex);
            bool flag = num == -1;
            if (!flag)
            {
                try
                {
                    bool flag2 = num >= 0 && num < this.dataGridView1.Rows.Count;
                    if (flag2)
                    {
                        DataGridViewCell dataGridViewCell = this.dataGridView1.Rows[num].Cells[2];
                        string text;
                        if (dataGridViewCell == null)
                        {
                            text = null;
                        }
                        else
                        {
                            object value = dataGridViewCell.Value;
                            text = ((value != null) ? value.ToString() : null);
                        }
                        string a = text ?? "";
                        bool flag3 = !this.dataGridView1.Rows[num].IsNewRow;
                        if (flag3)
                        {
                            this.dataGridView1.Rows.RemoveAt(num);
                            bool flag4 = (a == "1" && this.radioButtonLine1.Checked) || (a == "2" && this.radioButtonLine2.Checked);
                            if (flag4)
                            {
                                this.SaveDataAll();
                            }
                            MessageBox.Show(string.Format("Đã xóa dòng {0}", num), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                        }
                        else
                        {
                            MessageBox.Show("Không thể xóa dòng mới", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        }
                    }
                    else
                    {
                        MessageBox.Show("Index không hợp lệ", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Lỗi khi xóa dòng: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                }
            }
        }
        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            // Kiểm tra xem có click vào cột button không
            if (e.RowIndex >= 0) // Đảm bảo không click vào header
            {
                DataGridView dgv = sender as DataGridView;

                // Lấy mã đơn hàng từ row được click
                string maDH = GetMaDHFromDisplayRow(e.RowIndex);

                if (string.IsNullOrEmpty(maDH))
                {
                    //MessageBox.Show("Không thể xác định mã đơn hàng!", "Lỗi",
                    //               MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                // Lấy tên khách hàng từ row được click
                string tenKH = GetTenKHFromDisplayRow(e.RowIndex);
                if (string.IsNullOrEmpty(tenKH))
                {
                    //MessageBox.Show("Không thể xác định tên khách hàng!", "Lỗi",
                    //               MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                // Lấy index thực tế trong originalData
                int realIndex = GetRealIndexByMaDH(maDH, tenKH);

                if (realIndex == -1)
                {
                    //MessageBox.Show($"Không tìm thấy index thực tế cho mã đơn hàng: {maDH}", "Lỗi",
                    //               MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // Kiểm tra tên cột được click
                string columnName = dgv.Columns[e.ColumnIndex].Name;
                switch (columnName)
                {
                    case "chat": // Cột Line
                        HandleLineButtonClick(realIndex); // Truyền cả realIndex và maDH
                        break;
                    case "findIphone": // Cột Rung
                        HandleRungButtonClick(realIndex);
                        break;
                    case "tatRung": // Cột Tắt Rung
                        HandleOffRungButtonClick(realIndex);
                        break;
                    case "loiNhac": // Cột Lời Nhắc
                        HandleLoiNhacButtonClick(realIndex);
                        break;
                    case "deleteRow": // Cột Lời Nhắc
                        HandleXoaButtonClick(realIndex);
                        break;
                    case "Move": // Cột Lời Nhắc
                        HandleMoveButtonClick(realIndex);
                        break;

                }
            }
        }

        private void HandleMoveButtonClick(int rowIndex)
        {
            Dictionary<string, Form1.DebtReminderManager.DebtReminderInfo> dictionary = new Dictionary<string, Form1.DebtReminderManager.DebtReminderInfo>();
            int num = this.FindDisplayRowIndexByRealIndex(rowIndex);
            bool flag = num == -1;
            if (!flag)
            {
                try
                {
                    bool flag2 = num >= 0 && num < this.dataGridView1.Rows.Count;
                    if (flag2)
                    {
                        bool flag3 = !this.dataGridView1.Rows[num].IsNewRow;
                        if (flag3)
                        {
                            string dataDeleteFromMongoDB = this.GetDataDeleteFromMongoDB();
                            DataGridViewCell dataGridViewCell = this.dataGridView1.Rows[num].Cells["MaHD"];
                            string text;
                            if (dataGridViewCell == null)
                            {
                                text = null;
                            }
                            else
                            {
                                object value = dataGridViewCell.Value;
                                text = ((value != null) ? value.ToString() : null);
                            }
                            string text2 = text;
                            DataGridViewCell dataGridViewCell2 = this.dataGridView1.Rows[num].Cells[1];
                            string text3;
                            if (dataGridViewCell2 == null)
                            {
                                text3 = null;
                            }
                            else
                            {
                                object value2 = dataGridViewCell2.Value;
                                text3 = ((value2 != null) ? value2.ToString() : null);
                            }
                            string text4 = text3 ?? "";
                            DataGridViewCell dataGridViewCell3 = this.dataGridView1.Rows[num].Cells[2];
                            string text5;
                            if (dataGridViewCell3 == null)
                            {
                                text5 = null;
                            }
                            else
                            {
                                object value3 = dataGridViewCell3.Value;
                                text5 = ((value3 != null) ? value3.ToString() : null);
                            }
                            string text6 = text5 ?? "";
                            string key = text2 + "|" + text4;
                            string date = "";
                            DataGridViewCell dataGridViewCell4 = this.dataGridView1.Rows[num].Cells[14];
                            object obj = (dataGridViewCell4 != null) ? dataGridViewCell4.Value : null;
                            bool flag4 = obj != null && obj != DBNull.Value;
                            if (flag4)
                            {
                                DateTime dateTime = DateTime.MinValue;
                                bool flag5;
                                if (obj is DateTime)
                                {
                                    dateTime = (DateTime)obj;
                                    flag5 = true;
                                }
                                else
                                {
                                    flag5 = false;
                                }
                                bool flag6 = flag5;
                                if (flag6)
                                {
                                    date = dateTime.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture);
                                }
                                else
                                {
                                    string text7 = obj.ToString().Trim();
                                    bool flag7 = !string.IsNullOrEmpty(text7);
                                    if (flag7)
                                    {
                                        string[] array = new string[]
                                        {
                                    "dd/MM/yyyy",
                                    "d/MM/yyyy",
                                    "dd/M/yyyy",
                                    "d/M/yyyy",
                                    "dd-MM-yyyy",
                                    "d-MM-yyyy",
                                    "dd-M-yyyy",
                                    "d-M-yyyy",
                                    "yyyy-MM-dd",
                                    "dd/MM/yy",
                                    "d/MM/yy",
                                    "dd/M/yy",
                                    "d/M/yy"
                                        };
                                        bool flag8 = false;
                                        foreach (string format in array)
                                        {
                                            DateTime dateTime2;
                                            bool flag9 = DateTime.TryParseExact(text7, format, CultureInfo.InvariantCulture, DateTimeStyles.None, out dateTime2);
                                            if (flag9)
                                            {
                                                date = dateTime2.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture);
                                                flag8 = true;
                                                break;
                                            }
                                        }
                                        bool flag10 = !flag8;
                                        if (flag10)
                                        {
                                            CultureInfo provider = new CultureInfo("vi-VN");
                                            DateTime dateTime2;
                                            bool flag11 = DateTime.TryParse(text7, provider, DateTimeStyles.None, out dateTime2);
                                            if (flag11)
                                            {
                                                date = dateTime2.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture);
                                                flag8 = true;
                                            }
                                        }
                                        bool flag12 = !flag8;
                                        if (flag12)
                                        {
                                            date = text7;
                                        }
                                    }
                                }
                            }
                            Form1.DebtReminderManager.DebtReminderInfo debtReminderInfo = new Form1.DebtReminderManager.DebtReminderInfo();
                            debtReminderInfo.MaHD = text2;
                            debtReminderInfo.TenKh = text4;
                            DataGridViewCell dataGridViewCell5 = this.dataGridView1.Rows[num].Cells[10];
                            string text8;
                            if (dataGridViewCell5 == null)
                            {
                                text8 = null;
                            }
                            else
                            {
                                object value4 = dataGridViewCell5.Value;
                                text8 = ((value4 != null) ? value4.ToString() : null);
                            }
                            debtReminderInfo.userIcloud = (text8 ?? "");
                            DataGridViewCell dataGridViewCell6 = this.dataGridView1.Rows[num].Cells[11];
                            string text9;
                            if (dataGridViewCell6 == null)
                            {
                                text9 = null;
                            }
                            else
                            {
                                object value5 = dataGridViewCell6.Value;
                                text9 = ((value5 != null) ? value5.ToString() : null);
                            }
                            debtReminderInfo.PassIcloud = (text9 ?? "");
                            DataGridViewCell dataGridViewCell7 = this.dataGridView1.Rows[num].Cells[9];
                            string text10;
                            if (dataGridViewCell7 == null)
                            {
                                text10 = null;
                            }
                            else
                            {
                                object value6 = dataGridViewCell7.Value;
                                text10 = ((value6 != null) ? value6.ToString() : null);
                            }
                            debtReminderInfo.NamePhoneChange = (text10 ?? "");
                            DataGridViewCell dataGridViewCell8 = this.dataGridView1.Rows[num].Cells[3];
                            string text11;
                            if (dataGridViewCell8 == null)
                            {
                                text11 = null;
                            }
                            else
                            {
                                object value7 = dataGridViewCell8.Value;
                                text11 = ((value7 != null) ? value7.ToString() : null);
                            }
                            debtReminderInfo.Note = (text11 ?? "");
                            debtReminderInfo.Line = text6;
                            debtReminderInfo.date = date;
                            Form1.DebtReminderManager.DebtReminderInfo value8 = debtReminderInfo;
                            bool flag13 = !string.IsNullOrWhiteSpace(text2);
                            if (flag13)
                            {
                                dictionary[key] = value8;
                                string text12 = JsonConvert.SerializeObject(dictionary, Formatting.Indented);
                                this.dataGridView1.Rows.RemoveAt(num);
                                bool flag14 = (text6 == "1" && this.radioButtonxoa.Checked) || (text6 == "1" && this.radioButtonAll.Checked);
                                if (flag14)
                                {
                                    string dataFromMongoDB = this.GetDataFromMongoDB1();
                                    string jsonContent = Form1.MergeJsonWithNewtonsoft(dataFromMongoDB, text12);
                                    this.reminderManager.SaveDataBackMongoDB1(jsonContent);
                                }
                                else
                                {
                                    bool flag15 = (text6 == "2" && this.radioButtonxoa.Checked) || (text6 == "2" && this.radioButtonAll.Checked);
                                    if (flag15)
                                    {
                                        string dataFromMongoDB2 = this.GetDataFromMongoDB2();
                                        string jsonContent2 = Form1.MergeJsonWithNewtonsoft(dataFromMongoDB2, text12);
                                        this.reminderManager.SaveDataBackMongoDB2(jsonContent2);
                                    }
                                }
                                bool flag16 = !string.IsNullOrEmpty(text12) && !this.radioButtonAll.Checked;
                                if (flag16)
                                {
                                    string jsonContent3 = Form1.MergeJsonWithNewtonsoft(dataDeleteFromMongoDB, text12);
                                    this.reminderManager.SaveDataDeleteMongoDB(jsonContent3);
                                }
                                bool flag17 = (text6 == "1" && this.radioButtonLine1.Checked) || (text6 == "2" && this.radioButtonLine2.Checked);
                                if (flag17)
                                {
                                    this.SaveDataAll();
                                }
                            }
                            else
                            {
                                MessageBox.Show("Không thể di chuyển dòng: Mã hóa đơn không hợp lệ", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            }
                        }
                        else
                        {
                            MessageBox.Show("Không thể di chuyển dòng mới", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        }
                    }
                    else
                    {
                        MessageBox.Show("Chỉ số dòng không hợp lệ", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Lỗi khi chuyển dữ liệu: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                }
            }
        }

        private static string MergeJsonWithNewtonsoft(string json1, string json2)
        {
            string result;
            try
            {
                bool flag = string.IsNullOrEmpty(json1) && string.IsNullOrEmpty(json2);
                if (flag)
                {
                    result = "{}";
                }
                else
                {
                    bool flag2 = string.IsNullOrEmpty(json1);
                    if (flag2)
                    {
                        result = (string.IsNullOrEmpty(json2) ? "{}" : json2);
                    }
                    else
                    {
                        bool flag3 = string.IsNullOrEmpty(json2);
                        if (flag3)
                        {
                            result = json1;
                        }
                        else
                        {
                            JObject jobject = JObject.Parse(json1);
                            JObject jobject2 = JObject.Parse(json2);
                            jobject.Merge(jobject2, new JsonMergeSettings
                            {
                                MergeArrayHandling = MergeArrayHandling.Union  // Sửa từ 1 thành enum
                            });
                            result = jobject.ToString(Formatting.Indented);  // Sửa cả 2 tham số
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Lỗi khi merge JSON: " + ex.Message);
                result = null;
            }
            return result;
        }
        // Thêm Dictionary để track Chrome driver theo row index
        private Dictionary<int, IWebDriver> rowDriverMap = new Dictionary<int, IWebDriver>();
        // Class để quản lý từng process riêng biệt
        public class iCloudProcess
        {
            public int RowIndex { get; set; }
            public string UserIcloud { get; set; }
            public string PassIcloud { get; set; }
            public string DeviceInfo { get; set; }
            public int Pro5 { get; set; }
            public System.Timers.Timer Timer { get; set; }
            public CancellationTokenSource CancellationTokenSource { get; set; }
            public bool IsRunning { get; set; }
            public IWebDriver Driver { get; set; }
        }
        private int FindDisplayRowIndexByRealIndex(int realIndex)
        {
            // Lấy mã đơn hàng và tên khách hàng từ originalData dựa trên realIndex
            if (realIndex < 0 || realIndex >= originalData.Count)
                return -1;

            string[] parts = originalData[realIndex].Split('|');
            if (parts.Length < 2) // Cần ít nhất maDH và tenKH
                return -1;

            string targetMaDH = parts[0];
            string targetTenKH = parts[1];

            // Tìm trong dataGridView1 hiện tại bằng cả maDH và tenKH
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                var maDHValue = dataGridView1.Rows[i].Cells[0].Value;
                var tenKHValue = dataGridView1.Rows[i].Cells[1].Value;

                if (maDHValue != null && tenKHValue != null &&
                    maDHValue.ToString() == targetMaDH &&
                    tenKHValue.ToString() == targetTenKH)
                {
                    return i;
                }
            }
            return -1; // Không tìm thấy
        }
        private string CreateCompositeKey(string maDH, string tenKH)
        {
            return $"{maDH}|{tenKH}";
        }
        // Thêm hàm lấy tên khách hàng từ row hiển thị
        private string GetTenKHFromDisplayRow(int displayRowIndex)
        {
            if (displayRowIndex >= 0 && displayRowIndex < dataGridView1.Rows.Count)
            {
                var cellValue = dataGridView1.Rows[displayRowIndex].Cells[1].Value; // Cột 1 là tenKH
                if (cellValue != null)
                    return cellValue.ToString();
            }
            return "";
        }
        // Hàm lấy index thực tế từ mã đơn hàng
        public int GetRealIndexByMaDH(string maDH, string tenKH)
        {
            // Phương pháp 1: Thử tìm trong originalIndexMap trước (nhanh hơn)
            string compositeKey = CreateCompositeKey(maDH, tenKH);
            if (originalIndexMap != null && originalIndexMap.ContainsKey(compositeKey))
            {
                return originalIndexMap[compositeKey];
            }

            // Phương pháp 2: Nếu không tìm thấy trong originalIndexMap, tìm trong originalDataNew
            if (originalDataNew != null)
            {
                int index = 0;
                foreach (var kvp in originalDataNew)
                {
                    if (kvp.Key == maDH && kvp.Value.TenKh == tenKH)
                    {
                        return index;
                    }
                    index++;
                }
            }

            // Phương pháp 3: Cuối cùng, tìm trong originalData (List<string>)
            if (originalData != null)
            {
                for (int i = 0; i < originalData.Count; i++)
                {
                    string[] parts = originalData[i].Split('|');
                    if (parts.Length >= 2 && parts[0] == maDH && parts[1] == tenKH)
                    {
                        return i;
                    }
                }
            }

            return -1; // Không tìm thấy ở đâu cả
        }

        // Hàm lưu trạng thái hiện tại trước khi thay đổi hiển thị
        private void SaveCurrentStatus()
        {
            for (int i = 0; i < this.dataGridView1.Rows.Count; i++)
            {
                bool flag = this.dataGridView1.Rows[i].Cells[0].Value != null && this.dataGridView1.Rows[i].Cells[1].Value != null && this.dataGridView1.Rows[i].Cells[8].Value != null;
                if (flag)
                {
                    string maDH = this.dataGridView1.Rows[i].Cells[0].Value.ToString();
                    string tenKH = this.dataGridView1.Rows[i].Cells[1].Value.ToString();
                    string value = this.dataGridView1.Rows[i].Cells[8].Value.ToString();
                    object value2 = this.dataGridView1.Rows[i].Cells[9].Value;
                    string text;
                    if (value2 == null)
                    {
                        text = null;
                    }
                    else
                    {
                        string text2 = value2.ToString();
                        text = ((text2 != null) ? text2.Trim().ToUpper() : null);
                    }
                    string value3 = text ?? "";
                    string key = this.CreateCompositeKey(maDH, tenKH);
                    bool flag2 = this.statusData.ContainsKey(key);
                    if (flag2)
                    {
                        this.statusData[key] = value;
                    }
                    else
                    {
                        this.statusData.Add(key, value);
                    }
                    bool flag3 = this.namePhoneData.ContainsKey(key);
                    if (flag3)
                    {
                        this.namePhoneData[key] = value3;
                    }
                    else
                    {
                        this.namePhoneData.Add(key, value3);
                    }
                }
            }
        }
        // Hàm khôi phục trạng thái sau khi hiển thị dữ liệu
        private void RestoreStatus()
        {
            for (int i = 0; i < this.dataGridView1.Rows.Count; i++)
            {
                bool flag = this.dataGridView1.Rows[i].Cells[0].Value != null && this.dataGridView1.Rows[i].Cells[1].Value != null;
                if (flag)
                {
                    string maDH = this.dataGridView1.Rows[i].Cells[0].Value.ToString();
                    string tenKH = this.dataGridView1.Rows[i].Cells[1].Value.ToString();
                    string key = this.CreateCompositeKey(maDH, tenKH);
                    bool flag2 = this.statusData.ContainsKey(key);
                    if (flag2)
                    {
                        this.dataGridView1.Rows[i].Cells[8].Value = this.statusData[key];
                    }
                    bool flag3 = this.namePhoneData.ContainsKey(key);
                    if (flag3)
                    {
                        this.dataGridView1.Rows[i].Cells[9].Value = this.namePhoneData[key];
                    }
                }
            }
        }
        private void LoadAndSaveData(List<string> newData)
        {
            bool flag = newData == null;
            if (!flag)
            {
                int count = Form1.originalData.Count;
                Form1.originalData.AddRange(newData);
                for (int i = 0; i < newData.Count; i++)
                {
                    string[] array = newData[i].Split(new char[]
                    {
                        '|'
                    });
                    bool flag2 = array.Length >= 2;
                    if (flag2)
                    {
                        string maDH = array[0];
                        string tenKH = array[1];
                        string key = this.CreateCompositeKey(maDH, tenKH);
                        int value = count + i;
                        bool flag3 = this.originalIndexMap.ContainsKey(key);
                        if (flag3)
                        {
                            this.originalIndexMap[key] = value;
                        }
                        else
                        {
                            this.originalIndexMap.Add(key, value);
                        }
                    }
                }
                this.DisplayData(Form1.originalData);
            }
        }
        // Hàm lấy mã đơn hàng từ row hiển thị
        private string GetMaDHFromDisplayRow(int displayRowIndex)
        {

            if (displayRowIndex >= 0 && displayRowIndex < dataGridView1.Rows.Count)
            {
                var cellValue = dataGridView1.Rows[displayRowIndex].Cells[0].Value;
                if (cellValue != null)
                    return cellValue.ToString();
            }
            return "";
        }
        // Thêm vào class chính của bạn
        private Dictionary<int, iCloudProcess> activeProcesses = new Dictionary<int, iCloudProcess>();
        private void HandleOffRungButtonClick(int rowIndex)
        {
            try
            {
                int num = this.FindDisplayRowIndexByRealIndex(rowIndex);
                bool flag = num == -1;
                if (!flag)
                {
                    object value = this.dataGridView1.Rows[num].Cells[1].Value;
                    string text = ((value != null) ? value.ToString() : null) ?? "";
                    bool flag2 = string.IsNullOrEmpty(text);
                    if (flag2)
                    {
                        MessageBox.Show("Không có thông tin thiết bị để tắt", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                    }
                    else
                    {
                        DialogResult dialogResult = MessageBox.Show("Bạn có muốn tắt Chrome đang chạy cho " + text + "?", "Xác nhận tắt", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        bool flag3 = dialogResult != DialogResult.Yes;
                        if (!flag3)
                        {
                            bool flag4 = false;
                            bool flag5 = this.activeProcesses.ContainsKey(rowIndex);
                            if (flag5)
                            {
                                Form1.iCloudProcess iCloudProcess = this.activeProcesses[rowIndex];
                                iCloudProcess.IsRunning = false;
                                CancellationTokenSource cancellationTokenSource = iCloudProcess.CancellationTokenSource;
                                if (cancellationTokenSource != null)
                                {
                                    cancellationTokenSource.Cancel();
                                }
                                flag4 = true;
                                this.richTextBox1.AppendText(string.Format("[{0:HH:mm:ss}] Đã gửi lệnh dừng process cho Row {1} - Device: {2}\r\n", DateTime.Now, rowIndex, text));
                            }
                            IWebDriver webDriver = null;
                            List<IWebDriver> obj = Form1.activeDrivers;
                            lock (obj)
                            {
                                bool flag7 = this.rowDriverMap.ContainsKey(rowIndex);
                                if (flag7)
                                {
                                    bool flag8 = false;
                                    bool flag9 = this.activeProcesses.ContainsKey(rowIndex);
                                    if (flag9)
                                    {
                                        Form1.iCloudProcess iCloudProcess2 = this.activeProcesses[rowIndex];
                                        bool flag10 = iCloudProcess2.DeviceInfo == text;
                                        if (flag10)
                                        {
                                            flag8 = true;
                                            this.richTextBox1.AppendText(string.Format("[{0:HH:mm:ss}] ✓ Xác nhận đúng device: {1}\r\n", DateTime.Now, text));
                                        }
                                        else
                                        {
                                            this.richTextBox1.AppendText(string.Format("[{0:HH:mm:ss}] ⚠ Device không khớp - Expected: {1}, Found: {2}\r\n", DateTime.Now, text, iCloudProcess2.DeviceInfo));
                                        }
                                    }
                                    bool flag11 = flag8 || !this.activeProcesses.ContainsKey(rowIndex);
                                    if (flag11)
                                    {
                                        webDriver = this.rowDriverMap[rowIndex];
                                        this.rowDriverMap.Remove(rowIndex);
                                        this.richTextBox1.AppendText(string.Format("[{0:HH:mm:ss}] Tìm thấy driver qua rowDriverMap cho Row {1} - Device: {2}\r\n", DateTime.Now, rowIndex, text));
                                    }
                                    else
                                    {
                                        this.richTextBox1.AppendText(string.Format("[{0:HH:mm:ss}] ✗ Bỏ qua driver do device không khớp cho Row {1}\r\n", DateTime.Now, rowIndex));
                                    }
                                }
                                bool flag12 = webDriver != null;
                                if (flag12)
                                {
                                    bool flag13 = true;
                                    bool flag14 = this.chromeStartTimes.ContainsKey(rowIndex);
                                    if (flag14)
                                    {
                                        DateTime d = this.chromeStartTimes[rowIndex];
                                        double totalSeconds = (DateTime.Now - d).TotalSeconds;
                                        this.richTextBox1.AppendText(string.Format("[{0:HH:mm:ss}] Thời gian từ lúc khởi chạy Chrome: {1:F1} giây\r\n", DateTime.Now, totalSeconds));
                                        bool flag15 = totalSeconds <= 25.0;
                                        if (flag15)
                                        {
                                            flag13 = false;
                                            this.richTextBox1.AppendText(string.Format("[{0:HH:mm:ss}] ⚡ Tắt nhanh (≤25s): Bỏ qua kiểm tra VerifyDriverIsCorrect\r\n", DateTime.Now));
                                        }
                                        else
                                        {
                                            this.richTextBox1.AppendText(string.Format("[{0:HH:mm:ss}] \ud83d\udd0d Tắt chậm (>25s): Sẽ chạy VerifyDriverIsCorrect ngay lập tức\r\n", DateTime.Now));
                                        }
                                    }
                                    else
                                    {
                                        this.richTextBox1.AppendText(string.Format("[{0:HH:mm:ss}] ⚠ Không tìm thấy thời gian khởi chạy, sẽ chạy VerifyDriverIsCorrect\r\n", DateTime.Now));
                                    }
                                    bool flag16 = true;
                                    bool flag17 = flag13;
                                    if (flag17)
                                    {
                                        flag16 = this.VerifyDriverIsCorrect(webDriver, text, rowIndex);
                                    }
                                    bool flag18 = flag16;
                                    if (flag18)
                                    {
                                        Form1.activeDrivers.Remove(webDriver);
                                        webDriver.Quit();
                                        this.richTextBox1.AppendText(string.Format("[{0:HH:mm:ss}] ✓ Đã tắt Chrome driver cho Row {1} - Device: {2}\r\n", DateTime.Now, rowIndex, text));
                                        bool flag19 = this.chromeStartTimes.ContainsKey(rowIndex);
                                        if (flag19)
                                        {
                                            this.chromeStartTimes.Remove(rowIndex);
                                        }
                                        this.DeleteChromeProfile(text);
                                    }
                                    else
                                    {
                                        this.richTextBox1.AppendText(string.Format("[{0:HH:mm:ss}] ✗ Driver không khớp với thiết bị {1}, không tắt để tránh nhầm lẫn\r\n", DateTime.Now, text));
                                        MessageBox.Show("Driver không khớp với thiết bị " + text + ". Không thực hiện tắt để tránh nhầm lẫn.", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                    }
                                }
                                else
                                {
                                    this.richTextBox1.AppendText(string.Format("[{0:HH:mm:ss}] ✗ Lỗi khi tắt Chrome driver: Không tắt được chrome\r\n", DateTime.Now));
                                }
                            }
                            bool flag20 = this.activeProcesses.ContainsKey(rowIndex);
                            if (flag20)
                            {
                                this.activeProcesses.Remove(rowIndex);
                            }
                            this.dataGridView1.Rows[num].Cells[8].Value = "Đã tắt";
                            bool flag21 = flag4 && this.activeProcesses.ContainsKey(rowIndex);
                            if (flag21)
                            {
                                Form1.iCloudProcess iCloudProcess3 = this.activeProcesses[rowIndex];
                            }
                            this.richTextBox1.AppendText(string.Format("[{0:HH:mm:ss}] ✓ Hoàn tất tắt process cho Row {1} - Device: {2}\r\n", DateTime.Now, rowIndex, text));
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                this.richTextBox1.AppendText(string.Format("[{0:HH:mm:ss}] ✗ Lỗi tổng quát trong HandleOffRungButtonClick: {1}\r\n", DateTime.Now, ex.Message));
                MessageBox.Show("Đã xảy ra lỗi: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Hand);
            }
        }
        private bool VerifyDriverIsCorrect(IWebDriver driver, string deviceInfo, int rowIndex)
        {
            bool result;
            try
            {
                bool flag = driver == null;
                if (flag)
                {
                    this.richTextBox1.AppendText(string.Format("[{0:HH:mm:ss}] ✗ Driver is null\r\n", DateTime.Now));
                    result = false;
                }
                else
                {
                    WebDriverWait webDriverWait = new WebDriverWait(driver, TimeSpan.FromSeconds(10.0));
                    string url = driver.Url;
                    bool flag2 = !url.Contains("icloud.com") || !url.Contains("find");
                    if (flag2)
                    {
                        this.richTextBox1.AppendText(string.Format("[{0:HH:mm:ss}] ✗ Driver không ở trang Find My: {1}\r\n", DateTime.Now, url));
                        result = false;
                    }
                    else
                    {
                        object value = this.dataGridView1.Rows[rowIndex].Cells[9].Value;
                        string text;
                        if (value == null)
                        {
                            text = null;
                        }
                        else
                        {
                            string text2 = value.ToString();
                            text = ((text2 != null) ? text2.Trim().ToUpper() : null);
                        }
                        string a = text ?? "";
                        bool flag3 = false;
                        IWebElement webElement = webDriverWait.Until<IWebElement>(ExpectedConditions.ElementIsVisible(By.CssSelector(".device-details-view .name")));
                        string text3 = webElement.Text.Trim().ToUpper();
                        string value2 = Regex.Match(deviceInfo, "\\d+").Value;
                        string value3 = Regex.Match(text3, "\\d+").Value;
                        bool flag4 = (value3.Length > 0 && string.Equals(value3, value2, StringComparison.OrdinalIgnoreCase)) || a == text3;
                        bool flag5 = flag4;
                        if (flag5)
                        {
                            flag3 = true;
                        }
                        bool flag6 = flag3;
                        if (flag6)
                        {
                            this.richTextBox1.AppendText(string.Format("[{0:HH:mm:ss}] ✓ Driver được xác nhận đúng cho thiết bị: {1}\r\n", DateTime.Now, deviceInfo));
                            result = true;
                        }
                        else
                        {
                            this.richTextBox1.AppendText(string.Format("[{0:HH:mm:ss}] ✗ Không tìm thấy thiết bị khớp với: {1}\r\n", DateTime.Now, deviceInfo));
                            result = false;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                this.richTextBox1.AppendText(string.Format("[{0:HH:mm:ss}] ✗ Lỗi trong VerifyDriverIsCorrect: {1}\r\n", DateTime.Now, ex.Message));
                result = false;
            }
            return result;
        }
        private void TatChromeKhiLoi(int rowIndex)
        {
            try
            {
                int displayRowIndex = FindDisplayRowIndexByRealIndex(rowIndex);
                if (displayRowIndex == -1)
                {

                    return;
                }
                string deviceInfo = dataGridView1.Rows[displayRowIndex].Cells[1].Value?.ToString() ?? "";

                // QUAN TRỌNG: Dừng process trước khi tắt driver
                if (activeProcesses.ContainsKey(rowIndex))
                {
                    var process = activeProcesses[rowIndex];

                    // Dừng vòng lặp
                    process.IsRunning = false;

                    // Hủy CancellationToken để dừng Task
                    process.CancellationTokenSource?.Cancel();
                    richTextBox1.Invoke(new Action(() => {
                        richTextBox1.AppendText($"[{DateTime.Now:HH:mm:ss}] Đã tắt Chrome driver cho Row {rowIndex} vì không đăng nhập được vào icloud\r\n");
                    }));
                    
                }

                // Tìm và tắt Chrome driver của row này
                IWebDriver driverToClose = null;

                lock (activeDrivers)
                {
                    // Cách 1: Sử dụng rowDriverMap nếu có
                    if (rowDriverMap.ContainsKey(rowIndex))
                    {
                        driverToClose = rowDriverMap[rowIndex];
                        rowDriverMap.Remove(rowIndex);
                    }
                    else
                    {
                        // Cách 2: Tìm driver dựa vào profile path (backup method)
                        foreach (var driver in activeDrivers.ToList())
                        {
                            try
                            {
                                // Kiểm tra URL hoặc title có chứa thông tin thiết bị không
                                string currentUrl = driver.Url;
                                if (currentUrl.Contains("icloud.com"))
                                {
                                    // Giả sử chỉ có 1 driver đang chạy cho thiết bị này
                                    driverToClose = driver;
                                    break;
                                }
                            }
                            catch (Exception)
                            {
                                // Driver có thể đã bị đóng, bỏ qua
                                continue;
                            }
                        }
                    }

                    if (driverToClose != null)
                    {
                        try
                        {
                            activeDrivers.Remove(driverToClose);
                            driverToClose.Quit();
                            richTextBox1.Invoke(new Action(() => {
                                richTextBox1.AppendText($"[{DateTime.Now:HH:mm:ss}] Đã tắt Chrome driver cho Row {rowIndex} vì không đăng nhập được vào icloud\r\n");
                            }));
                            
                            DeleteChromeProfile(deviceInfo);
                        }
                        catch (Exception ex)
                        {
                            richTextBox1.Invoke(new Action(() => {
                                richTextBox1.AppendText($"[{DateTime.Now:HH:mm:ss}] Lỗi khi tắt Chrome khi không đăng nhập được icloud: {ex.Message}\r\n");
                            }));
                            
                        }
                    }
                }

                // Cập nhật status trong DataGridView
                dataGridView1.Rows[displayRowIndex].Cells[8].Value = "Lỗi Icloud";
                richTextBox1.Invoke(new Action(() => {
                    richTextBox1.AppendText($"[{DateTime.Now:HH:mm:ss}] Đã tắt hoàn toàn process cho {deviceInfo}\r\n");
                }));
                
            }
            catch (Exception ex)
            {
                richTextBox1.Invoke(new Action(() => {
                    richTextBox1.AppendText($"[{DateTime.Now:HH:mm:ss}] Lỗi tổng quát: {ex.Message}\r\n");
                }));
                
            }
        }
        private async void HandleRungButtonClick(int rowIndex)
        {
            int displayRowIndex = FindDisplayRowIndexByRealIndex(rowIndex);
            if (displayRowIndex == -1)
            {
                
                return;
            }
           
            string deviceInfo = "";
            string userIcloud = "";
            string passIcloud = "";
            string usercloudinData = "";
            string passcloudinData = "";
            

            // Kiểm tra xem row này đã có process chạy chưa
            if (activeProcesses.ContainsKey(rowIndex) && activeProcesses[rowIndex].IsRunning)
            {
                MessageBox.Show($"Row {rowIndex} đã có process đang chạy. Vui lòng tắt trước khi chạy lại.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                // Lấy position tiếp theo thay vì clear queue
                int pro5 = ChromePositionManager.GetNextPosition();
                this.Invoke((MethodInvoker)delegate ()
                {
                    deviceInfo = dataGridView1.Rows[displayRowIndex].Cells[1].Value?.ToString() ?? "";
                    userIcloud = comboBoxUsername.SelectedItem?.ToString();
                    passIcloud = textBoxPassIcloud.Text?.Trim();
                    usercloudinData = dataGridView1.Rows[rowIndex].Cells[10].Value?.ToString()?.Trim() ?? "";
                    passcloudinData = dataGridView1.Rows[rowIndex].Cells[11].Value?.ToString()?.Trim() ?? "";
                });
                if (string.IsNullOrEmpty(deviceInfo))
                {
                    MessageBox.Show("Thông tin khách hàng đang trống, hãy kiểm tra lại...", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                if (usercloudinData == "" || passcloudinData == "")
                {
                    if (string.IsNullOrEmpty(userIcloud) || string.IsNullOrEmpty(passIcloud))
                    {
                        MessageBox.Show("Vui lòng nhập đầy đủ thông tin iCloud", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                }
                else
                {
                    userIcloud = usercloudinData;
                    passIcloud = passcloudinData;
                }
                

                // Xác nhận trước khi thực hiện
                var confirmResult = MessageBox.Show($"Bạn có muốn phát âm thanh cho {deviceInfo}?",
                                                  "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (confirmResult != DialogResult.Yes)
                {
                    return;
                }

                // Tạo process mới cho row này
                var process = new iCloudProcess
                {
                    RowIndex = rowIndex,
                    UserIcloud = userIcloud,
                    PassIcloud = passIcloud,
                    DeviceInfo = deviceInfo,
                    Pro5 = pro5,
                    IsRunning = true,
                    CancellationTokenSource = new CancellationTokenSource()
                };

                activeProcesses[rowIndex] = process;

                // Cập nhật status trong DataGridView
                dataGridView1.Rows[displayRowIndex].Cells[8].Value = "Đang chạy";

                // Chạy process trong background
                await Task.Run(() => RunContinuousProcess(process));
            }
            catch (Exception ex)
            {
                richTextBox1.AppendText($"[{DateTime.Now}] Lỗi tổng quát: {ex.Message}\n");
                MessageBox.Show($"Đã xảy ra lỗi: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async void runToDie(int rowIndex)
        {
            int displayRowIndex = FindDisplayRowIndexByRealIndex(rowIndex);
            if (displayRowIndex == -1)
            {

                return;
            }

            string deviceInfo = "";
            string userIcloud = "";
            string passIcloud = "";
            string usercloudinData = "";
            string passcloudinData = "";


            // Kiểm tra xem row này đã có process chạy chưa
            if (activeProcesses.ContainsKey(rowIndex) && activeProcesses[rowIndex].IsRunning)
            {
                MessageBox.Show($"Row {rowIndex} đã có process đang chạy. Vui lòng tắt trước khi chạy lại.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                // Lấy position tiếp theo thay vì clear queue
                int pro5 = ChromePositionManager.GetNextPosition();
                this.Invoke((MethodInvoker)delegate ()
                {
                    deviceInfo = dataGridView1.Rows[displayRowIndex].Cells[1].Value?.ToString() ?? "";
                    userIcloud = comboBoxUsername.SelectedItem?.ToString();
                    passIcloud = textBoxPassIcloud.Text?.Trim();
                    usercloudinData = dataGridView1.Rows[rowIndex].Cells[10].Value?.ToString()?.Trim() ?? "";
                    passcloudinData = dataGridView1.Rows[rowIndex].Cells[11].Value?.ToString()?.Trim() ?? "";
                });
                if (string.IsNullOrEmpty(deviceInfo))
                {
                    MessageBox.Show("Thông tin khách hàng đang trống, hãy kiểm tra lại...", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                if (usercloudinData == "" || passcloudinData == "")
                {
                    if (string.IsNullOrEmpty(userIcloud) || string.IsNullOrEmpty(passIcloud))
                    {
                        MessageBox.Show("Vui lòng nhập đầy đủ thông tin iCloud", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        button14.Invoke((MethodInvoker)delegate ()
                        {
                            button14.Enabled = true;
                        });
                        return;
                    }

                }
                else
                {
                    userIcloud = usercloudinData;
                    passIcloud = passcloudinData;
                }


                //// Xác nhận trước khi thực hiện
                //var confirmResult = MessageBox.Show($"Bạn có muốn phát âm thanh cho {deviceInfo}?",
                //                                  "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                //if (confirmResult != DialogResult.Yes)
                //{
                //    return;
                //}

                // Tạo process mới cho row này
                var process = new iCloudProcess
                {
                    RowIndex = rowIndex,
                    UserIcloud = userIcloud,
                    PassIcloud = passIcloud,
                    DeviceInfo = deviceInfo,
                    Pro5 = pro5,
                    IsRunning = true,
                    CancellationTokenSource = new CancellationTokenSource()
                };

                activeProcesses[rowIndex] = process;

                // Cập nhật status trong DataGridView
                dataGridView1.Rows[displayRowIndex].Cells[8].Value = "Đang chạy";

                // Chạy process trong background
                await Task.Run(() => RunContinuousProcess(process));
            }
            catch (Exception ex)
            {
                richTextBox1.AppendText($"[{DateTime.Now}] Lỗi tổng quát: {ex.Message}\n");
                MessageBox.Show($"Đã xảy ra lỗi: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private async void RunContinuousProcess(iCloudProcess process)
        {
            richTextBox1.Invoke(new Action(() => {
                richTextBox1.AppendText($"[{DateTime.Now:HH:mm:ss}] Bắt đầu process cho Row {process.RowIndex} - {process.DeviceInfo}\n");
            }));

            // VÒNG LẶP để chạy liên tục
            while (process.IsRunning && !process.CancellationTokenSource.Token.IsCancellationRequested)
            {
                try
                {
                    richTextBox1.Invoke(new Action(() => {
                        richTextBox1.AppendText($"[{DateTime.Now:HH:mm:ss}] Row {process.RowIndex}: Bắt đầu chạy (1 tiếng)...\n");
                    }));

                    // Tạo CancellationToken với timeout 1 TIẾNG cho mỗi lần chạy
                    using (var timeoutCts = new CancellationTokenSource(TimeSpan.FromHours(1))) // FIX: Đổi từ 60 phút thành 1 tiếng
                    using (var combinedCts = CancellationTokenSource.CreateLinkedTokenSource(
                        process.CancellationTokenSource.Token, timeoutCts.Token))
                    {
                        try
                        {
                            // FIX: Sử dụng await thay vì Task.Wait() để tránh blocking
                            await Task.Run(() => ProcessiCloudRequest(
                                process.UserIcloud,
                                process.PassIcloud,
                                process.DeviceInfo,
                                process.Pro5,
                                process.RowIndex), combinedCts.Token);

                            richTextBox1.Invoke(new Action(() => {
                                richTextBox1.AppendText($"[{DateTime.Now:HH:mm:ss}] Row {process.RowIndex}: Hoàn thành trong thời gian cho phép\n");
                            }));
                            await DeleteChromeProfile(process.DeviceInfo);
                        }
                        catch (OperationCanceledException) when (timeoutCts.Token.IsCancellationRequested)
                        {
                            richTextBox1.Invoke(new Action(() => {
                                richTextBox1.AppendText($"[{DateTime.Now:HH:mm:ss}] Row {process.RowIndex}: Timeout sau 1 TIẾNG - Dừng và chuẩn bị chạy lại\n");
                            }));

                            // Dừng Chrome processes cho lần chạy này
                            await StopCurrentProcess(process.RowIndex);
                        }
                        catch (OperationCanceledException) when (process.CancellationTokenSource.Token.IsCancellationRequested)
                        {
                            richTextBox1.Invoke(new Action(() => {
                                richTextBox1.AppendText($"[{DateTime.Now:HH:mm:ss}] Row {process.RowIndex}: Process bị dừng bởi user\n");
                            }));
                            break; // Thoát vòng lặp
                        }
                    }

                    // Nghỉ 1 giây trước khi chạy lần tiếp theo (chỉ khi process vẫn running)
                    if (process.IsRunning && !process.CancellationTokenSource.Token.IsCancellationRequested)
                    {
                        richTextBox1.Invoke(new Action(() => {
                            richTextBox1.AppendText($"[{DateTime.Now:HH:mm:ss}] Row {process.RowIndex}: Nghỉ 1 giây trước khi chạy lại\n");
                        }));

                        try
                        {
                            await Task.Delay(1000, process.CancellationTokenSource.Token);
                        }
                        catch (OperationCanceledException)
                        {
                            richTextBox1.Invoke(new Action(() => {
                                richTextBox1.AppendText($"[{DateTime.Now:HH:mm:ss}] Row {process.RowIndex}: Process bị dừng trong lúc delay\n");
                            }));
                            break;
                        }
                    }
                }
                catch (Exception ex)
                {
                    richTextBox1.Invoke(new Action(() => {
                        richTextBox1.AppendText($"[{DateTime.Now:HH:mm:ss}] Row {process.RowIndex}: Lỗi - {ex.Message}\n");
                    }));

                    // Nghỉ 5 giây trước khi thử lại
                    if (process.IsRunning && !process.CancellationTokenSource.Token.IsCancellationRequested)
                    {
                        richTextBox1.Invoke(new Action(() => {
                            richTextBox1.AppendText($"[{DateTime.Now:HH:mm:ss}] Row {process.RowIndex}: Nghỉ 5 giây trước khi thử lại\n");
                        }));

                        try
                        {
                            await Task.Delay(5000, process.CancellationTokenSource.Token);
                        }
                        catch (OperationCanceledException)
                        {
                            richTextBox1.Invoke(new Action(() => {
                                richTextBox1.AppendText($"[{DateTime.Now:HH:mm:ss}] Row {process.RowIndex}: Process bị dừng trong lúc chờ retry\n");
                            }));
                            break;
                        }
                    }
                }
            }

            // Cleanup khi process kết thúc
            richTextBox1.Invoke(new Action(() => {
                richTextBox1.AppendText($"[{DateTime.Now:HH:mm:ss}] Row {process.RowIndex}: Kết thúc process và cleanup\n");
            }));

            CleanupProcess(process);
        }
        // Cập nhật lại method StopCurrentProcess
        private async Task StopCurrentProcess(int rowIndex)
        {
            try
            {
                int displayRowIndex = FindDisplayRowIndexByRealIndex(rowIndex);
                if (displayRowIndex == -1)
                {
                    
                    return;
                }
                // Dừng process trước
                if (activeProcesses.ContainsKey(rowIndex))
                {
                    var process = activeProcesses[rowIndex];
                    process.IsRunning = false;
                    process.CancellationTokenSource?.Cancel();
                }

                string deviceInfo = dataGridView1.Rows[displayRowIndex].Cells[1].Value?.ToString() ?? "";
                IWebDriver driverToClose = null;

                lock (activeDrivers)
                {
                    // Cách 1: Sử dụng rowDriverMap nếu có
                    if (rowDriverMap.ContainsKey(rowIndex))
                    {
                        driverToClose = rowDriverMap[rowIndex];
                        rowDriverMap.Remove(rowIndex);
                    }
                    else
                    {
                        // Cách 2: Tìm driver dựa vào profile path (backup method)
                        foreach (var driver in activeDrivers.ToList())
                        {
                            try
                            {
                                string currentUrl = driver.Url;
                                if (currentUrl.Contains("icloud.com"))
                                {
                                    driverToClose = driver;
                                    break;
                                }
                            }
                            catch (Exception)
                            {
                                continue;
                            }
                        }
                    }

                    if (driverToClose != null)
                    {
                        try
                        {
                            activeDrivers.Remove(driverToClose);
                            driverToClose.Quit();

                            // Chờ một chút để Chrome đóng hoàn toàn
                            Task.Delay(2000);

                            dataGridView1.Rows[displayRowIndex].Cells[8].Value = "Đã tắt";
                        }
                        catch (Exception ex)
                        {
                            richTextBox1.AppendText($"[{DateTime.Now:HH:mm:ss}] Lỗi khi tắt Chrome: {ex.Message}\r\n");
                        }
                    }
                }

                // XÓA THƯMỤC PROFILE CHROME
                await DeleteChromeProfile(deviceInfo);
            }
            catch (Exception ex)
            {
                richTextBox1.AppendText($"[{DateTime.Now:HH:mm:ss}] Lỗi StopCurrentProcess: {ex.Message}\r\n");
            }
        }

        // Thêm method mới để xóa Chrome profile
        private async Task DeleteChromeProfile(string deviceInfo)
        {
            try
            {
                string chromePortableDirectory = Path.Combine(Directory.GetCurrentDirectory(), "GoogleChromePortable", "App", "Chrome-bin");
                string profilePath = Path.Combine(chromePortableDirectory, "User_Data_" + deviceInfo);

                if (Directory.Exists(profilePath))
                {
                    // Chờ thêm một chút để đảm bảo Chrome đã giải phóng tất cả file
                    await Task.Delay(1000);

                    // Thử xóa với retry logic
                    int maxRetries = 3;
                    for (int i = 0; i < maxRetries; i++)
                    {
                        try
                        {
                            // Xóa thuộc tính readonly trước khi xóa
                            SetDirectoryAttributesNormal(profilePath);
                            Directory.Delete(profilePath, true);
                            break;
                        }
                        catch (UnauthorizedAccessException)
                        {
                            if (i == maxRetries - 1)
                            {
                                richTextBox1.Invoke(new Action(() => {
                                    richTextBox1.AppendText($"[{DateTime.Now:HH:mm:ss}] Không thể xóa profile của : {deviceInfo}\r\n");
                                }));
                            }
                            else
                            {
                                await Task.Delay(2000); // Chờ 2 giây trước khi thử lại
                            }
                        }
                        catch (DirectoryNotFoundException)
                        {
                            // Thư mục đã được xóa rồi
                            break;
                        }
                        catch (IOException ex)
                        {
                            if (i == maxRetries - 1)
                            {
                                richTextBox1.Invoke(new Action(() => {
                                    richTextBox1.AppendText($"[{DateTime.Now:HH:mm:ss}] Lỗi xóa profile: {ex.Message}\r\n");
                                }));
                            }
                            else
                            {
                                await Task.Delay(2000);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                richTextBox1.Invoke(new Action(() => {
                    richTextBox1.AppendText($"[{DateTime.Now:HH:mm:ss}] Lỗi DeleteChromeProfile: {ex.Message}\r\n");
                }));
            }
        }

        // Helper method để xóa thuộc tính readonly
        private void SetDirectoryAttributesNormal(string dirPath)
        {
            try
            {
                foreach (string file in Directory.GetFiles(dirPath, "*", SearchOption.AllDirectories))
                {
                    System.IO.File.SetAttributes(file, FileAttributes.Normal);
                }

                foreach (string dir in Directory.GetDirectories(dirPath, "*", SearchOption.AllDirectories))
                {
                    System.IO.File.SetAttributes(dir, FileAttributes.Normal);
                }

                System.IO.File.SetAttributes(dirPath, FileAttributes.Normal);
            }
            catch (Exception)
            {
                // Bỏ qua lỗi set attributes
            }
        }

        // Cập nhật method CleanupProcess để đảm bảo cleanup hoàn toàn
        private void CleanupProcess(iCloudProcess process)
        {
            try
            {
                // Đảm bảo process được đánh dấu dừng
                process.IsRunning = false;

                // Hủy CancellationToken nếu chưa hủy
                if (process.CancellationTokenSource != null && !process.CancellationTokenSource.Token.IsCancellationRequested)
                {
                    process.CancellationTokenSource.Cancel();
                }

                // Dispose CancellationTokenSource
                process.CancellationTokenSource?.Dispose();

                // Tắt Chrome driver
                if (process.Driver != null)
                {
                    try
                    {
                        lock (activeDrivers)
                        {
                            activeDrivers.Remove(process.Driver);
                        }
                        process.Driver.Quit();
                        process.Driver = null;
                    }
                    catch (Exception ex)
                    {
                        richTextBox1.Invoke(new Action(() => {
                            richTextBox1.AppendText($"[{DateTime.Now}] Lỗi khi tắt driver trong cleanup: {ex.Message}\n");
                        }));
                    }
                }

                // Remove from maps
                if (rowDriverMap.ContainsKey(process.RowIndex))
                {
                    rowDriverMap.Remove(process.RowIndex);
                }

                if (activeProcesses.ContainsKey(process.RowIndex))
                {
                    activeProcesses.Remove(process.RowIndex);
                }
                int displayRowIndex = FindDisplayRowIndexByRealIndex(process.RowIndex);
                if (displayRowIndex == -1)
                {
                    
                    return;
                }
                // Cập nhật UI
                //this.Invoke(new Action(() => {
                //    dataGridView1.Rows[displayRowIndex].Cells[8].Value = "Đã tắt";
                //}));

                richTextBox1.Invoke(new Action(() => {
                    richTextBox1.AppendText($"[{DateTime.Now}] Đã cleanup hoàn toàn process cho Row {process.RowIndex}\n");
                }));
            }
            catch (Exception ex)
            {
                richTextBox1.Invoke(new Action(() => {
                    richTextBox1.AppendText($"[{DateTime.Now}] Lỗi cleanup Row {process.RowIndex}: {ex.Message}\n");
                }));
            }
        }
        //        private void ProcessiCloudRequest(string userIcloud, string passIcloud, string deviceInfo, int pro5, int index)
        //        {
        //            this.chromeStartTimes[index] = DateTime.Now;
        //            Random random = new Random();
        //            string[] array = System.IO.File.ReadAllLines("UA.txt");
        //            string useragen = array[random.Next(0, array.Length)];
        //            IWebDriver chromeDriver = null;
        //            ChromeDriverService chromeDriverService = null;
        //            int displayRowIndex1 = FindDisplayRowIndexByRealIndex(index);
        //            if (displayRowIndex1 == -1)
        //            {

        //                return;
        //            }
        //            string namePhone = dataGridView1.Rows[displayRowIndex1].Cells[9].Value?.ToString()?.Trim().ToUpper() ?? "";
        //            try
        //            {
        //                //richTextBox1.AppendText($"Bắt đầu xử lý yêu cầu cho {deviceInfo}");
        //                this.Invoke((MethodInvoker)delegate
        //                {
        //                    richTextBox1.AppendText($"Bắt đầu xử lý yêu cầu cho {deviceInfo} \r\n");
        //                });
        //                chromeDriverService = ChromeDriverService.CreateDefaultService();
        //                chromeDriverService.HideCommandPromptWindow = true;

        //                // Đường dẫn tới thư mục chứa Chrome Portable
        //                string chromePortableDirectory = Path.Combine(Directory.GetCurrentDirectory(), "GoogleChromePortable", "App", "Chrome-bin");

        //                // Tạo profile riêng cho mỗi session
        //                //string sessionId = DateTime.Now.Ticks.ToString();
        //                string ProfilePath = Path.Combine(chromePortableDirectory, "User_Data_" + deviceInfo);

        //                if (!Directory.Exists(ProfilePath))
        //                {
        //                    Directory.CreateDirectory(ProfilePath);
        //                }

        //                // Tìm port available

        //                ChromeOptions chromeOptions = new ChromeOptions();
        //                chromeOptions.BinaryLocation = Path.Combine(chromePortableDirectory, "Chrome.exe");
        //                chromeOptions.AddArgument($"--user-data-dir={ProfilePath}");
        //                // Bật password manager
        //                chromeOptions.AddUserProfilePreference("credentials_enable_service", true);
        //                chromeOptions.AddUserProfilePreference("profile.password_manager_enabled", true);

        //                // Cải thiện để tránh phát hiện bot
        //                chromeOptions.AddArgument("--disable-blink-features=AutomationControlled");
        //                chromeOptions.AddExcludedArgument("enable-automation");
        //                chromeOptions.AddAdditionalOption("useAutomationExtension", false);

        //                // Thêm User-Agent thực tế
        //                chromeOptions.AddArgument("--user-agent="+useragen);
        //                //chromeOptions.AddArgument("--disable-web-security");
        //                chromeOptions.AddArgument("--disable-features=VizDisplayCompositor");
        //                // Giữ các argument cần thiết
        //                chromeOptions.AddArgument("--mute-audio");
        //                chromeOptions.AddArgument("--no-first-run");
        //                chromeOptions.AddArgument("--disable-default-apps");
        //                chromeOptions.AddArgument("--disable-popup-blocking");
        //                chromeOptions.AddArgument("--remote-allow-origins=*");
        //                chromeOptions.AddArgument("--disable-logging");
        //                chromeOptions.AddArgument("--disable-extensions");
        //                chromeOptions.AddArgument("--disable-plugins");
        //                chromeOptions.AddArgument("--disable-images"); // Tăng tốc độ load
        //                chromeOptions.AddArgument("--disable-javascript-harmony-shipping");
        //                // Performance optimization (giữ lại những cần thiết)
        //                chromeOptions.AddArgument("--disable-gpu");
        //                chromeOptions.AddArgument("--no-sandbox");
        //                chromeOptions.AddArgument("--disable-dev-shm-usage");
        //                // Thay đổi timezone ngẫu nhiên
        //                string[] timezones = { "America/New_York", "Europe/London", "Asia/Tokyo", "Australia/Sydney" };
        //                chromeOptions.AddArgument($"--timezone={timezones[random.Next(timezones.Length)]}");

        //                #region calc position for profile
        //                {
        //                    // calc size
        //                    Screen[] screens = Screen.AllScreens;
        //                    Rectangle secondScreenBounds = screens[0].Bounds;
        //                    int max_width = secondScreenBounds.Width;
        //                    int max_height = secondScreenBounds.Height;

        //                    int width = ConfigInfo.chrome_width;
        //                    int height = ConfigInfo.chrome_height;
        //                    chromeOptions.AddArgument($"--window-size={width},{height}");

        //                    // calc max position for pro5
        //                    int distance_x = ConfigInfo.chrome_distance_x;
        //                    int distance_y = ConfigInfo.chrome_distance_y;

        //                    // Tính số cột và hàng tối đa có thể hiển thị
        //                    int max_column = (max_width - width) / distance_x + 1;
        //                    int max_row = (max_height - height) / distance_y + 1;

        //                    // Đảm bảo có ít nhất 1 cột và 1 hàng
        //                    max_column = Math.Max(1, max_column);
        //                    max_row = Math.Max(1, max_row);

        //                    // Tính vị trí dựa trên pro5 (bắt đầu từ 1)
        //                    int adjustedPosition = pro5 - 1; // Chuyển về base 0 để tính toán
        //                    int column = (adjustedPosition % max_column) + 1; // Cột từ 1 đến max_column
        //                    int row = (adjustedPosition / max_column) + 1;    // Hàng từ 1 trở lên

        //                    // Nếu vượt quá màn hình, wrap lại
        //                    if (row > max_row)
        //                    {
        //                        row = ((row - 1) % max_row) + 1;
        //                    }

        //                    // Tính toán vị trí pixel
        //                    int margin_width_position = (column - 1) * distance_x;
        //                    int margin_height_position = (row - 1) * distance_y;

        //                    // Đảm bảo không vượt quá biên màn hình
        //                    margin_width_position = Math.Min(margin_width_position, max_width - width);
        //                    margin_height_position = Math.Min(margin_height_position, max_height - height);

        //                    string position = $"--window-position={margin_width_position},{margin_height_position}";
        //                    chromeOptions.AddArgument(position);

        //                }
        //                #endregion


        //                chromeDriver = new ChromeDriver(chromeDriverService, chromeOptions);
        //                IJavaScriptExecutor jsExecutor = (IJavaScriptExecutor)chromeDriver;
        //                jsExecutor.ExecuteScript(@"
        //    Object.defineProperty(navigator, 'webdriver', {get: () => undefined});
        //    Object.defineProperty(navigator, 'plugins', {get: () => [1, 2, 3, 4, 5]});
        //    Object.defineProperty(navigator, 'languages', {get: () => ['en-US', 'en']});
        //    window.chrome = { runtime: {} };
        //    Object.defineProperty(navigator, 'permissions', {
        //        query: () => Promise.resolve({ state: 'granted' })
        //    });
        //// Fake WebGL fingerprint
        //    const getParameter = WebGLRenderingContext.prototype.getParameter;
        //    WebGLRenderingContext.prototype.getParameter = function(parameter) {{
        //        if (parameter === 37445) return 'Intel Inc.';
        //        if (parameter === 37446) return 'Intel(R) HD Graphics {rand.Next(4000, 6000)}';
        //        return getParameter.call(this, parameter);
        //    }};

        //    // Fake canvas fingerprint
        //    const toDataURL = HTMLCanvasElement.prototype.toDataURL;
        //    HTMLCanvasElement.prototype.toDataURL = function() {{
        //        const context = this.getContext('2d');
        //        context.fillStyle = 'rgb({rand.Next(0, 255)}, {rand.Next(0, 255)}, {rand.Next(0, 255)})';
        //        context.fillRect(0, 0, 1, 1);
        //        return toDataURL.call(this);
        //    }};
        //");
        //                // Thêm driver vào danh sách quản lý
        //                lock (activeDrivers)
        //                {
        //                    activeDrivers.Add(chromeDriver);
        //                    rowDriverMap[index] = chromeDriver;
        //                }


        //                // Thiết lập timeout ngắn hơn
        //                chromeDriver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(30);
        //                chromeDriver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);

        //                // Mở iCloud Find My
        //                chromeDriver.Navigate().GoToUrl("https://www.icloud.com/find/");
        //                Thread.Sleep(3000);
        //                this.Invoke((MethodInvoker)delegate
        //                {
        //                    richTextBox1.AppendText($"Đang tiến hành login tài khoản icloud: {userIcloud}... \r\n");
        //                });
        //                // Click sign in button
        //                var wait = new WebDriverWait(chromeDriver, TimeSpan.FromSeconds(20));
        //                try
        //                {
        //                    var signInBtn = wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("//ui-button[contains(@class, 'sign-in-button')]")));
        //                    signInBtn.Click();
        //                    Thread.Sleep(3000);
        //                }
        //                catch (WebDriverTimeoutException)
        //                {
        //                    throw new Exception("Không tìm thấy nút Sign In");
        //                }

        //                // Switch to iframe and login
        //                try
        //                {
        //                    wait.Until(ExpectedConditions.FrameToBeAvailableAndSwitchToIt("aid-auth-widget"));
        //                }
        //                catch (WebDriverTimeoutException)
        //                {
        //                    throw new Exception("Không tìm thấy iframe đăng nhập");
        //                }

        //                // Nhập email
        //                var emailField = wait.Until(ExpectedConditions.ElementToBeClickable(By.Id("account_name_text_field")));
        //                emailField.Clear();
        //                emailField.SendKeys(userIcloud);
        //                Thread.Sleep(1000);

        //                // Click checkbox và next
        //                var nextPass = wait.Until(ExpectedConditions.ElementToBeClickable(By.ClassName("form-checkbox-indicator")));
        //                nextPass.Click();
        //                Thread.Sleep(1000);

        //                var singin = wait.Until(ExpectedConditions.ElementToBeClickable(By.Id("sign-in")));
        //                singin.Click();
        //                Thread.Sleep(2000);

        //                // Nhập password
        //                var passField = wait.Until(ExpectedConditions.ElementToBeClickable(By.Id("password_text_field")));
        //                passField.Clear();
        //                passField.SendKeys(passIcloud);
        //                Thread.Sleep(1000);

        //                // Click sign in
        //                IWebElement signInButton = chromeDriver.FindElement(By.Id("sign-in"));
        //                bool isDisabled = signInButton.GetAttribute("disabled") != null;

        //                if (isDisabled)
        //                {
        //                    IJavaScriptExecutor js = (IJavaScriptExecutor)chromeDriver;
        //                    js.ExecuteScript("arguments[0].removeAttribute('disabled');", signInButton);
        //                    js.ExecuteScript("arguments[0].click();", signInButton);
        //                }
        //                else
        //                {
        //                    signInButton.Click();
        //                }
        //                // Đợi một chút để form xử lý
        //                Thread.Sleep(3000);

        //                // Kiểm tra xem có lỗi đăng nhập không
        //                try
        //                {
        //                    // Tìm element báo lỗi
        //                    var errorElement = chromeDriver.FindElement(By.CssSelector(".form-cell-wrapper.is-error, #invalid_user_name_pwd_err_msg[aria-hidden='false'], .si-error-message"));
        //                    if (errorElement != null && errorElement.Displayed)
        //                    {
        //                        // Có lỗi đăng nhập
        //                        string errorMessage = "Mật khẩu hoặc tài khoản không đúng";
        //                        try
        //                        {
        //                            var errorText = chromeDriver.FindElement(By.Id("invalid_user_name_pwd_err_msg"));
        //                            if (!string.IsNullOrEmpty(errorText.Text))
        //                            {
        //                                errorMessage = errorText.Text;
        //                            }
        //                            TatChromeKhiLoi(index);
        //                        }
        //                        catch { }

        //                        this.Invoke((MethodInvoker)delegate
        //                        {
        //                            richTextBox1.AppendText($"[{DateTime.Now:HH:mm:ss}] ĐĂNG NHẬP THẤT BẠI - {deviceInfo}: {errorMessage}\r\n");
        //                            richTextBox1.ScrollToCaret();
        //                        });

        //                        throw new Exception($"Đăng nhập thất bại cho tài khoản {userIcloud}: {errorMessage}");
        //                    }
        //                }
        //                catch (NoSuchElementException)
        //                {
        //                    // Không tìm thấy element lỗi có nghĩa là đăng nhập thành công
        //                }

        //                // Kiểm tra xem có chuyển sang trang chính không (đăng nhập thành công)
        //                try
        //                {
        //                    // Đợi một chút để trang load
        //                    Thread.Sleep(2000);

        //                    // Kiểm tra xem có còn trong iframe đăng nhập không
        //                    var currentUrl = chromeDriver.Url;
        //                    if (currentUrl.Contains("idmsa.apple.com") || currentUrl.Contains("signin"))
        //                    {
        //                        // Vẫn còn trong trang đăng nhập, có thể là lỗi
        //                        throw new Exception("Vẫn còn trong trang đăng nhập, có thể đăng nhập thất bại");
        //                    }
        //                }
        //                catch (Exception ex)
        //                {
        //                    // Ghi log lỗi nhưng vẫn thử tiếp tục
        //                    Console.WriteLine($"Cảnh báo: {ex.Message}");
        //                }
        //                // Switch back to main content
        //                chromeDriver.SwitchTo().DefaultContent();
        //                this.Invoke((MethodInvoker)delegate
        //                {
        //                    richTextBox1.AppendText($"Đã đăng nhập thành công tài khoản icloud {userIcloud}, đợi trang Find My load...\r\n");
        //                });
        //                // Đợi trang Find My load
        //                Thread.Sleep(10000);

        //                // Tìm và switch vào iframe Find My
        //                wait = new WebDriverWait(chromeDriver, TimeSpan.FromSeconds(20));
        //                IWebElement iframe = null;

        //                string[] iframeSelectors = {
        //            "iframe.child-application",
        //            "iframe[src*='find']",
        //            "iframe[title*='Find']",
        //            ".child-application"
        //        };

        //                foreach (string selector in iframeSelectors)
        //                {
        //                    try
        //                    {
        //                        iframe = wait.Until(ExpectedConditions.ElementToBeClickable(By.CssSelector(selector)));
        //                        if (iframe != null)
        //                        {

        //                            //richTextBox1.AppendText($"Tìm thấy iframe với selector: {selector}");
        //                            break;
        //                        }
        //                    }
        //                    catch (WebDriverTimeoutException)
        //                    {
        //                        continue;
        //                    }
        //                }

        //                if (iframe == null)
        //                {
        //                    TatChromeKhiLoi(index);
        //                    throw new Exception("Không thể tìm thấy iframe Find My");
        //                }

        //                chromeDriver.SwitchTo().Frame(iframe);
        //                //richTextBox1.AppendText("Đã switch vào iframe Find My");

        //                // Đợi device list load
        //                Thread.Sleep(5000);

        //                try
        //                {
        //                    wait.Until(driver => driver.FindElements(By.CssSelector(".fmip-device-list-item")).Count > 0);
        //                }
        //                catch (WebDriverTimeoutException)
        //                {
        //                    throw new Exception("Không thể load danh sách thiết bị");
        //                }

        //                // Tìm và click vào iPhone
        //                var deviceElements = chromeDriver.FindElements(By.CssSelector(".fmip-device-list-item"));
        //                this.Invoke((MethodInvoker)delegate
        //                {
        //                    richTextBox1.AppendText($"Tìm thấy {deviceElements.Count} thiết bị trên tài khoản {userIcloud}\r\n");
        //                });
        //                bool iPhoneFound = false;
        //                foreach (var device in deviceElements)
        //                {
        //                    try
        //                    {
        //                        var nameElement = device.FindElement(By.CssSelector("[data-testid='show-device-name'], .device-name, .name"));

        //                        var phoneName = GetIPhoneName(deviceInfo);

        //                        var deviceName = nameElement.Text.Trim().ToUpper();
        //                        this.Invoke((MethodInvoker)delegate
        //                        {
        //                            richTextBox1.AppendText($"Tìm thấy thiết bị: {deviceName} \r\n");
        //                        });
        //                        string number = Regex.Match(deviceInfo, @"\d+").Value;
        //                        string parts = Regex.Match(deviceName, @"\d+").Value;
        //                        bool isExactMatch3 = (parts.Length > 0 &&
        //                                            string.Equals(parts, number, StringComparison.OrdinalIgnoreCase));

        //                        // Kiểm tra match theo số hoặc tên
        //                        bool shouldProcessDevice = isExactMatch3 || (namePhone == deviceName);

        //                        if (shouldProcessDevice)
        //                        {
        //                            int displayRowIndex = FindDisplayRowIndexByRealIndex(index);
        //                            if (displayRowIndex == -1)
        //                            {

        //                                return;
        //                            }
        //                            this.Invoke((MethodInvoker)delegate
        //                            {
        //                                richTextBox1.AppendText("Tìm thấy iPhone, đang tiến hành kiểm tra on hay off... \r\n");
        //                            });
        //                            var imageElement = device.FindElement(By.CssSelector("img.image"));
        //                            var imageSrc = imageElement.GetAttribute("src");

        //                            // Kiểm tra trạng thái dựa vào tên file ảnh
        //                            bool isOnline = imageSrc.Contains("online-sourcelist.png");
        //                            bool isOffline = imageSrc.Contains("offline-sourcelist.png");

        //                            string deviceStatus = "";
        //                            if (isOnline || isOffline)
        //                            {
        //                                deviceStatus = "Quét...";
        //                                this.Invoke((MethodInvoker)delegate
        //                                {
        //                                    richTextBox1.AppendText("Thiết bị đang được tiến hành click... \r\n");
        //                                });
        //                                dataGridView1.Rows[displayRowIndex].Cells[8].Value = deviceStatus;
        //                                nameElement.Click();
        //                                iPhoneFound = true;
        //                                Thread.Sleep(3000);

        //                                // Nếu cần kiểm tra model thiết bị, làm sau khi click
        //                                bool shouldPlaySound = true;
        //                                if (namePhone == deviceName && !string.IsNullOrEmpty(phoneName))
        //                                {
        //                                    try
        //                                    {
        //                                        // Lấy thông tin model sau khi đã click vào thiết bị
        //                                        var typeElement = wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector(".device-details-view .subTitle-name")));
        //                                        var typeIphone = typeElement.Text.Trim().ToUpper();

        //                                        this.Invoke((MethodInvoker)delegate
        //                                        {
        //                                            richTextBox1.AppendText($"Kiểm tra loại thiết bị: {typeIphone} vs {phoneName.ToUpper()}\r\n");
        //                                        });

        //                                        // Chỉ phát âm thanh nếu model khớp
        //                                        shouldPlaySound = (phoneName.ToUpper() == typeIphone);

        //                                        if (!shouldPlaySound)
        //                                        {
        //                                            this.Invoke((MethodInvoker)delegate
        //                                            {
        //                                                richTextBox1.AppendText("Model thiết bị không khớp, quay lại danh sách...\r\n");
        //                                            });
        //                                            // Quay lại danh sách thiết bị
        //                                            var allDevicesButton = chromeDriver.FindElement(By.CssSelector(".close-button-x"));
        //                                            allDevicesButton.Click();
        //                                            Thread.Sleep(3000);
        //                                            continue; // Tiếp tục với thiết bị tiếp theo
        //                                        }
        //                                    }
        //                                    catch (Exception ex)
        //                                    {
        //                                        this.Invoke((MethodInvoker)delegate
        //                                        {
        //                                            richTextBox1.AppendText($"Không thể lấy thông tin model: {ex.Message}\r\n");
        //                                        });
        //                                        // Nếu không lấy được model, vẫn tiếp tục phát âm thanh
        //                                    }
        //                                }

        //                                if (shouldPlaySound)
        //                                {
        //                                    // Tìm và click nút Play Sound
        //                                    string[] playSoundSelectors = {
        //                        ".fm-test-playsndbtn",
        //                        ".play-sound-button",
        //                        "ui-button[title*='Play sound']",
        //                        "[class*='play-sound']"
        //                    };

        //                                    IWebElement playSoundButton = null;
        //                                    foreach (string selector in playSoundSelectors)
        //                                    {
        //                                        try
        //                                        {
        //                                            playSoundButton = wait.Until(ExpectedConditions.ElementToBeClickable(By.CssSelector(selector)));
        //                                            if (playSoundButton != null && playSoundButton.Displayed)
        //                                            {
        //                                                break;
        //                                            }
        //                                        }
        //                                        catch (WebDriverTimeoutException)
        //                                        {
        //                                            continue;
        //                                        }
        //                                    }

        //                                    if (playSoundButton != null)
        //                                    {
        //                                        // Hiển thị thông báo trên UI thread
        //                                        this.Invoke((MethodInvoker)delegate
        //                                        {
        //                                            richTextBox1.AppendText($"Đã tìm thấy iPhone của {deviceInfo}. Bắt đầu phát âm thanh...\r\n");
        //                                        });

        //                                        // Click Play Sound 12000 lần
        //                                        for (int i = 0; i < 12000; i++)
        //                                        {
        //                                            try
        //                                            {
        //                                                playSoundButton.Click();
        //                                                Thread.Sleep(random.Next(4000,9000)); // Đợi 7 giây giữa các lần click
        //                                            }
        //                                            catch (Exception clickEx)
        //                                            {
        //                                                this.Invoke((MethodInvoker)delegate
        //                                                {
        //                                                    richTextBox1.AppendText($"Lỗi khi click lần {i + 1} cho thiết bị {deviceInfo}: {clickEx.Message}\r\n");
        //                                                });
        //                                                break;
        //                                            }
        //                                        }

        //                                        this.Invoke((MethodInvoker)delegate
        //                                        {
        //                                            dataGridView1.Rows[displayRowIndex].Cells[8].Value = "Đã Tắt";
        //                                        });
        //                                    }
        //                                    else
        //                                    {
        //                                        this.Invoke((MethodInvoker)delegate
        //                                        {
        //                                            dataGridView1.Rows[displayRowIndex].Cells[8].Value = "Không tìm thấy nút Play Sound";
        //                                        });
        //                                    }
        //                                }
        //                                break; // Thoát khỏi vòng lặp sau khi xử lý xong
        //                            }
        //                            else
        //                            {
        //                                deviceStatus = "Unknown";
        //                                dataGridView1.Rows[displayRowIndex].Cells[8].Value = deviceStatus;
        //                            }
        //                        }
        //                    }
        //                    catch (Exception deviceEx)
        //                    {
        //                        this.Invoke((MethodInvoker)delegate
        //                        {
        //                            richTextBox1.AppendText($"Lỗi khi xử lý thiết bị: {deviceEx.Message}\r\n");
        //                        });
        //                        break;
        //                    }
        //                }

        //                if (!iPhoneFound)
        //                {
        //                    this.Invoke((MethodInvoker)delegate
        //                    {
        //                        richTextBox1.AppendText("Không tìm thấy thiết bị iPhone trong danh sách \r\n");
        //                    });
        //                    this.Invoke((MethodInvoker)delegate
        //                    {
        //                        richTextBox1.AppendText("Không tìm thấy thiết bị iPhone trong danh sách");
        //                        //MessageBox.Show("Không tìm thấy thiết bị iPhone trong danh sách", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
        //                    });
        //                }
        //            }
        //            catch (Exception ex)
        //            {
        //                this.Invoke((MethodInvoker)delegate
        //                {
        //                    richTextBox1.AppendText($"Lỗi: {ex.Message}\r\n");
        //                });
        //                KillChromeProcessesWithUserDataDir(deviceInfo);
        //            }
        //            finally
        //            {
        //                // Đóng browser và cleanup
        //                try
        //                {
        //                    if (chromeDriver != null)
        //                    {
        //                        lock (activeDrivers)
        //                        {
        //                            activeDrivers.Remove(chromeDriver);
        //                            rowDriverMap.Remove(index); // Remove from map
        //                        }
        //                        chromeDriver.Quit();
        //                        this.Invoke((MethodInvoker)delegate
        //                        {
        //                            richTextBox1.AppendText($"Đã đóng Chrome instance cho thiết bị {deviceInfo}\r\n");
        //                        });
        //                    }

        //                    if (chromeDriverService != null)
        //                    {
        //                        chromeDriverService.Dispose();
        //                    }
        //                }
        //                catch (Exception cleanupEx)
        //                {
        //                    this.Invoke((MethodInvoker)delegate
        //                    {
        //                        richTextBox1.AppendText($"Lỗi khi cleanup: {cleanupEx.Message} \r\n");
        //                    });
        //                }
        //            }
        //        }

        private void ProcessiCloudRequest(string userIcloud, string passIcloud, string deviceInfo, int pro5, int index)
        {
            this.chromeStartTimes[index] = DateTime.Now;
            Random random = new Random();
            string[] array = System.IO.File.ReadAllLines("UA.txt");
            string str = array[random.Next(0, array.Length)];
            IWebDriver webDriver = null;
            ChromeDriverService chromeDriverService = null;
            int num = this.FindDisplayRowIndexByRealIndex(index);
            bool flag = num == -1;
            if (!flag)
            {
                object value = this.dataGridView1.Rows[num].Cells[9].Value;
                string text;
                if (value == null)
                {
                    text = null;
                }
                else
                {
                    string text2 = value.ToString();
                    text = ((text2 != null) ? text2.Trim().ToUpper() : null);
                }
                string a = text ?? "";
                try
                {
                    base.Invoke(new MethodInvoker(delegate ()
                    {
                        this.richTextBox1.AppendText("Bắt đầu xử lý yêu cầu cho " + deviceInfo + " \r\n");
                    }));
                    chromeDriverService = ChromeDriverService.CreateDefaultService();
                    chromeDriverService.HideCommandPromptWindow = true;
                    string path = Path.Combine(Directory.GetCurrentDirectory(), "GoogleChromePortable", "App", "Chrome-bin");
                    string text3 = Path.Combine(path, "User_Data_" + deviceInfo);
                    bool flag2 = !Directory.Exists(text3);
                    if (flag2)
                    {
                        Directory.CreateDirectory(text3);
                    }
                    ChromeOptions chromeOptions = new ChromeOptions();
                    Dictionary<string, object> dictionary = new Dictionary<string, object>();
                    dictionary.Add("profile.default_content_setting_values.geolocation", 1);
                    dictionary.Add("profile.default_content_settings.popups", 0);
                    chromeOptions.AddUserProfilePreference("profile.default_content_setting_values.geolocation", 1);
                    chromeOptions.AddArgument("--disable-geolocation");
                    chromeOptions.BinaryLocation = Path.Combine(path, "Chrome.exe");
                    chromeOptions.AddArgument("--user-data-dir=" + text3);
                    chromeOptions.AddUserProfilePreference("credentials_enable_service", true);
                    chromeOptions.AddUserProfilePreference("profile.password_manager_enabled", true);
                    chromeOptions.AddArgument("--disable-blink-features=AutomationControlled");
                    chromeOptions.AddExcludedArgument("enable-automation");
                    chromeOptions.AddAdditionalOption("useAutomationExtension", true);
                    chromeOptions.AddArgument("--user-agent=" + str);
                    chromeOptions.AddArgument("--disable-features=VizDisplayCompositor");
                    chromeOptions.AddArgument("--mute-audio");
                    chromeOptions.AddArgument("--no-first-run");
                    chromeOptions.AddArgument("--no-default-browser-check");
                    chromeOptions.AddArgument("--disable-default-apps");
                    chromeOptions.AddArgument("--disable-popup-blocking");
                    chromeOptions.AddArgument("--remote-allow-origins=*");
                    chromeOptions.AddArgument("--disable-logging");
                    chromeOptions.AddArgument("--enable-extensions");
                    chromeOptions.AddArgument("--disable-plugins");
                    chromeOptions.AddArgument("--disable-images");
                    chromeOptions.AddArgument("--disable-javascript-harmony-shipping");
                    chromeOptions.AddArgument("--disable-gpu");
                    chromeOptions.AddArgument("--no-sandbox");
                    chromeOptions.AddArgument("--disable-dev-shm-usage");
                    string[] array2 = new string[]
                    {
                        "America/New_York",
                        "Europe/London",
                        "Asia/Tokyo",
                        "Australia/Sydney"
                    };
                    chromeOptions.AddArgument("--timezone=" + array2[random.Next(array2.Length)]);
                    Screen[] allScreens = Screen.AllScreens;
                    Rectangle bounds = allScreens[0].Bounds;
                    int width = bounds.Width;
                    int height = bounds.Height;
                    int chrome_width = ConfigInfo.chrome_width;
                    int chrome_height = ConfigInfo.chrome_height;
                    chromeOptions.AddArgument(string.Format("--window-size={0},{1}", chrome_width, chrome_height));
                    int chrome_distance_x = ConfigInfo.chrome_distance_x;
                    int chrome_distance_y = ConfigInfo.chrome_distance_y;
                    int num2 = (width - chrome_width) / chrome_distance_x + 1;
                    int num3 = (height - chrome_height) / chrome_distance_y + 1;
                    num2 = Math.Max(1, num2);
                    num3 = Math.Max(1, num3);
                    int num4 = pro5 - 1;
                    int num5 = num4 % num2 + 1;
                    int num6 = num4 / num2 + 1;
                    bool flag3 = num6 > num3;
                    if (flag3)
                    {
                        num6 = (num6 - 1) % num3 + 1;
                    }
                    int num7 = (num5 - 1) * chrome_distance_x;
                    int num8 = (num6 - 1) * chrome_distance_y;
                    num7 = Math.Min(num7, width - chrome_width);
                    num8 = Math.Min(num8, height - chrome_height);
                    string text4 = string.Format("--window-position={0},{1}", num7, num8);
                    chromeOptions.AddArgument(text4);
                    webDriver = new ChromeDriver(chromeDriverService, chromeOptions);
                    IJavaScriptExecutor javaScriptExecutor = (IJavaScriptExecutor)webDriver;
                    javaScriptExecutor.ExecuteScript("// Comprehensive Anti-Detection Script for iCloud Automation\r\n\r\n// 1. Hide webdriver property\r\nObject.defineProperty(navigator, 'webdriver', {\r\n    get: () => undefined,\r\n    configurable: true\r\n});\r\n\r\n// 2. Realistic plugins array\r\nObject.defineProperty(navigator, 'plugins', {\r\n    get: () => [\r\n        {\r\n            name: \"Chrome PDF Plugin\",\r\n            filename: \"internal-pdf-viewer\",\r\n            description: \"Portable Document Format\",\r\n            length: 1\r\n        },\r\n        {\r\n            name: \"Chrome PDF Viewer\",\r\n            filename: \"mhjfbmdgcfjbbpaeojofohoefgiehjai\", \r\n            description: \"Portable Document Format\",\r\n            length: 1\r\n        },\r\n        {\r\n            name: \"Native Client\",\r\n            filename: \"internal-nacl-plugin\",\r\n            description: \"Native Client\",\r\n            length: 2\r\n        }\r\n    ]\r\n});\r\n\r\n// 3. Set realistic languages\r\nObject.defineProperty(navigator, 'languages', {\r\n    get: () => ['en-US', 'en', 'vi-VN', 'vi']\r\n});\r\n\r\n// 4. Mock Chrome runtime\r\nwindow.chrome = {\r\n    runtime: {\r\n        onConnect: undefined,\r\n        onMessage: undefined,\r\n        connect: () => {},\r\n        sendMessage: () => {}\r\n    },\r\n    csi: () => {},\r\n    loadTimes: () => {}\r\n};\r\n\r\n// 5. Mock permissions API\r\nObject.defineProperty(navigator, 'permissions', {\r\n    get: () => ({\r\n        query: () => Promise.resolve({ state: 'granted' })\r\n    })\r\n});\r\n\r\n// 6. Enhanced WebGL fingerprint spoofing\r\nconst getParameter = WebGLRenderingContext.prototype.getParameter;\r\nWebGLRenderingContext.prototype.getParameter = function(parameter) {\r\n    // UNMASKED_VENDOR_WEBGL\r\n    if (parameter === 37445) {\r\n        const vendors = ['Intel Inc.', 'NVIDIA Corporation', 'AMD'];\r\n        return vendors[Math.floor(Math.random() * vendors.length)];\r\n    }\r\n    // UNMASKED_RENDERER_WEBGL  \r\n    if (parameter === 37446) {\r\n        const renderers = [\r\n            'Intel(R) HD Graphics 620',\r\n            'NVIDIA GeForce GTX 1060',\r\n            'AMD Radeon RX 580'\r\n        ];\r\n        return renderers[Math.floor(Math.random() * renderers.length)];\r\n    }\r\n    return getParameter.call(this, parameter);\r\n};\r\n\r\n// 7. Enhanced canvas fingerprint protection\r\nconst toDataURL = HTMLCanvasElement.prototype.toDataURL;\r\nHTMLCanvasElement.prototype.toDataURL = function() {\r\n    const context = this.getContext('2d');\r\n    const imageData = context.getImageData(0, 0, this.width, this.height);\r\n    \r\n    // Add subtle noise to avoid detection\r\n    for(let i = 0; i < imageData.data.length; i += 4) {\r\n        imageData.data[i] += Math.floor(Math.random() * 2) - 1;     // Red\r\n        imageData.data[i + 1] += Math.floor(Math.random() * 2) - 1; // Green  \r\n        imageData.data[i + 2] += Math.floor(Math.random() * 2) - 1; // Blue\r\n    }\r\n    context.putImageData(imageData, 0, 0);\r\n    return toDataURL.call(this);\r\n};\r\n\r\n// 8. Spoof screen resolution variations\r\nObject.defineProperty(screen, 'width', {\r\n    get: () => 1920 + Math.floor(Math.random() * 100)\r\n});\r\nObject.defineProperty(screen, 'height', {\r\n    get: () => 1080 + Math.floor(Math.random() * 100)\r\n});\r\n\r\n// 9. Mock hardware concurrency\r\nObject.defineProperty(navigator, 'hardwareConcurrency', {\r\n    get: () => 4 + Math.floor(Math.random() * 4) // 4-8 cores\r\n});\r\n\r\n// 10. Spoof memory info\r\nObject.defineProperty(navigator, 'deviceMemory', {\r\n    get: () => [4, 8, 16][Math.floor(Math.random() * 3)]\r\n});\r\n\r\n// 11. Mock battery API\r\nObject.defineProperty(navigator, 'getBattery', {\r\n    get: () => () => Promise.resolve({\r\n        charging: Math.random() > 0.5,\r\n        chargingTime: Math.random() * 1000,\r\n        dischargingTime: Math.random() * 10000,\r\n        level: Math.random()\r\n    })\r\n});\r\n\r\n// 12. Spoof timezone\r\nObject.defineProperty(Intl.DateTimeFormat.prototype, 'resolvedOptions', {\r\n    get: () => function() {\r\n        return {\r\n            locale: 'en-US',\r\n            timeZone: 'Asia/Ho_Chi_Minh',\r\n            calendar: 'gregory',\r\n            numberingSystem: 'latn'\r\n        };\r\n    }\r\n});\r\n\r\n// 13. Mock connection info\r\nObject.defineProperty(navigator, 'connection', {\r\n    get: () => ({\r\n        effectiveType: '4g',\r\n        rtt: Math.floor(Math.random() * 100) + 50,\r\n        downlink: Math.random() * 10 + 5,\r\n        saveData: false\r\n    })\r\n});\r\n\r\n// 14. Prevent automation detection through timing\r\nconst originalSetTimeout = window.setTimeout;\r\nwindow.setTimeout = function(callback, delay) {\r\n    // Add slight randomness to timing\r\n    const randomDelay = delay + Math.floor(Math.random() * 10) - 5;\r\n    return originalSetTimeout.call(this, callback, Math.max(0, randomDelay));\r\n};\r\n\r\n// 15. Mock media devices\r\nObject.defineProperty(navigator, 'mediaDevices', {\r\n    get: () => ({\r\n        enumerateDevices: () => Promise.resolve([\r\n            {\r\n                deviceId: 'default',\r\n                kind: 'audioinput',\r\n                label: 'Default - Microphone',\r\n                groupId: 'default'\r\n            },\r\n            {\r\n                deviceId: 'default',\r\n                kind: 'audiooutput', \r\n                label: 'Default - Speaker',\r\n                groupId: 'default'\r\n            }\r\n        ])\r\n    })\r\n});\r\n\r\n// 16. Remove automation indicators\r\ndelete navigator.__webdriver_evaluate;\r\ndelete navigator.__webdriver_script_function;\r\ndelete navigator.__webdriver_script_func;\r\ndelete navigator.__webdriver_script_fn;\r\ndelete navigator.__fxdriver_evaluate;\r\ndelete navigator.__fxdriver_unwrapped;\r\ndelete navigator.__driver_unwrapped;\r\ndelete navigator.__webdriver_unwrapped;\r\ndelete navigator.__driver_evaluate;\r\ndelete navigator.__selenium_evaluate;\r\ndelete navigator.__selenium_unwrapped;\r\ndelete navigator.__webdriver_script_fn;\r\n\r\nconsole.log('Anti-detection script loaded successfully');", Array.Empty<object>());
                    List<IWebDriver> obj = Form1.activeDrivers;
                    lock (obj)
                    {
                        Form1.activeDrivers.Add(webDriver);
                        this.rowDriverMap[index] = webDriver;
                    }
                    webDriver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(30.0);
                    webDriver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10.0);
                    webDriver.Navigate().GoToUrl("https://www.icloud.com/find/");
                    Thread.Sleep(random.Next(3000, 5000));
                    base.Invoke(new MethodInvoker(delegate ()
                    {
                        this.richTextBox1.AppendText("Đang tiến hành login tài khoản icloud: " + userIcloud + "... \r\n");
                    }));
                    WebDriverWait webDriverWait = new WebDriverWait(webDriver, TimeSpan.FromSeconds(20.0));
                    try
                    {
                        IWebElement webElement = webDriverWait.Until<IWebElement>(ExpectedConditions.ElementToBeClickable(By.XPath("//ui-button[contains(@class, 'sign-in-button')]")));
                        webElement.Click();
                        Thread.Sleep(random.Next(3000, 5000));
                    }
                    catch (WebDriverTimeoutException)
                    {
                        throw new Exception("Không tìm thấy nút Sign In");
                    }
                    try
                    {
                        webDriverWait.Until<IWebDriver>(ExpectedConditions.FrameToBeAvailableAndSwitchToIt("aid-auth-widget"));
                    }
                    catch (WebDriverTimeoutException)
                    {
                        throw new Exception("Không tìm thấy iframe đăng nhập");
                    }
                    IWebElement webElement2 = webDriverWait.Until<IWebElement>(ExpectedConditions.ElementToBeClickable(By.Id("account_name_text_field")));
                    webElement2.Clear();
                    webElement2.SendKeys(userIcloud);
                    Thread.Sleep(random.Next(1000, 3000));
                    IWebElement webElement3 = webDriverWait.Until<IWebElement>(ExpectedConditions.ElementToBeClickable(By.ClassName("form-checkbox-indicator")));
                    webElement3.Click();
                    Thread.Sleep(random.Next(1000, 3000));
                    IWebElement webElement4 = webDriverWait.Until<IWebElement>(ExpectedConditions.ElementToBeClickable(By.Id("sign-in")));
                    webElement4.Click();
                    Thread.Sleep(random.Next(2000, 4000));
                    IWebElement webElement5 = webDriverWait.Until<IWebElement>(ExpectedConditions.ElementToBeClickable(By.Id("password_text_field")));
                    webElement5.Clear();
                    webElement5.SendKeys(passIcloud);
                    Thread.Sleep(random.Next(1000, 3000));
                    IWebElement webElement6 = webDriver.FindElement(By.Id("sign-in"));
                    bool flag5 = webElement6.GetAttribute("disabled") != null;
                    bool flag6 = flag5;
                    if (flag6)
                    {
                        IJavaScriptExecutor javaScriptExecutor2 = (IJavaScriptExecutor)webDriver;
                        javaScriptExecutor2.ExecuteScript("arguments[0].removeAttribute('disabled');", new object[]
                        {
                            webElement6
                        });
                        javaScriptExecutor2.ExecuteScript("arguments[0].click();", new object[]
                        {
                            webElement6
                        });
                    }
                    else
                    {
                        webElement6.Click();
                    }
                    Thread.Sleep(random.Next(3000, 5000));
                    try
                    {
                        IWebElement webElement7 = webDriver.FindElement(By.CssSelector(".form-cell-wrapper.is-error, #invalid_user_name_pwd_err_msg[aria-hidden='false'], .si-error-message"));
                        bool flag7 = webElement7 != null && webElement7.Displayed;
                        if (flag7)
                        {
                            string errorMessage = "Mật khẩu hoặc tài khoản không đúng";
                            try
                            {
                                IWebElement webElement8 = webDriver.FindElement(By.Id("invalid_user_name_pwd_err_msg"));
                                bool flag8 = !string.IsNullOrEmpty(webElement8.Text);
                                if (flag8)
                                {
                                    errorMessage = webElement8.Text;
                                }
                                this.TatChromeKhiLoi(index);
                            }
                            catch
                            {
                            }
                            base.Invoke(new MethodInvoker(delegate ()
                            {
                                this.richTextBox1.AppendText(string.Format("[{0:HH:mm:ss}] ĐĂNG NHẬP THẤT BẠI - {1}: {2}\r\n", DateTime.Now, deviceInfo, errorMessage));
                                this.richTextBox1.ScrollToCaret();
                            }));
                            throw new Exception("Đăng nhập thất bại cho tài khoản " + userIcloud + ": " + errorMessage);
                        }
                    }
                    catch (NoSuchElementException)
                    {
                    }
                    try
                    {
                        Thread.Sleep(random.Next(2000, 4000));
                        string url = webDriver.Url;
                        bool flag9 = url.Contains("idmsa.apple.com") || url.Contains("signin");
                        if (flag9)
                        {
                            throw new Exception("Vẫn còn trong trang đăng nhập, có thể đăng nhập thất bại");
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Cảnh báo: " + ex.Message);
                    }
                    webDriver.SwitchTo().DefaultContent();
                    base.Invoke(new MethodInvoker(delegate ()
                    {
                        this.richTextBox1.AppendText("Đã đăng nhập thành công tài khoản icloud " + userIcloud + ", đợi trang Find My load...\r\n");
                    }));
                    Thread.Sleep(random.Next(7000, 9000));
                    webDriverWait = new WebDriverWait(webDriver, TimeSpan.FromSeconds(20.0));
                    IWebElement webElement9 = null;
                    string[] array3 = new string[]
                    {
                        "iframe.child-application",
                        "iframe[src*='find']",
                        "iframe[title*='Find']",
                        ".child-application"
                    };
                    string[] array4 = array3;
                    int k = 0;
                    while (k < array4.Length)
                    {
                        string text5 = array4[k];
                        try
                        {
                            webElement9 = webDriverWait.Until<IWebElement>(ExpectedConditions.ElementToBeClickable(By.CssSelector(text5)));
                            bool flag10 = webElement9 != null;
                            if (flag10)
                            {
                                break;
                            }
                        }
                        catch (WebDriverTimeoutException)
                        {
                        }
                    IL_933:
                        k++;
                        continue;
                        goto IL_933;
                    }
                    bool flag11 = webElement9 == null;
                    if (flag11)
                    {
                        this.TatChromeKhiLoi(index);
                        throw new Exception("Không thể tìm thấy iframe Find My");
                    }
                    webDriver.SwitchTo().Frame(webElement9);
                    Thread.Sleep(random.Next(5000, 7000));
                    try
                    {
                        webDriverWait.Until<bool>((IWebDriver driver) => driver.FindElements(By.CssSelector(".fmip-device-list-item")).Count > 0);
                    }
                    catch (WebDriverTimeoutException)
                    {
                        throw new Exception("Không thể load danh sách thiết bị");
                    }
                    ReadOnlyCollection<IWebElement> deviceElements = webDriver.FindElements(By.CssSelector(".fmip-device-list-item"));
                    base.Invoke(new MethodInvoker(delegate ()
                    {
                        this.richTextBox1.AppendText(string.Format("Tìm thấy {0} thiết bị trên tài khoản {1}\r\n", deviceElements.Count, userIcloud));
                    }));
                    bool flag12 = false;
                    foreach (IWebElement webElement10 in deviceElements)
                    {
                        try
                        {
                            IWebElement webElement11 = webElement10.FindElement(By.CssSelector("[data-testid='show-device-name'], .device-name, .name"));
                            string phoneName = Form1.GetIPhoneName(deviceInfo);
                            string text6 = webElement11.Text.Trim().ToUpper();
                            string value2 = Regex.Match(deviceInfo, "\\d+").Value;
                            string value3 = Regex.Match(text6, "\\d+").Value;
                            bool flag13 = (value3.Length > 0 && string.Equals(value3, value2, StringComparison.OrdinalIgnoreCase)) || a == text6;
                            bool flag14 = flag13;
                            if (flag14)
                            {
                                int displayRowIndex = this.FindDisplayRowIndexByRealIndex(index);
                                bool flag15 = displayRowIndex == -1;
                                if (flag15)
                                {
                                    return;
                                }
                                base.Invoke(new MethodInvoker(delegate ()
                                {
                                    this.richTextBox1.AppendText("Tìm thấy iPhone, đang tiến hành kiểm tra on hay off... \r\n");
                                }));
                                IWebElement webElement12 = webElement10.FindElement(By.CssSelector("img.image"));
                                string attribute = webElement12.GetAttribute("src");
                                bool flag16 = attribute.Contains("online-sourcelist.png");
                                bool flag17 = attribute.Contains("offline-sourcelist.png");
                                bool flag18 = flag16;
                                string str2;
                                if (flag18)
                                {
                                    str2 = "online";
                                }
                                else
                                {
                                    str2 = "offline";
                                }
                                bool flag19 = flag16 || flag17;
                                string value4;
                                if (flag19)
                                {
                                    value4 = str2 + " Quét...";
                                    base.Invoke(new MethodInvoker(delegate ()
                                    {
                                        this.richTextBox1.AppendText("Thiết bị đang được tiến hành click... \r\n");
                                    }));
                                    this.dataGridView1.Rows[displayRowIndex].Cells[8].Value = value4;
                                    webElement11.Click();
                                    flag12 = true;
                                    Thread.Sleep(random.Next(3000, 5000));
                                    bool flag20 = true;
                                    bool flag21 = a == text6 && !string.IsNullOrEmpty(phoneName);
                                    if (flag21)
                                    {
                                        try
                                        {
                                            IWebElement webElement13 = webDriverWait.Until<IWebElement>(ExpectedConditions.ElementIsVisible(By.CssSelector(".device-details-view .subTitle-name")));
                                            string typeIphone = webElement13.Text.Trim().ToUpper();
                                            base.Invoke(new MethodInvoker(delegate ()
                                            {
                                                this.richTextBox1.AppendText(string.Concat(new string[]
                                                {
                                                    "Kiểm tra loại thiết bị: ",
                                                    typeIphone,
                                                    " vs ",
                                                    phoneName.ToUpper(),
                                                    "\r\n"
                                                }));
                                            }));
                                            flag20 = (phoneName.ToUpper() == typeIphone);
                                            bool flag22 = !flag20;
                                            if (flag22)
                                            {
                                                base.Invoke(new MethodInvoker(delegate ()
                                                {
                                                    this.richTextBox1.AppendText("Model thiết bị không khớp, quay lại danh sách...\r\n");
                                                }));
                                                IWebElement webElement14 = webDriver.FindElement(By.CssSelector(".close-button-x"));
                                                webElement14.Click();
                                                Thread.Sleep(random.Next(3000, 5000));
                                                continue;
                                            }
                                        }
                                        catch (Exception ex2)
                                        {
                                            Exception ex3 = ex2;
                                            Exception ex = ex3;
                                            base.Invoke(new MethodInvoker(delegate ()
                                            {
                                                this.richTextBox1.AppendText("Không thể lấy thông tin model: " + ex.Message + "\r\n");
                                            }));
                                        }
                                    }
                                    bool flag23 = flag20;
                                    if (flag23)
                                    {
                                        string[] array5 = new string[]
                                        {
                                            ".fm-test-playsndbtn",
                                            ".play-sound-button",
                                            "ui-button[title*='Play sound']",
                                            "[class*='play-sound']"
                                        };
                                        IWebElement webElement15 = null;
                                        string[] array6 = array5;
                                        int j = 0;
                                        while (j < array6.Length)
                                        {
                                            string text7 = array6[j];
                                            try
                                            {
                                                webElement15 = webDriverWait.Until<IWebElement>(ExpectedConditions.ElementToBeClickable(By.CssSelector(text7)));
                                                bool flag24 = webElement15 != null && webElement15.Displayed;
                                                if (flag24)
                                                {
                                                    break;
                                                }
                                            }
                                            catch (WebDriverTimeoutException)
                                            {
                                            }
                                        IL_E88:
                                            j++;
                                            continue;
                                            goto IL_E88;
                                        }
                                        bool flag25 = webElement15 != null;
                                        if (flag25)
                                        {
                                            
                                            base.Invoke(new MethodInvoker(delegate ()
                                            {
                                                this.richTextBox1.AppendText("Đã tìm thấy iPhone của " + deviceInfo + ". Bắt đầu phát âm thanh...\r\n");
                                            }));
                                            int i2;
                                            int i;
                                            for (i = 0; i < 12000; i = i2 + 1)
                                            {
                                                try
                                                {
                                                    webElement15.Click();
                                                    Thread.Sleep(random.Next(4000, 9000));
                                                }
                                                catch (Exception ex4)
                                                {
                                                    Exception ex3 = ex4;
                                                    Exception clickEx = ex3;
                                                    base.Invoke(new MethodInvoker(delegate ()
                                                    {
                                                        this.richTextBox1.AppendText(string.Format("Lỗi khi click lần {0} cho thiết bị {1}: {2}\r\n", i + 1, deviceInfo, clickEx.Message));
                                                    }));
                                                    break;
                                                }
                                                i2 = i;
                                            }
                                        }
                                        else
                                        {
                                            base.Invoke(new MethodInvoker(delegate ()
                                            {
                                                this.dataGridView1.Rows[displayRowIndex].Cells[8].Value = "Không tìm thấy nút Play Sound";
                                            }));
                                        }
                                    }
                                    break;
                                }
                                value4 = "Unknown";
                                this.dataGridView1.Rows[displayRowIndex].Cells[8].Value = value4;
                            }
                        }
                        catch (Exception ex5)
                        {
                            base.Invoke(new MethodInvoker(delegate ()
                            {
                                this.richTextBox1.AppendText("Lỗi khi xử lý thiết bị: " + ex5.Message + "\r\n");
                            }));
                            break;
                        }
                    }
                    bool flag26 = !flag12;
                    if (flag26)
                    {
                        this.TatChromeKhiKhongThayName(index);
                        base.Invoke(new MethodInvoker(delegate ()
                        {
                            this.richTextBox1.AppendText("Không tìm thấy thiết bị iPhone trong danh sách \r\n");
                        }));
                        base.Invoke(new MethodInvoker(delegate ()
                        {
                            this.richTextBox1.AppendText("Không tìm thấy thiết bị iPhone trong danh sách");
                        }));
                    }
                }
                catch (Exception ex6)
                {
                    base.Invoke(new MethodInvoker(delegate ()
                    {
                        this.richTextBox1.AppendText("Lỗi: " + ex6.Message + "\r\n");
                    }));
                    this.KillChromeProcessesWithUserDataDir(deviceInfo);
                }
                finally
                {
                    try
                    {
                        bool flag27 = webDriver != null;
                        if (flag27)
                        {
                            List<IWebDriver> obj2 = Form1.activeDrivers;
                            lock (obj2)
                            {
                                Form1.activeDrivers.Remove(webDriver);
                                this.rowDriverMap.Remove(index);
                            }
                            webDriver.Quit();
                            base.Invoke(new MethodInvoker(delegate ()
                            {
                                this.richTextBox1.AppendText("Đã đóng Chrome instance cho thiết bị " + deviceInfo + "\r\n");
                            }));
                        }
                        bool flag29 = chromeDriverService != null;
                        if (flag29)
                        {
                            chromeDriverService.Dispose();
                        }
                    }
                    catch (Exception ex7)
                    {
                        base.Invoke(new MethodInvoker(delegate ()
                        {
                            this.richTextBox1.AppendText("Lỗi khi cleanup: " + ex7.Message + " \r\n");
                        }));
                    }
                }
            }
        }

        private void TatChromeKhiKhongThayName(int index)
        {
            try
            {
                int displayRowIndex = this.FindDisplayRowIndexByRealIndex(index);
                bool flag = displayRowIndex == -1;
                if (!flag)
                {
                    object value = this.dataGridView1.Rows[displayRowIndex].Cells[1].Value;
                    string deviceInfo = (((value != null) ? value.ToString() : null) ?? "");

                    bool flag2 = this.activeProcesses.ContainsKey(index);
                    if (flag2)
                    {
                        Form1.iCloudProcess iCloudProcess = this.activeProcesses[index];
                        iCloudProcess.IsRunning = false;
                        CancellationTokenSource cancellationTokenSource = iCloudProcess.CancellationTokenSource;
                        if (cancellationTokenSource != null)
                        {
                            cancellationTokenSource.Cancel();
                        }
                        this.richTextBox1.Invoke(new Action(delegate ()
                        {
                            this.richTextBox1.AppendText(string.Format("[{0:HH:mm:ss}] Đã tắt Chrome driver cho Row {1} vì bị đổi tên phone\r\n", DateTime.Now, index));
                        }));
                    }

                    IWebDriver webDriver = null;
                    List<IWebDriver> obj = Form1.activeDrivers;
                    lock (obj)
                    {
                        bool flag4 = this.rowDriverMap.ContainsKey(index);
                        if (flag4)
                        {
                            webDriver = this.rowDriverMap[index];
                            this.rowDriverMap.Remove(index);
                        }
                        else
                        {
                            foreach (IWebDriver webDriver2 in Form1.activeDrivers.ToList<IWebDriver>())
                            {
                                try
                                {
                                    string url = webDriver2.Url;
                                    bool flag5 = url.Contains("icloud.com");
                                    if (flag5)
                                    {
                                        webDriver = webDriver2;
                                        break;
                                    }
                                }
                                catch (Exception)
                                {
                                }
                            }
                        }

                        bool flag6 = webDriver != null;
                        if (flag6)
                        {
                            try
                            {
                                Form1.activeDrivers.Remove(webDriver);
                                webDriver.Quit();
                                string capturedDeviceInfo = deviceInfo; // Capture biến để dùng trong delegate
                                this.richTextBox1.Invoke(new Action(delegate ()
                                {
                                    this.richTextBox1.AppendText(string.Format("[{0:HH:mm:ss}] Đã tắt Chrome driver cho Row {1} vì đổi tên phone\r\n", DateTime.Now, index));
                                }));
                                this.DeleteChromeProfile(capturedDeviceInfo);
                            }
                            catch (Exception ex)
                            {
                                string capturedMessage = ex.Message; // Capture biến
                                this.richTextBox1.Invoke(new Action(delegate ()
                                {
                                    this.richTextBox1.AppendText(string.Format("[{0:HH:mm:ss}] Lỗi khi tắt Chrome khi rung do đổi tên phone: {1}\r\n", DateTime.Now, capturedMessage));
                                }));
                            }
                        }
                    }

                    this.dataGridView1.Rows[displayRowIndex].Cells[8].Value = "Bị Đổi Tên";
                    string finalDeviceInfo = deviceInfo; // Capture biến
                    this.richTextBox1.Invoke(new Action(delegate ()
                    {
                        this.richTextBox1.AppendText(string.Format("[{0:HH:mm:ss}] Đã tắt hoàn toàn process cho {1}\r\n", DateTime.Now, finalDeviceInfo));
                    }));
                }
            }
            catch (Exception ex)
            {
                string errorMessage = ex.Message; // Capture biến
                this.richTextBox1.Invoke(new Action(delegate ()
                {
                    this.richTextBox1.AppendText(string.Format("[{0:HH:mm:ss}] Lỗi tổng quát: {1}\r\n", DateTime.Now, errorMessage));
                }));
            }
        }

        private void KillChromeProcessesWithUserDataDir(string deviceInfo)
        {
            try
            {
                string value = "User_Data_" + deviceInfo;
                Process[] processesByName = Process.GetProcessesByName("chrome");
                Process[] array = processesByName;
                for (int i = 0; i < array.Length; i++)
                {
                    Process process = array[i];
                    try
                    {
                        string processCommandLine = this.GetProcessCommandLine(process.Id);
                        bool flag = !string.IsNullOrEmpty(processCommandLine) && processCommandLine.Contains(value);
                        if (flag)
                        {
                            this.richTextBox1.Invoke(new Action(delegate ()
                            {
                                this.richTextBox1.AppendText(string.Format("[{0:HH:mm:ss}] Đang kill Chrome process PID: {1} cho {2}\r\n", DateTime.Now, process.Id, deviceInfo));
                            }));
                            process.Kill();
                            process.WaitForExit(5000);
                        }
                    }
                    catch (Exception ex)
                    {
                        this.richTextBox1.Invoke(new Action(delegate ()
                        {
                            this.richTextBox1.AppendText(string.Format("[{0:HH:mm:ss}] Lỗi khi kill process {1}: {2}\r\n", DateTime.Now, process.Id, ex.Message));
                        }));
                    }
                }
                this.richTextBox1.Invoke(new Action(delegate ()
                {
                    this.richTextBox1.AppendText(string.Format("[{0:HH:mm:ss}] Đã kill tất cả Chrome processes cho {1}\r\n", DateTime.Now, deviceInfo));
                }));
            }
            catch (Exception ex3)
            {
                Exception ex2 = ex3;
                Exception ex = ex2;
                this.richTextBox1.Invoke(new Action(delegate ()
                {
                    this.richTextBox1.AppendText(string.Format("[{0:HH:mm:ss}] Lỗi KillChromeProcesses: {1}\r\n", DateTime.Now, ex.Message));
                }));
            }
        }
        private void LoadSavedAccounts()
        {
            savedAccounts = new Dictionary<string, string>();

            if (System.IO.File.Exists(accountsFilePath))
            {
                try
                {
                    string json = System.IO.File.ReadAllText(accountsFilePath);
                    savedAccounts = System.Text.Json.JsonSerializer.Deserialize<Dictionary<string, string>>(json) ?? new Dictionary<string, string>();
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Lỗi khi đọc file tài khoản: {ex.Message}");
                }
            }
        }
        private string GetProcessCommandLine(int processId)
        {
            try
            {
                using (ManagementObjectSearcher searcher = new ManagementObjectSearcher(
                    $"SELECT CommandLine FROM Win32_Process WHERE ProcessId = {processId}"))
                {
                    using (ManagementObjectCollection objects = searcher.Get())
                    {
                        foreach (ManagementObject obj in objects)
                        {
                            return obj["CommandLine"]?.ToString() ?? "";
                        }
                    }
                }
            }
            catch
            {
                return "";
            }
            return "";
        }
        private void LoadSavedAccounts1()
        {
            savedAccounts1 = new Dictionary<string, string>();

            if (System.IO.File.Exists(accountsFilePath1))
            {
                try
                {
                    string json = System.IO.File.ReadAllText(accountsFilePath1);
                    savedAccounts1 = System.Text.Json.JsonSerializer.Deserialize<Dictionary<string, string>>(json) ?? new Dictionary<string, string>();
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Lỗi khi đọc file tài khoản: {ex.Message}");
                }
            }
        }
        // Load danh sách username vào ComboBox
        private void LoadUsernamesIntoComboBox()
        {
            comboBoxUsername.Items.Clear();
            foreach (string username in savedAccounts.Keys)
            {
                comboBoxUsername.Items.Add(username);
            }
        }
        private void LoadUsernamesIntoComboBox1()
        {
            comboBoxUsername1.Items.Clear();
            foreach (string username in savedAccounts1.Keys)
            {
                comboBoxUsername1.Items.Add(username);
            }
        }
        // Xử lý khi chọn username từ ComboBox
        private void ComboBoxUsername_SelectedIndexChanged(object sender, EventArgs e)
        {
            string selectedUsername = comboBoxUsername.SelectedItem?.ToString();

            if (!string.IsNullOrEmpty(selectedUsername) && savedAccounts.ContainsKey(selectedUsername))
            {
                // Tự động điền password tương ứng
                textBoxPassIcloud.Text = DecryptPassword(savedAccounts[selectedUsername]);
            }
        }
        private void ComboBoxUsername1_SelectedIndexChanged(object sender, EventArgs e)
        {
            string selectedUsername = comboBoxUsername1.SelectedItem?.ToString();

            if (!string.IsNullOrEmpty(selectedUsername) && savedAccounts1.ContainsKey(selectedUsername))
            {
                // Tự động điền password tương ứng
                textBoxPass.Text = DecryptPassword(savedAccounts1[selectedUsername]);
            }
        }
        // Method để cleanup tất cả drivers khi đóng form
        private void CleanupAllDrivers()
        {
            lock (activeDrivers)
            {
                foreach (var driver in activeDrivers.ToList())
                {
                    try
                    {
                        driver?.Quit();
                    }
                    catch (Exception ex)
                    {
                        richTextBox1.AppendText($"Lỗi khi đóng driver: {ex.Message}");
                    }
                }
                activeDrivers.Clear();
            }
        }

        // Gọi method này trong Form_FormClosing event
        private void Form_FormClosing(object sender, FormClosingEventArgs e)
        {
            CleanupAllDrivers();
        }
        private void DisableFeatures()
        {
            button2.Enabled = false;
            button3.Enabled = false;
            button4.Enabled = false;
            button5.Enabled = false;
            button6.Enabled = false;
            button7.Enabled = false;
            button8.Enabled = false;
            button10.Enabled = false;
            button11.Enabled = false;
            button12.Enabled = false;
            comboBoxUsername1.Enabled = false;
            textBoxPass.Enabled = false;
            comboBoxUsername.Enabled = false;
            textBoxPassIcloud.Enabled = false;
            dataGridView1.Enabled = false;
            radioButtonAll.Enabled = false; 
            radioButtonLine1.Enabled = false;
            radioButtonLine2.Enabled = false;
            radioButtonxoa.Enabled = false; 
            button13.Enabled = false;
            button14.Enabled = false;
            button15.Enabled = false;
            button20.Enabled = false;
            button21.Enabled = false;
            button22.Enabled = false;
            button23.Enabled = false;
            button24.Enabled = false;
            button19.Enabled = false;   




        }
        private void button1_Click(object sender, EventArgs e)
        {
            Form1.shouldMerge = true;
            this.button5.Invoke(new MethodInvoker(delegate ()
            {
                this.radioButtonAll.Checked = true;
            }));
            List<string> userTomorowNew = this.GetUserTomorowNew("USR38122602", "SHP19263338");
            this.LoadAndSaveData(userTomorowNew);
        }
        private List<string> GetUserTomorowNew(string userId, string shopId)
        {
            List<string> result;
            string tenSo = "";
            try
            {
                if(shopId == "SHP19263338")
                {
                    tenSo = "Sổ 1";
                }else if(shopId == "SHP10275991")
                {
                    tenSo = "Sổ 2";
                }else if(shopId == "SHP93791247")
                {
                    tenSo = "Sổ 3";
                }else if(shopId == "SHP01201952")
                {
                    tenSo = "Sổ 6";
                }else if(shopId == "SHP78278175")
                {
                    tenSo = "Sổ 7";
                }
                List<string> list = new List<string>();
                string text = this.comboBoxUsername1.Text.Trim();
                string text2 = this.textBoxPass.Text;
                string text3 = new HttpRequest
                {
                    UserAgent = xNet.Http.ChromeUserAgent()
                }.Get(string.Concat(new string[]
                {
                    "https://annam.pro/api/get_contracts1.php?userId=",
                    userId,
                    "&shopId=",
                    shopId,
                    "&status=all"
                }), null).ToString();
                JObject jobject = JObject.Parse(text3);
                bool flag = jobject["tomorrowPayments"].Count<JToken>() > 0;
                if (flag)
                {
                    int num = jobject["tomorrowPayments"].Count<JToken>();
                    for (int i = 0; i < num; i++)
                    {
                        string text4 = jobject["tomorrowPayments"][i]["code_id"].ToString();
                        string text5 = jobject["tomorrowPayments"][i]["customer_name"].ToString();
                        string text6 = jobject["tomorrowPayments"][i]["money_per_period"].ToString().Split(new char[]
                        {
                            '.'
                        })[0];
                        string text7 = jobject["tomorrowPayments"][i]["current_status"].ToString();
                        string s = jobject["tomorrowPayments"][i]["next_payment_date"].ToString();
                        DateTime now;
                        bool flag2 = !DateTime.TryParseExact(s, "M/d/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out now);
                        if (flag2)
                        {
                            bool flag3 = !DateTime.TryParse(s, out now);
                            if (flag3)
                            {
                                now = DateTime.Now;
                            }
                        }
                        string text8 = now.ToString("dd/MM/yyyy");
                        list.Add(string.Concat(new string[]
                        {
                            text4,
                            "|",
                            text5,
                            "|",
                            text6,
                            "|",
                            text7,
                            " ",
                            text6,
                            " tiền họ|",
                            i.ToString(),
                            "|",tenSo,"|",
                            text8
                        }));
                    }
                    result = list;
                }
                else
                {
                    result = null;
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message, ex);
            }
            return result;
        }
        private List<string> GetUserQuaHan(string userId, string shopId)
        {
            List<string> result;
            string tenSo = "";
            try
            {
                // Xác định tên sổ
                if (shopId == "SHP19263338")
                {
                    tenSo = "Sổ 1";
                }
                else if (shopId == "SHP10275991")
                {
                    tenSo = "Sổ 2";
                }
                else if (shopId == "SHP93791247")
                {
                    tenSo = "Sổ 3";
                }
                else if (shopId == "SHP01201952")
                {
                    tenSo = "Sổ 6";
                }
                else if (shopId == "SHP78278175")
                {
                    tenSo = "Sổ 7";
                }

                // Lấy dữ liệu từ MongoDB
                string mongoData1 = this.GetDataFromMongoDB1();
                string mongoData2 = this.GetDataFromMongoDB2();

                Dictionary<string, Form1.DebtReminderManager.DebtReminderInfo> dbData1 = ParseMongoDBData(mongoData1);
                Dictionary<string, Form1.DebtReminderManager.DebtReminderInfo> dbData2 = ParseMongoDBData(mongoData2);

                List<string> list = new List<string>();
                string text = this.comboBoxUsername1.Text.Trim();
                string text2 = this.textBoxPass.Text;

                string text3 = new HttpRequest
                {
                    UserAgent = xNet.Http.ChromeUserAgent()
                }.Get(string.Concat(new string[]
                {
            "https://annam.pro/api/get_contracts1.php?userId=",
            userId,
            "&shopId=",
            shopId,
            "&status=all"
                }), null).ToString();

                JObject jobject = JObject.Parse(text3);
                bool flag = jobject["data"].Count<JToken>() > 0;

                if (flag)
                {
                    int num = jobject["data"].Count<JToken>();
                    for (int i = 0; i < num; i++)
                    {
                        string text7 = jobject["data"][i]["current_status"].ToString();

                        // Chỉ xử lý các record có current_status là "Quá hạn"
                        if (text7 == "Quá hạn")
                        {
                            string text4 = jobject["data"][i]["code_id"].ToString(); // MaHD
                            string text5 = jobject["data"][i]["customer_name"].ToString(); // TenKH
                            string s = jobject["data"][i]["next_payment_date"].ToString();

                            DateTime now;
                            bool flag2 = !DateTime.TryParseExact(s, "M/d/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out now);
                            if (flag2)
                            {
                                bool flag3 = !DateTime.TryParse(s, out now);
                                if (flag3)
                                {
                                    now = DateTime.Now;
                                }
                            }
                            string text8 = now.ToString("dd/MM/yyyy");

                            // Kiểm tra dữ liệu từ MongoDB
                            string lineInfo = "";
                            string noteInfo = "";
                            string userIcloud = text;
                            string passIcloud = text2;
                            string namePhoneChange = "";
                            string dateInfo = "";

                            Form1.DebtReminderManager.DebtReminderInfo dbInfo = null;

                            // Tạo key để tìm trong MongoDB (MaHD|TenKH)
                            string searchKey = text4 + "|" + text5;

                            // Kiểm tra trong MongoDB1 trước
                            if (dbData1.ContainsKey(searchKey))
                            {
                                dbInfo = dbData1[searchKey];
                            }
                            // Nếu không có trong MongoDB1, kiểm tra MongoDB2
                            else if (dbData2.ContainsKey(searchKey))
                            {
                                dbInfo = dbData2[searchKey];
                            }

                            // Nếu tìm thấy trong MongoDB, lấy thông tin
                            if (dbInfo != null)
                            {
                                lineInfo = dbInfo.Line ?? "";
                                noteInfo = dbInfo.Note ?? "";
                                userIcloud = !string.IsNullOrEmpty(dbInfo.userIcloud) ? dbInfo.userIcloud : text;
                                passIcloud = !string.IsNullOrEmpty(dbInfo.PassIcloud) ? dbInfo.PassIcloud : text2;
                                namePhoneChange = dbInfo.NamePhoneChange ?? "";
                                dateInfo = dbInfo.date ?? "";
                            }

                            // Thêm vào danh sách kết quả theo thứ tự giống LoadDataByMonGoDB
                            // Format: Key|Note|UserIcloud|PassIcloud|NamePhoneChange|STT|Line|Date|TenSo
                            list.Add(string.Concat(new string[]
                            {
                        searchKey,           // [0] Key (MaHD|TenKH)
                        "|",
                        noteInfo,            // [1] Note
                        "|",
                        userIcloud,          // [2] UserIcloud
                        "|",
                        passIcloud,          // [3] PassIcloud
                        "|",
                        namePhoneChange,     // [4] NamePhoneChange
                        "|",
                        i.ToString(),        // [5] STT
                        "|",
                        lineInfo,            // [6] Line
                        "|",
                        text8,               // [7] Date
                        "|",
                        tenSo                // [8] TenSo
                            }));
                        }
                    }
                    result = list;
                }
                else
                {
                    result = null;
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message, ex);
            }
            return result;
        }
        private Dictionary<string, Form1.DebtReminderManager.DebtReminderInfo> ParseMongoDBData(string jsonData)
        {
            try
            {
                if (string.IsNullOrEmpty(jsonData))
                    return new Dictionary<string, Form1.DebtReminderManager.DebtReminderInfo>();

                return JsonConvert.DeserializeObject<Dictionary<string, Form1.DebtReminderManager.DebtReminderInfo>>(jsonData)
                       ?? new Dictionary<string, Form1.DebtReminderManager.DebtReminderInfo>();
            }
            catch
            {
                return new Dictionary<string, Form1.DebtReminderManager.DebtReminderInfo>();
            }
        }
        private void button9_Click_1(object sender, EventArgs e)
        {
            string userIcloud = comboBoxUsername1.SelectedItem?.ToString();
            string passIcloud = textBoxPass.Text?.Trim();

            if (string.IsNullOrEmpty(userIcloud) || string.IsNullOrEmpty(passIcloud))
            {
                MessageBox.Show("Vui lòng nhập đầy đủ thông tin acount 2Gold", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            List<string> dsKH = login2gold2();
            LoadAndSaveData(dsKH);
        }

        // Hàm DisplayData mới cho Dictionary<string, DebtReminderInfo>
        private void DisplayData(Dictionary<string, DebtReminderInfo> dsKH)
        {
            // Lưu trạng thái hiện tại trước khi xóa
            SaveCurrentStatus();
            dataGridView1.Rows.Clear();

            int count = dsKH.Count;
            int startIndex = dataGridView1.Rows.Count; // Lấy vị trí hiện tại để thêm vào

            // Đảm bảo có đủ số dòng
            while (dataGridView1.Rows.Count < startIndex + count)
            {
                dataGridView1.Rows.Add();
            }

            int i = 0;
            foreach (var kvp in dsKH)
            {
                DebtReminderInfo debtInfo = kvp.Value;

                // Thêm vào vị trí startIndex + i thay vì i
                dataGridView1.Rows[startIndex + i].Cells[0].Value = debtInfo.MaHD;
                dataGridView1.Rows[startIndex + i].Cells[1].Value = debtInfo.TenKh;
                dataGridView1.Rows[startIndex + i].Cells[9].Value = debtInfo.NamePhoneChange;
                dataGridView1.Rows[startIndex + i].Cells[3].Value = debtInfo.Note;
                dataGridView1.Rows[startIndex + i].Cells[10].Value = debtInfo.userIcloud;
                dataGridView1.Rows[startIndex + i].Cells[11].Value = debtInfo.PassIcloud;

                i++;
            }

            // Khôi phục trạng thái sau khi hiển thị dữ liệu
            RestoreStatus();
            reminderManager.LoadReminderData(dataGridView1);
        }
        private void DisplayData(List<string> dsKH)
        {
            this.SaveCurrentStatus();
            this.dataGridView1.Rows.Clear();
            int count = dsKH.Count;
            for (int i = 0; i < count; i++)
            {
                string text = dsKH[i];
                string[] array = text.Split(new char[]
                {
                    '|'
                });
                bool flag = array[0] == "";
                if (!flag)
                {
                    bool flag2 = array.Length == 7;
                    if (flag2)
                    {
                        int index = this.dataGridView1.Rows.Add();
                        this.dataGridView1.Rows[index].Cells[0].Value = array[0];
                        this.dataGridView1.Rows[index].Cells[1].Value = array[1];
                        this.dataGridView1.Rows[index].Cells[3].Value = array[3];
                        this.dataGridView1.Rows[index].Cells[13].Value = array[4];
                        this.dataGridView1.Rows[index].Cells[8].Value = array[5];
                        this.dataGridView1.Rows[index].Cells[14].Value = array[6];
                    }
                    else
                    {
                        int index2 = this.dataGridView1.Rows.Add();
                        this.dataGridView1.Rows[index2].Cells[0].Value = array[0];
                        this.dataGridView1.Rows[index2].Cells[2].Value = array[7];
                        this.dataGridView1.Rows[index2].Cells[1].Value = array[1];
                        this.dataGridView1.Rows[index2].Cells[10].Value = array[3];
                        this.dataGridView1.Rows[index2].Cells[11].Value = array[4];
                        this.dataGridView1.Rows[index2].Cells[3].Value = array[2];
                        this.dataGridView1.Rows[index2].Cells[9].Value = array[5];
                        this.dataGridView1.Rows[index2].Cells[13].Value = array[6];
                        this.dataGridView1.Rows[index2].Cells[14].Value = array[8];
                        this.dataGridView1.Rows[index2].Cells[8].Value = array[9];
                    }
                }
            }
            this.RestoreStatus();
            this.reminderManager.LoadReminderData(this.dataGridView1);
        }
        private List<string> login2gold()
        {
            try {
                List<string> info = new List<string>();
                string user = comboBoxUsername1.Text.Trim();
                string pass = textBoxPass.Text;
                HttpRequest request = new HttpRequest();
                request.UserAgent = xNet.Http.ChromeUserAgent();
                string cookie = request.Get("https://2gold.biz/").Cookies.ToString();
                request.Cookies = new CookieDictionary();
                // gán cookie vào http
                for (int c = 0; c < cookie.Split(';').Length; c++)
                {
                    try
                    {
                        string name = cookie.Split(';')[c].Split('=')[0].Trim();
                        string value = cookie.Split(';')[c].Substring(cookie.Split(';')[c].IndexOf('=') + 1).Trim();
                        if (request.Cookies.ContainsKey(name))
                            request.Cookies.Remove(name);

                        request.Cookies.Add(name, value);
                    }
                    catch (Exception ex) { }
                }
                request.AddHeader("accept", "*/*");
                request.AddHeader("accept-language", "en-US,en;q=0.9,vi;q=0.8");
                request.AddHeader("origin", "https://2gold.biz");
                request.AddHeader("priority", "u=1, i");
                request.AddHeader("referer", "https://2gold.biz/User/Login");
                request.AddHeader("sec-ch-ua", "\"Google Chrome\";v=\"137\", \"Chromium\";v=\"137\", \"Not/A)Brand\";v=\"24\"");
                request.AddHeader("sec-ch-ua-mobile", "?1");
                request.AddHeader("sec-ch-ua-platform", "\"Android\"");
                request.AddHeader("sec-fetch-dest", "empty");
                request.AddHeader("sec-fetch-mode", "cors");
                request.AddHeader("sec-fetch-site", "same-origin");
                request.AddHeader("x-kl-ajax-request", "Ajax_Request");
                request.AddHeader("x-requested-with", "XMLHttpRequest");
                var body = "Username=" + user + "&Password=" + pass;
                string response = request.Post("https://2gold.biz/User/ProcessLogin", body, "application/x-www-form-urlencoded; charset=UTF-8").ToString();
                JObject json = JObject.Parse(response);
                if (json["Message"].ToString().Contains("Đăng nhập thành công"))
                {
                    string UserID = json["Data"]["UserID"].ToString();
                    string ShopID = json["Data"]["ShopID"].ToString();
                    string Token = json["Data"]["Token"].ToString();
                    request.AddHeader("accept", "*/*");
                    request.AddHeader("accept-language", "en-US,en;q=0.9,vi;q=0.8");
                    request.AddHeader("origin", "https://2gold.biz");
                    request.AddHeader("priority", "u=1, i");
                    request.AddHeader("referer", "https://2gold.biz/");
                    request.AddHeader("sec-ch-ua", "\"Google Chrome\";v=\"137\", \"Chromium\";v=\"137\", \"Not/A)Brand\";v=\"24\"");
                    request.AddHeader("sec-ch-ua-mobile", "?0");
                    request.AddHeader("sec-ch-ua-platform", "\"Windows\"");
                    request.AddHeader("sec-fetch-dest", "empty");
                    request.AddHeader("sec-fetch-mode", "cors");
                    request.AddHeader("sec-fetch-site", "same-site");
                    var body2 = "UserID="+ UserID + "&ShopID="+ ShopID + "&Token="+ Token;
                    string response2 = request.Post("https://api.2gold.biz/api/PaymentNotify/GetNotifyCount", body2, "application/x-www-form-urlencoded; charset=UTF-8").ToString();
                    json = JObject.Parse(response2);
                    if (json["Message"].ToString().Contains("Lấy số liệu nhắc nợ thành công"))
                    {
                        int InstallmentTotal = int.Parse(json["Data"]["InstallmentTotal"].ToString());
                        double result = InstallmentTotal / 50.0;
                        int pageCount = (int)Math.Ceiling(result);
                        string firtNum = "";
                        string secondNum = "";
                        for (int j = 0; j < pageCount; j++)
                        {
                            
                            if(j == 0)
                            {
                                firtNum = "1";
                            }
                            else
                            {
                                firtNum = (j+1).ToString();
                                secondNum = pageCount.ToString();
                            }
                            request.AddHeader("accept", "application/json, text/javascript, */*; q=0.01");
                            request.AddHeader("accept-language", "en-US,en;q=0.9,vi;q=0.8");
                            request.AddHeader("origin", "https://2gold.biz");
                            request.AddHeader("priority", "u=1, i");
                            request.AddHeader("referer", "https://2gold.biz/");
                            request.AddHeader("sec-ch-ua", "\"Google Chrome\";v=\"137\", \"Chromium\";v=\"137\", \"Not/A)Brand\";v=\"24\"");
                            request.AddHeader("sec-ch-ua-mobile", "?0");
                            request.AddHeader("sec-ch-ua-platform", "\"Windows\"");
                            request.AddHeader("sec-fetch-dest", "empty");
                            request.AddHeader("sec-fetch-mode", "cors");
                            request.AddHeader("sec-fetch-site", "same-site");
                            var body1 = "datatable%5Bpagination%5D%5Bpage%5D=" + firtNum + "&datatable%5Bpagination%5D%5Bpages%5D=" + secondNum + "&datatable%5Bpagination%5D%5Bperpage%5D=" + InstallmentTotal + "&datatable%5Bsort%5D%5Bsort%5D=&datatable%5Bsort%5D%5Bfield%5D=&datatable%5Bquery%5D%5BgeneralSearch%5D=&datatable%5Bquery%5D%5BStatus%5D=1000&datatable%5Bquery%5D%5BShopID%5D=" + ShopID + "&datatable%5Bquery%5D%5BTypeProduct%5D=2&datatable%5Bquery%5D%5BUserID%5D=" + UserID + "&datatable%5Bquery%5D%5BToken%5D=" + Token + "&datatable%5Bquery%5D%5BStaffId%5D=0&datatable%5Bquery%5D%5BPageCurrent%5D=1&datatable%5Bquery%5D%5BPerPageCurrent%5D=&datatable%5Bquery%5D%5BcolumnCurrent%5D=&datatable%5Bquery%5D%5BsortCurrent%5D=";
                            string response1 = request.Post("https://api.2gold.biz/api/PaymentNotify/List", body1, "application/x-www-form-urlencoded; charset=UTF-8").ToString();
                            json = JObject.Parse(response1);
                            if (json["data"].Count() > 0)
                            {
                                int countData = json["data"].Count();
                                for (int i = 0; i < countData; i++)
                                {
                                    string maHD = json["data"][i]["StrCodeID"].ToString();
                                    string tenKH = json["data"][i]["CustomerName"].ToString();
                                    string diachi = json["data"][i]["CustomerAddress"].ToString();
                                    string noCu = json["data"][i]["DebitMoney"].ToString();
                                    string tiencandong = json["data"][i]["PaymentNotify"].ToString();
                                    string lydo = json["data"][i]["Note"].ToString();
                                    info.Add(maHD + "|" + tenKH +  "|" + tiencandong + "|" + lydo);
                                }

                            }
                        }
                        

                    }


                    
                }
                else
                {
                    throw new Exception("Đăng Nhập Thất Bại");
                }
                return info;
            }
            catch(Exception ex) 
            { 
                throw new Exception(ex.Message, ex);
            }

        }
        private List<string> login2gold2()
        {
            try
            {
                List<string> info = new List<string>();
                string user = comboBoxUsername1.Text.Trim();
                string pass = textBoxPass.Text;
                HttpRequest request = new HttpRequest();
                request.UserAgent = xNet.Http.ChromeUserAgent();
                string cookie = request.Get("https://2gold.biz/").Cookies.ToString();
                request.Cookies = new CookieDictionary();
                // gán cookie vào http
                for (int c = 0; c < cookie.Split(';').Length; c++)
                {
                    try
                    {
                        string name = cookie.Split(';')[c].Split('=')[0].Trim();
                        string value = cookie.Split(';')[c].Substring(cookie.Split(';')[c].IndexOf('=') + 1).Trim();
                        if (request.Cookies.ContainsKey(name))
                            request.Cookies.Remove(name);

                        request.Cookies.Add(name, value);
                    }
                    catch (Exception ex) { }
                }
                request.AddHeader("accept", "*/*");
                request.AddHeader("accept-language", "en-US,en;q=0.9,vi;q=0.8");
                request.AddHeader("origin", "https://2gold.biz");
                request.AddHeader("priority", "u=1, i");
                request.AddHeader("referer", "https://2gold.biz/User/Login");
                request.AddHeader("sec-ch-ua", "\"Google Chrome\";v=\"137\", \"Chromium\";v=\"137\", \"Not/A)Brand\";v=\"24\"");
                request.AddHeader("sec-ch-ua-mobile", "?1");
                request.AddHeader("sec-ch-ua-platform", "\"Android\"");
                request.AddHeader("sec-fetch-dest", "empty");
                request.AddHeader("sec-fetch-mode", "cors");
                request.AddHeader("sec-fetch-site", "same-origin");
                request.AddHeader("x-kl-ajax-request", "Ajax_Request");
                request.AddHeader("x-requested-with", "XMLHttpRequest");
                var body = "Username=" + user + "&Password=" + pass;
                string response = request.Post("https://2gold.biz/User/ProcessLogin", body, "application/x-www-form-urlencoded; charset=UTF-8").ToString();
                JObject json = JObject.Parse(response);
                if (json["Message"].ToString().Contains("Đăng nhập thành công"))
                {
                    string UserID = json["Data"]["UserID"].ToString();
                    //string ShopID = json["Data"]["ShopID"].ToString();
                    string ShopID = "53576";
                    string Token = json["Data"]["Token"].ToString();
                    request.AddHeader("accept", "*/*");
                    request.AddHeader("accept-language", "en-US,en;q=0.9,vi;q=0.8");
                    request.AddHeader("origin", "https://2gold.biz");
                    request.AddHeader("priority", "u=1, i");
                    request.AddHeader("referer", "https://2gold.biz/");
                    request.AddHeader("sec-ch-ua", "\"Google Chrome\";v=\"137\", \"Chromium\";v=\"137\", \"Not/A)Brand\";v=\"24\"");
                    request.AddHeader("sec-ch-ua-mobile", "?0");
                    request.AddHeader("sec-ch-ua-platform", "\"Windows\"");
                    request.AddHeader("sec-fetch-dest", "empty");
                    request.AddHeader("sec-fetch-mode", "cors");
                    request.AddHeader("sec-fetch-site", "same-site");
                    var body2 = "UserID=" + UserID + "&ShopID=" + ShopID + "&Token=" + Token;
                    string response2 = request.Post("https://api.2gold.biz/api/PaymentNotify/GetNotifyCount", body2, "application/x-www-form-urlencoded; charset=UTF-8").ToString();
                    json = JObject.Parse(response2);
                    if (json["Message"].ToString().Contains("Lấy số liệu nhắc nợ thành công"))
                    {
                        int InstallmentTotal = int.Parse(json["Data"]["InstallmentTotal"].ToString());
                        double result = InstallmentTotal / 50.0;
                        int pageCount = (int)Math.Ceiling(result);
                        string firtNum = "";
                        string secondNum = "";
                        for (int j = 0; j < pageCount; j++)
                        {

                            if (j == 0)
                            {
                                firtNum = "1";
                            }
                            else
                            {
                                firtNum = (j + 1).ToString();
                                secondNum = pageCount.ToString();
                            }
                            request.AddHeader("accept", "application/json, text/javascript, */*; q=0.01");
                            request.AddHeader("accept-language", "en-US,en;q=0.9,vi;q=0.8");
                            request.AddHeader("origin", "https://2gold.biz");
                            request.AddHeader("priority", "u=1, i");
                            request.AddHeader("referer", "https://2gold.biz/");
                            request.AddHeader("sec-ch-ua", "\"Google Chrome\";v=\"137\", \"Chromium\";v=\"137\", \"Not/A)Brand\";v=\"24\"");
                            request.AddHeader("sec-ch-ua-mobile", "?0");
                            request.AddHeader("sec-ch-ua-platform", "\"Windows\"");
                            request.AddHeader("sec-fetch-dest", "empty");
                            request.AddHeader("sec-fetch-mode", "cors");
                            request.AddHeader("sec-fetch-site", "same-site");
                            var body1 = "datatable%5Bpagination%5D%5Bpage%5D=" + firtNum + "&datatable%5Bpagination%5D%5Bpages%5D=" + secondNum + "&datatable%5Bpagination%5D%5Bperpage%5D=" + InstallmentTotal + "&datatable%5Bsort%5D%5Bsort%5D=&datatable%5Bsort%5D%5Bfield%5D=&datatable%5Bquery%5D%5BgeneralSearch%5D=&datatable%5Bquery%5D%5BStatus%5D=1000&datatable%5Bquery%5D%5BShopID%5D=" + ShopID + "&datatable%5Bquery%5D%5BTypeProduct%5D=2&datatable%5Bquery%5D%5BUserID%5D=" + UserID + "&datatable%5Bquery%5D%5BToken%5D=" + Token + "&datatable%5Bquery%5D%5BStaffId%5D=0&datatable%5Bquery%5D%5BPageCurrent%5D=1&datatable%5Bquery%5D%5BPerPageCurrent%5D=&datatable%5Bquery%5D%5BcolumnCurrent%5D=&datatable%5Bquery%5D%5BsortCurrent%5D=";
                            string response1 = request.Post("https://api.2gold.biz/api/PaymentNotify/List", body1, "application/x-www-form-urlencoded; charset=UTF-8").ToString();
                            json = JObject.Parse(response1);
                            if (json["data"].Count() > 0)
                            {
                                int countData = json["data"].Count();
                                for (int i = 0; i < countData; i++)
                                {
                                    string maHD = json["data"][i]["StrCodeID"].ToString();
                                    string tenKH = json["data"][i]["CustomerName"].ToString();
                                    string diachi = json["data"][i]["CustomerAddress"].ToString();
                                    string noCu = json["data"][i]["DebitMoney"].ToString();
                                    string tiencandong = json["data"][i]["PaymentNotify"].ToString();
                                    string lydo = json["data"][i]["Note"].ToString();
                                    info.Add(maHD + "|" + tenKH + "|" + tiencandong + "|" + lydo);
                                }

                            }
                        }


                    }



                }
                else
                {
                    throw new Exception("Đăng Nhập Thất Bại");
                }
                return info;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message, ex);
            }

        }
        private List<string> GetUserTomorow()
        {
            try
            {
                List<string> info = new List<string>();
                string user = comboBoxUsername1.Text.Trim();
                string pass = textBoxPass.Text;
                HttpRequest request = new HttpRequest();
                request.UserAgent = xNet.Http.ChromeUserAgent();
                string cookie = request.Get("https://2gold.biz/").Cookies.ToString();
                request.Cookies = new CookieDictionary();
                // gán cookie vào http
                for (int c = 0; c < cookie.Split(';').Length; c++)
                {
                    try
                    {
                        string name = cookie.Split(';')[c].Split('=')[0].Trim();
                        string value = cookie.Split(';')[c].Substring(cookie.Split(';')[c].IndexOf('=') + 1).Trim();
                        if (request.Cookies.ContainsKey(name))
                            request.Cookies.Remove(name);

                        request.Cookies.Add(name, value);
                    }
                    catch (Exception ex) { }
                }
                request.AddHeader("accept", "*/*");
                request.AddHeader("accept-language", "en-US,en;q=0.9,vi;q=0.8");
                request.AddHeader("origin", "https://2gold.biz");
                request.AddHeader("priority", "u=1, i");
                request.AddHeader("referer", "https://2gold.biz/User/Login");
                request.AddHeader("sec-ch-ua", "\"Google Chrome\";v=\"137\", \"Chromium\";v=\"137\", \"Not/A)Brand\";v=\"24\"");
                request.AddHeader("sec-ch-ua-mobile", "?1");
                request.AddHeader("sec-ch-ua-platform", "\"Android\"");
                request.AddHeader("sec-fetch-dest", "empty");
                request.AddHeader("sec-fetch-mode", "cors");
                request.AddHeader("sec-fetch-site", "same-origin");
                request.AddHeader("x-kl-ajax-request", "Ajax_Request");
                request.AddHeader("x-requested-with", "XMLHttpRequest");
                var body = "Username=" + user + "&Password=" + pass;
                string response = request.Post("https://2gold.biz/User/ProcessLogin", body, "application/x-www-form-urlencoded; charset=UTF-8").ToString();
                JObject json = JObject.Parse(response);
                if (json["Message"].ToString().Contains("Đăng nhập thành công"))
                {
                    string UserID = json["Data"]["UserID"].ToString();
                    string ShopID = json["Data"]["ShopID"].ToString();
                    string Token = json["Data"]["Token"].ToString();
                    request.AddHeader("accept", "application/json, text/javascript, */*; q=0.01");
                    request.AddHeader("accept-language", "en-US,en;q=0.9,vi;q=0.8");
                    request.AddHeader("origin", "https://2gold.biz");
                    request.AddHeader("priority", "u=1, i");
                    request.AddHeader("referer", "https://2gold.biz/");
                    request.AddHeader("sec-ch-ua", "\"Google Chrome\";v=\"137\", \"Chromium\";v=\"137\", \"Not/A)Brand\";v=\"24\"");
                    request.AddHeader("sec-ch-ua-mobile", "?0");
                    request.AddHeader("sec-ch-ua-platform", "\"Windows\"");
                    request.AddHeader("sec-fetch-dest", "empty");
                    request.AddHeader("sec-fetch-mode", "cors");
                    request.AddHeader("sec-fetch-site", "same-site");
                    var body1 = "datatable%5Bpagination%5D%5Bpage%5D=1&datatable%5Bpagination%5D%5Bpages%5D=1&datatable%5Bpagination%5D%5Bperpage%5D=100&datatable%5Bpagination%5D%5Btotal%5D&datatable%5Bsort%5D%5Bsort%5D=&datatable%5Bsort%5D%5Bfield%5D=&datatable%5Bquery%5D%5BgeneralSearch%5D=&datatable%5Bquery%5D%5BStatus%5D=8&datatable%5Bquery%5D%5BShopID%5D="+ShopID+"&datatable%5Bquery%5D%5BTypeProduct%5D=2&datatable%5Bquery%5D%5BUserID%5D="+UserID+"&datatable%5Bquery%5D%5BToken%5D="+Token+"&datatable%5Bquery%5D%5BStaffId%5D=0&datatable%5Bquery%5D%5BPageCurrent%5D=1&datatable%5Bquery%5D%5BPerPageCurrent%5D=50&datatable%5Bquery%5D%5BcolumnCurrent%5D=&datatable%5Bquery%5D%5BsortCurrent%5D=";
                    string response1 = request.Post("https://api.2gold.biz/api/PaymentNotify/List", body1, "application/x-www-form-urlencoded; charset=UTF-8").ToString();
                    json = JObject.Parse(response1);
                    if (json["data"].Count() > 0)
                    {
                        int countData = json["data"].Count();
                        for (int i = 0; i < countData; i++)
                        {
                            string maHD = json["data"][i]["StrCodeID"].ToString();
                            string tenKH = json["data"][i]["CustomerName"].ToString();
                            string diachi = json["data"][i]["CustomerAddress"].ToString();
                            string noCu = json["data"][i]["DebitMoney"].ToString();
                            string tiencandong = json["data"][i]["PaymentNotify"].ToString();
                            string lydo = json["data"][i]["Note"].ToString();
                            info.Add(maHD + "|" + tenKH + "|" + tiencandong + "|" + lydo);
                        }

                    }



                }
                else
                {
                    throw new Exception("Đăng Nhập Thất Bại");
                }
                return info;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message, ex);
            }

        }
        private List<string> GetUserTomorow2()
        {
            try
            {
                List<string> info = new List<string>();
                string user = comboBoxUsername1.Text.Trim();
                string pass = textBoxPass.Text;
                HttpRequest request = new HttpRequest();
                request.UserAgent = xNet.Http.ChromeUserAgent();
                string cookie = request.Get("https://2gold.biz/").Cookies.ToString();
                request.Cookies = new CookieDictionary();
                // gán cookie vào http
                for (int c = 0; c < cookie.Split(';').Length; c++)
                {
                    try
                    {
                        string name = cookie.Split(';')[c].Split('=')[0].Trim();
                        string value = cookie.Split(';')[c].Substring(cookie.Split(';')[c].IndexOf('=') + 1).Trim();
                        if (request.Cookies.ContainsKey(name))
                            request.Cookies.Remove(name);

                        request.Cookies.Add(name, value);
                    }
                    catch (Exception ex) { }
                }
                request.AddHeader("accept", "*/*");
                request.AddHeader("accept-language", "en-US,en;q=0.9,vi;q=0.8");
                request.AddHeader("origin", "https://2gold.biz");
                request.AddHeader("priority", "u=1, i");
                request.AddHeader("referer", "https://2gold.biz/User/Login");
                request.AddHeader("sec-ch-ua", "\"Google Chrome\";v=\"137\", \"Chromium\";v=\"137\", \"Not/A)Brand\";v=\"24\"");
                request.AddHeader("sec-ch-ua-mobile", "?1");
                request.AddHeader("sec-ch-ua-platform", "\"Android\"");
                request.AddHeader("sec-fetch-dest", "empty");
                request.AddHeader("sec-fetch-mode", "cors");
                request.AddHeader("sec-fetch-site", "same-origin");
                request.AddHeader("x-kl-ajax-request", "Ajax_Request");
                request.AddHeader("x-requested-with", "XMLHttpRequest");
                var body = "Username=" + user + "&Password=" + pass;
                string response = request.Post("https://2gold.biz/User/ProcessLogin", body, "application/x-www-form-urlencoded; charset=UTF-8").ToString();
                JObject json = JObject.Parse(response);
                if (json["Message"].ToString().Contains("Đăng nhập thành công"))
                {
                    string UserID = json["Data"]["UserID"].ToString();
                    string ShopID = "53576";
                    string Token = json["Data"]["Token"].ToString();
                    request.AddHeader("accept", "application/json, text/javascript, */*; q=0.01");
                    request.AddHeader("accept-language", "en-US,en;q=0.9,vi;q=0.8");
                    request.AddHeader("origin", "https://2gold.biz");
                    request.AddHeader("priority", "u=1, i");
                    request.AddHeader("referer", "https://2gold.biz/");
                    request.AddHeader("sec-ch-ua", "\"Google Chrome\";v=\"137\", \"Chromium\";v=\"137\", \"Not/A)Brand\";v=\"24\"");
                    request.AddHeader("sec-ch-ua-mobile", "?0");
                    request.AddHeader("sec-ch-ua-platform", "\"Windows\"");
                    request.AddHeader("sec-fetch-dest", "empty");
                    request.AddHeader("sec-fetch-mode", "cors");
                    request.AddHeader("sec-fetch-site", "same-site");
                    var body1 = "datatable%5Bpagination%5D%5Bpage%5D=1&datatable%5Bpagination%5D%5Bpages%5D=1&datatable%5Bpagination%5D%5Bperpage%5D=100&datatable%5Bpagination%5D%5Btotal%5D&datatable%5Bsort%5D%5Bsort%5D=&datatable%5Bsort%5D%5Bfield%5D=&datatable%5Bquery%5D%5BgeneralSearch%5D=&datatable%5Bquery%5D%5BStatus%5D=8&datatable%5Bquery%5D%5BShopID%5D=" + ShopID + "&datatable%5Bquery%5D%5BTypeProduct%5D=2&datatable%5Bquery%5D%5BUserID%5D=" + UserID + "&datatable%5Bquery%5D%5BToken%5D=" + Token + "&datatable%5Bquery%5D%5BStaffId%5D=0&datatable%5Bquery%5D%5BPageCurrent%5D=1&datatable%5Bquery%5D%5BPerPageCurrent%5D=50&datatable%5Bquery%5D%5BcolumnCurrent%5D=&datatable%5Bquery%5D%5BsortCurrent%5D=";
                    string response1 = request.Post("https://api.2gold.biz/api/PaymentNotify/List", body1, "application/x-www-form-urlencoded; charset=UTF-8").ToString();
                    json = JObject.Parse(response1);
                    if (json["data"].Count() > 0)
                    {
                        int countData = json["data"].Count();
                        for (int i = 0; i < countData; i++)
                        {
                            string maHD = json["data"][i]["StrCodeID"].ToString();
                            string tenKH = json["data"][i]["CustomerName"].ToString();
                            string diachi = json["data"][i]["CustomerAddress"].ToString();
                            string noCu = json["data"][i]["DebitMoney"].ToString();
                            string tiencandong = json["data"][i]["PaymentNotify"].ToString();
                            string lydo = json["data"][i]["Note"].ToString();
                            info.Add(maHD + "|" + tenKH + "|" + tiencandong + "|" + lydo);
                        }

                    }



                }
                else
                {
                    throw new Exception("Đăng Nhập Thất Bại");
                }
                return info;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message, ex);
            }

        }
        public static string GetCookies(IWebDriver chromeDriver)
        {
            string cookie = "";
            try
            {
                for (int i = 0; i < chromeDriver.Manage().Cookies.AllCookies.Count; i++)
                {
                    try
                    {
                        string name = chromeDriver.Manage().Cookies.AllCookies.ElementAt(i).Name.ToString().Trim();
                        string value = chromeDriver.Manage().Cookies.AllCookies.ElementAt(i).Value.ToString().Trim();
                        if (!"".Equals(name))
                            cookie += $"{name}={value};";
                    }
                    catch (Exception) { }
                }
            }
            catch { }
            return cookie;

        }

        private void button2_Click(object sender, EventArgs e)
        {
            string username = comboBoxUsername.Text.Trim();
            string password = textBoxPassIcloud.Text;

            if (string.IsNullOrEmpty(username) || string.IsNullOrEmpty(password))
            {
                MessageBox.Show("Vui lòng nhập đầy đủ thông tin để lưu!");
                return;
            }

            SaveAccount(username, password);
            MessageBox.Show("Đã lưu tài khoản thành công!");

        }
        // Lưu tài khoản vào file
        private void SaveAccount(string username, string password)
        {
            string encryptedPassword = EncryptPassword(password);

            if (!savedAccounts.ContainsKey(username))
            {
                savedAccounts[username] = encryptedPassword;
                comboBoxUsername.Items.Add(username);
            }
            else
            {
                savedAccounts[username] = encryptedPassword;
            }

            SaveAccountsToFile();
        }
        private void SaveAccount1(string username, string password)
        {
            string encryptedPassword = EncryptPassword(password);

            if (!savedAccounts1.ContainsKey(username))
            {
                savedAccounts1[username] = encryptedPassword;
                comboBoxUsername1.Items.Add(username);
            }
            else
            {
                savedAccounts1[username] = encryptedPassword;
            }

            SaveAccountsToFile1();
        }

        // Lưu dictionary vào file JSON
        private void SaveAccountsToFile()
        {
            try
            {
                string json = System.Text.Json.JsonSerializer.Serialize(savedAccounts, new JsonSerializerOptions { WriteIndented = true });
                System.IO.File.WriteAllText(accountsFilePath, json);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi khi lưu file: {ex.Message}");
            }
        }
        private void SaveAccountsToFile1()
        {
            try
            {
                string json = System.Text.Json.JsonSerializer.Serialize(savedAccounts1, new JsonSerializerOptions { WriteIndented = true });
                System.IO.File.WriteAllText(accountsFilePath1, json);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi khi lưu file: {ex.Message}");
            }
        }

        // Mã hóa đơn giản password (nên sử dụng thuật toán mạnh hơn trong thực tế)
        private string EncryptPassword(string password)
        {
            // Đây là mã hóa đơn giản, trong thực tế nên dùng AES hoặc các thuật toán mạnh hơn
            byte[] data = System.Text.Encoding.UTF8.GetBytes(password);
            return Convert.ToBase64String(data);
        }

        // Giải mã password
        private string DecryptPassword(string encryptedPassword)
        {
            try
            {
                byte[] data = Convert.FromBase64String(encryptedPassword);
                return System.Text.Encoding.UTF8.GetString(data);
            }
            catch
            {
                return encryptedPassword; // Trả về nguyên gốc nếu không thể giải mã
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string username = comboBoxUsername1.Text.Trim();
            string password = textBoxPass.Text;

            if (string.IsNullOrEmpty(username) || string.IsNullOrEmpty(password))
            {
                MessageBox.Show("Vui lòng nhập đầy đủ thông tin để lưu!");
                return;
            }

            SaveAccount1(username, password);
            MessageBox.Show("Đã lưu tài khoản thành công!");
        }

        private void button4_Click(object sender, EventArgs e)
        {
            this.Invoke((MethodInvoker)delegate
            {
                richTextBox1.Clear();
                richTextBox2.Clear();
            });
        }

        private void button5_Click(object sender, EventArgs e)
        {

            SaveDataAll();
        }
        private void SaveDataAll()
        {
            bool flag = originalData.Count == 0;
            if (flag)
            {
                MessageBox.Show("Chưa có dữ liệu để lưu! Vui lòng load dữ liệu trước.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            else
            {
                string line = "";
                bool @checked = this.radioButtonLine1.Checked;
                if (@checked)
                {
                    line = "1";
                }
                else
                {
                    bool checked2 = this.radioButtonLine2.Checked;
                    if (checked2)
                    {
                        line = "2";
                    }
                    else
                    {
                        bool checked3 = this.radioButtonAll.Checked;
                        if (checked3)
                        {
                            line = "3";
                        }else if(radioButtonQH.Checked)
                        {
                            line = "5";
                        }
                        else
                        {
                            line = "4";
                        }
                    }
                }
                bool flag2 = this.reminderManager.SaveReminderData(this.dataGridView1, line, true);
                if (flag2)
                {
                    base.Invoke(new MethodInvoker(delegate ()
                    {
                        this.richTextBox1.AppendText("Lưu thông tin nhắc nợ cho Line " + line + " Thành Công!!!\r\n");
                    }));
                }
                else
                {
                    base.Invoke(new MethodInvoker(delegate ()
                    {
                        this.richTextBox1.AppendText("Không lưu được dữ liệu Line " + line + "\r\n");
                    }));
                }
            }
        }

        public bool DeleteOldJsonFile()
        {
            bool result;
            try
            {
                bool flag = System.IO.File.Exists(this.filePath);
                if (flag)
                {
                    System.IO.File.Delete(this.filePath);
                    result = true;
                }
                else
                {
                    result = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi xóa file cũ: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                result = false;
            }
            return result;
        }
        private void button6_Click(object sender, EventArgs e)
        {
            IWebDriver chromeDriver = null;
            ChromeDriverService chromeDriverService = null;

            try
            {
                int pro5 = 1;
                string userIcloud = comboBoxUsername.SelectedItem?.ToString();
                string passIcloud = textBoxPassIcloud.Text?.Trim();
                //richTextBox1.AppendText($"Bắt đầu xử lý yêu cầu cho {deviceInfo}");
                this.Invoke((MethodInvoker)delegate
                {
                    richTextBox1.AppendText($"Bắt đầu kiểm tra máy online hay ofline \r\n");
                });
                chromeDriverService = ChromeDriverService.CreateDefaultService();
                chromeDriverService.HideCommandPromptWindow = true;

                // Đường dẫn tới thư mục chứa Chrome Portable
                string chromePortableDirectory = Path.Combine(Directory.GetCurrentDirectory(), "GoogleChromePortable", "App", "Chrome-bin");

                // Tạo profile riêng cho mỗi session
                //string sessionId = DateTime.Now.Ticks.ToString();
                string ProfilePath = Path.Combine(chromePortableDirectory, "User_Data");

                if (!Directory.Exists(ProfilePath))
                {
                    Directory.CreateDirectory(ProfilePath);
                }

                // Tìm port available

                ChromeOptions chromeOptions = new ChromeOptions();
                chromeOptions.BinaryLocation = Path.Combine(chromePortableDirectory, "Chrome.exe");
                chromeOptions.AddArgument($"--user-data-dir={ProfilePath}");
                chromeOptions.AddUserProfilePreference("credentials_enable_service", true);
                chromeOptions.AddUserProfilePreference("profile.password_manager_enabled", true);
                chromeOptions.AddArgument("--mute-audio");
                chromeOptions.AddArgument("--no-first-run");
                chromeOptions.AddArgument("--disable-default-apps");
                chromeOptions.AddArgument("--disable-popup-blocking");
                chromeOptions.AddArgument("--disable-extensions");
                chromeOptions.AddArgument("--disable-plugins");
                chromeOptions.AddArgument("--disable-images");
                chromeOptions.AddArgument("--remote-allow-origins=*");

                // Thêm arguments để tối ưu performance
                chromeOptions.AddArgument("--disable-gpu");
                chromeOptions.AddArgument("--no-sandbox");
                chromeOptions.AddArgument("--disable-dev-shm-usage");

                #region calc position for profile
                {
                    // calc size
                    Screen[] screens = Screen.AllScreens;
                    Rectangle secondScreenBounds = screens[0].Bounds;
                    int max_width = secondScreenBounds.Width;
                    int max_height = secondScreenBounds.Height;

                    int width = ConfigInfo.chrome_width;
                    int height = ConfigInfo.chrome_height;
                    chromeOptions.AddArgument($"--window-size={width},{height}");

                    // calc max position for pro5
                    int distance_x = ConfigInfo.chrome_distance_x;
                    int distance_y = ConfigInfo.chrome_distance_y;

                    // Tính số cột và hàng tối đa có thể hiển thị
                    int max_column = (max_width - width) / distance_x + 1;
                    int max_row = (max_height - height) / distance_y + 1;

                    // Đảm bảo có ít nhất 1 cột và 1 hàng
                    max_column = Math.Max(1, max_column);
                    max_row = Math.Max(1, max_row);

                    // Tính vị trí dựa trên pro5 (bắt đầu từ 1)
                    int adjustedPosition = pro5 - 1; // Chuyển về base 0 để tính toán
                    int column = (adjustedPosition % max_column) + 1; // Cột từ 1 đến max_column
                    int row = (adjustedPosition / max_column) + 1;    // Hàng từ 1 trở lên

                    // Nếu vượt quá màn hình, wrap lại
                    if (row > max_row)
                    {
                        row = ((row - 1) % max_row) + 1;
                    }

                    // Tính toán vị trí pixel
                    int margin_width_position = (column - 1) * distance_x;
                    int margin_height_position = (row - 1) * distance_y;

                    // Đảm bảo không vượt quá biên màn hình
                    margin_width_position = Math.Min(margin_width_position, max_width - width);
                    margin_height_position = Math.Min(margin_height_position, max_height - height);

                    string position = $"--window-position={margin_width_position},{margin_height_position}";
                    chromeOptions.AddArgument(position);

                }
                #endregion


                chromeDriver = new ChromeDriver(chromeDriverService, chromeOptions);

                // Thêm driver vào danh sách quản lý
                lock (activeDrivers)
                {
                    activeDrivers.Add(chromeDriver);
                }


                // Thiết lập timeout ngắn hơn
                chromeDriver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(30);
                chromeDriver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);

                // Mở iCloud Find My
                chromeDriver.Navigate().GoToUrl("https://www.icloud.com/find/");
                Thread.Sleep(3000);
                this.Invoke((MethodInvoker)delegate
                {
                    richTextBox1.AppendText($"Đang tiến hành login tài khoản icloud: {userIcloud}... \r\n");
                });
                // Click sign in button
                var wait = new WebDriverWait(chromeDriver, TimeSpan.FromSeconds(20));
                try
                {
                    var signInBtn = wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("//ui-button[contains(@class, 'sign-in-button')]")));
                    signInBtn.Click();
                    Thread.Sleep(3000);
                }
                catch (WebDriverTimeoutException)
                {
                    throw new Exception("Không tìm thấy nút Sign In");
                }

                // Switch to iframe and login
                try
                {
                    wait.Until(ExpectedConditions.FrameToBeAvailableAndSwitchToIt("aid-auth-widget"));
                }
                catch (WebDriverTimeoutException)
                {
                    throw new Exception("Không tìm thấy iframe đăng nhập");
                }

                // Nhập email
                var emailField = wait.Until(ExpectedConditions.ElementToBeClickable(By.Id("account_name_text_field")));
                emailField.Clear();
                emailField.SendKeys(userIcloud);
                Thread.Sleep(1000);

                // Click checkbox và next
                var nextPass = wait.Until(ExpectedConditions.ElementToBeClickable(By.ClassName("form-checkbox-indicator")));
                nextPass.Click();
                Thread.Sleep(1000);

                var singin = wait.Until(ExpectedConditions.ElementToBeClickable(By.Id("sign-in")));
                singin.Click();
                Thread.Sleep(2000);

                // Nhập password
                var passField = wait.Until(ExpectedConditions.ElementToBeClickable(By.Id("password_text_field")));
                passField.Clear();
                passField.SendKeys(passIcloud);
                Thread.Sleep(1000);

                // Click sign in
                IWebElement signInButton = chromeDriver.FindElement(By.Id("sign-in"));
                bool isDisabled = signInButton.GetAttribute("disabled") != null;

                if (isDisabled)
                {
                    IJavaScriptExecutor js = (IJavaScriptExecutor)chromeDriver;
                    js.ExecuteScript("arguments[0].removeAttribute('disabled');", signInButton);
                    js.ExecuteScript("arguments[0].click();", signInButton);
                }
                else
                {
                    signInButton.Click();
                }
                // Đợi một chút để form xử lý
                Thread.Sleep(3000);

                // Kiểm tra xem có lỗi đăng nhập không
                try
                {
                    // Tìm element báo lỗi
                    var errorElement = chromeDriver.FindElement(By.CssSelector(".form-cell-wrapper.is-error, #invalid_user_name_pwd_err_msg[aria-hidden='false'], .si-error-message"));
                    if (errorElement != null && errorElement.Displayed)
                    {
                        // Có lỗi đăng nhập
                        string errorMessage = "Mật khẩu hoặc tài khoản không đúng";
                        try
                        {
                            var errorText = chromeDriver.FindElement(By.Id("invalid_user_name_pwd_err_msg"));
                            if (!string.IsNullOrEmpty(errorText.Text))
                            {
                                errorMessage = errorText.Text;
                            }
                        }
                        catch { }

                        this.Invoke((MethodInvoker)delegate
                        {
                            richTextBox1.AppendText($"[{DateTime.Now:HH:mm:ss}] ĐĂNG NHẬP THẤT BẠI: {errorMessage}\r\n");
                            richTextBox1.ScrollToCaret();
                        });

                        throw new Exception($"Đăng nhập thất bại cho tài khoản {userIcloud}: {errorMessage}");
                    }
                }
                catch (NoSuchElementException)
                {
                    // Không tìm thấy element lỗi có nghĩa là đăng nhập thành công
                }

                // Kiểm tra xem có chuyển sang trang chính không (đăng nhập thành công)
                try
                {
                    // Đợi một chút để trang load
                    Thread.Sleep(2000);

                    // Kiểm tra xem có còn trong iframe đăng nhập không
                    var currentUrl = chromeDriver.Url;
                    if (currentUrl.Contains("idmsa.apple.com") || currentUrl.Contains("signin"))
                    {
                        // Vẫn còn trong trang đăng nhập, có thể là lỗi
                        throw new Exception("Vẫn còn trong trang đăng nhập, có thể đăng nhập thất bại");
                    }
                }
                catch (Exception ex)
                {
                    // Ghi log lỗi nhưng vẫn thử tiếp tục
                    Console.WriteLine($"Cảnh báo: {ex.Message}");
                }
                // Switch back to main content
                chromeDriver.SwitchTo().DefaultContent();
                this.Invoke((MethodInvoker)delegate
                {
                    richTextBox1.AppendText($"Đã đăng nhập thành công tài khoản icloud {userIcloud}, đợi trang Find My load...\r\n");
                });
                // Đợi trang Find My load
                Thread.Sleep(10000);

                // Tìm và switch vào iframe Find My
                wait = new WebDriverWait(chromeDriver, TimeSpan.FromSeconds(20));
                IWebElement iframe = null;

                string[] iframeSelectors = {
            "iframe.child-application",
            "iframe[src*='find']",
            "iframe[title*='Find']",
            ".child-application"
        };

                foreach (string selector in iframeSelectors)
                {
                    try
                    {
                        iframe = wait.Until(ExpectedConditions.ElementToBeClickable(By.CssSelector(selector)));
                        if (iframe != null)
                        {

                            //richTextBox1.AppendText($"Tìm thấy iframe với selector: {selector}");
                            break;
                        }
                    }
                    catch (WebDriverTimeoutException)
                    {
                        continue;
                    }
                }

                if (iframe == null)
                {
                    throw new Exception("Không thể tìm thấy iframe Find My");
                }

                chromeDriver.SwitchTo().Frame(iframe);
                //richTextBox1.AppendText("Đã switch vào iframe Find My");

                // Đợi device list load
                Thread.Sleep(5000);

                try
                {
                    wait.Until(driver => driver.FindElements(By.CssSelector(".fmip-device-list-item")).Count > 0);
                }
                catch (WebDriverTimeoutException)
                {
                    throw new Exception("Không thể load danh sách thiết bị");
                }

                // Tìm và click vào iPhone
                var deviceElements = chromeDriver.FindElements(By.CssSelector(".fmip-device-list-item"));
                this.Invoke((MethodInvoker)delegate
                {
                    richTextBox1.AppendText($"Tìm thấy {deviceElements.Count} thiết bị trên tài khoản {userIcloud}\r\n");
                });
                bool iPhoneFound = false;
                foreach (var device in deviceElements)
                {
                    try
                    {
                        var nameElement = device.FindElement(By.CssSelector("[data-testid='show-device-name'], .device-name, .name"));
                        var deviceName = nameElement.Text.Trim();
                        string number = Regex.Match(deviceName, @"\d+").Value;
                        // Lấy link hình ảnh
                        var imageElement = device.FindElement(By.CssSelector("img.image"));
                        var imageSrc = imageElement.GetAttribute("src");

                        // Kiểm tra trạng thái dựa vào tên file ảnh
                        bool isOnline = imageSrc.Contains("online-sourcelist.png");
                        bool isOffline = imageSrc.Contains("offline-sourcelist.png");

                        string deviceStatus = "";
                        if (isOnline)
                        {
                            deviceStatus = "Online";
                        }
                        else if (isOffline)
                        {
                            deviceStatus = "Offline";
                        }
                        else
                        {
                            deviceStatus = "Unknown";
                        }
                    }
                    catch (Exception deviceEx)
                    {
                        this.Invoke((MethodInvoker)delegate
                        {
                            richTextBox1.AppendText($"Lỗi khi xử lý thiết bị: {deviceEx.Message}\r\n");
                        });
                        continue;
                    }
                }

            }
            catch (Exception ex)
            {
                this.Invoke((MethodInvoker)delegate
                {
                    richTextBox1.AppendText($"Lỗi: {ex.Message}\r\n");
                });
                this.Invoke((MethodInvoker)delegate
                {
                    richTextBox1.AppendText($"Lỗi khi xử lý : {ex.Message}");
                    //MessageBox.Show($"Lỗi khi xử lý : {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                });
            }
            finally
            {
                // Đóng browser và cleanup
                try
                {
                    if (chromeDriver != null)
                    {
                        lock (activeDrivers)
                        {
                            activeDrivers.Remove(chromeDriver);
                        }
                        chromeDriver.Quit();
                        this.Invoke((MethodInvoker)delegate
                        {
                            richTextBox1.AppendText($"Đã đóng Chrome instance cho thiết bị \r\n");
                        });
                    }

                    if (chromeDriverService != null)
                    {
                        chromeDriverService.Dispose();
                    }
                }
                catch (Exception cleanupEx)
                {
                    this.Invoke((MethodInvoker)delegate
                    {
                        richTextBox1.AppendText($"Lỗi khi cleanup: {cleanupEx.Message} \r\n");
                    });
                }
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            string userIcloud = comboBoxUsername1.SelectedItem?.ToString();
            string passIcloud = textBoxPass.Text?.Trim();

            if (string.IsNullOrEmpty(userIcloud) || string.IsNullOrEmpty(passIcloud))
            {
                MessageBox.Show("Vui lòng nhập đầy đủ thông tin acount 2Gold", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            List<string> dsKH = GetUserTomorow();
            LoadAndSaveData(dsKH);
        }

        private void button8_Click(object sender, EventArgs e)
        {
            originalIndexMap.Clear();
            originalData.Clear();
            originalDataNew.Clear();
            dataGridView1.Rows.Clear();
            textBoxSearch.Text = "";
            richTextBox1.Clear();
            richTextBox2.Clear();
            this.button5.Invoke(new MethodInvoker(delegate ()
            {
                this.button5.Enabled = true;
            }));
        }

        private async void button9_Click(object sender, EventArgs e)
        {
            try
            {
                string filePath = "duLieuGoc.xlsx";
                Dictionary<string, int> mergedData = new Dictionary<string, int>();
                List<string> info = new List<string>();
                string user = comboBoxUsername1.Text.Trim();
                string pass = textBoxPass.Text;
                HttpRequest request = new HttpRequest();
                request.UserAgent = xNet.Http.ChromeUserAgent();
                string cookie = request.Get("https://2gold.biz/").Cookies.ToString();
                request.Cookies = new CookieDictionary();
                // gán cookie vào http
                for (int c = 0; c < cookie.Split(';').Length; c++)
                {
                    try
                    {
                        string name = cookie.Split(';')[c].Split('=')[0].Trim();
                        string value = cookie.Split(';')[c].Substring(cookie.Split(';')[c].IndexOf('=') + 1).Trim();
                        if (request.Cookies.ContainsKey(name))
                            request.Cookies.Remove(name);

                        request.Cookies.Add(name, value);
                    }
                    catch (Exception ex) { }
                }
                request.AddHeader("accept", "*/*");
                request.AddHeader("accept-language", "en-US,en;q=0.9,vi;q=0.8");
                request.AddHeader("origin", "https://2gold.biz");
                request.AddHeader("priority", "u=1, i");
                request.AddHeader("referer", "https://2gold.biz/User/Login");
                request.AddHeader("sec-ch-ua", "\"Google Chrome\";v=\"137\", \"Chromium\";v=\"137\", \"Not/A)Brand\";v=\"24\"");
                request.AddHeader("sec-ch-ua-mobile", "?1");
                request.AddHeader("sec-ch-ua-platform", "\"Android\"");
                request.AddHeader("sec-fetch-dest", "empty");
                request.AddHeader("sec-fetch-mode", "cors");
                request.AddHeader("sec-fetch-site", "same-origin");
                request.AddHeader("x-kl-ajax-request", "Ajax_Request");
                request.AddHeader("x-requested-with", "XMLHttpRequest");
                var body = "Username=" + user + "&Password=" + pass;
                var response3 = request.Post("https://2gold.biz/User/ProcessLogin", body, "application/x-www-form-urlencoded; charset=UTF-8");
                string response = response3.ToString(); 
                string ck = response3.Cookies.ToString().Split(';')[1].Trim();
                JObject json = JObject.Parse(response);
                if (json["Message"].ToString().Contains("Đăng nhập thành công"))
                {
                    string UserID = json["Data"]["UserID"].ToString();
                    string ShopID = json["Data"]["ShopID"].ToString();
                    string Token = json["Data"]["Token"].ToString();
                    request.AddHeader("accept", "*/*");
                    request.AddHeader("accept-language", "en-US,en;q=0.9,vi;q=0.8");
                    request.AddHeader("origin", "https://2gold.biz");
                    request.AddHeader("priority", "u=1, i");
                    request.AddHeader("referer", "https://2gold.biz/");
                    request.AddHeader("sec-ch-ua", "\"Google Chrome\";v=\"137\", \"Chromium\";v=\"137\", \"Not/A)Brand\";v=\"24\"");
                    request.AddHeader("sec-ch-ua-mobile", "?0");
                    request.AddHeader("sec-ch-ua-platform", "\"Windows\"");
                    request.AddHeader("sec-fetch-dest", "empty");
                    request.AddHeader("sec-fetch-mode", "cors");
                    request.AddHeader("sec-fetch-site", "same-site");
                    //var body2 = "UserID=" + UserID + "&ShopID=" + ShopID + "&Token=" + Token;
                    string response2 = request.Get("https://api.2gold.biz/api/Staff/AllStaffActive?UserID="+UserID+"&ShopID="+ShopID+"&Token="+Token).ToString();
                    json = JObject.Parse(response2);
                    if (json["Message"].ToString().Contains("Lấy dữ liệu nhân viên thành công"))
                    {
                        request.AddHeader("accept", "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7");
                        request.AddHeader("accept-language", "en-US,en;q=0.9,vi;q=0.8");
                        request.AddHeader("priority", "u=0, i");
                        request.AddHeader("referer", "https://2gold.biz/Calendar/Installment");
                        request.AddHeader("sec-ch-ua", "\"Google Chrome\";v=\"137\", \"Chromium\";v=\"137\", \"Not/A)Brand\";v=\"24\"");
                        request.AddHeader("sec-ch-ua-mobile", "?0");
                        request.AddHeader("sec-ch-ua-platform", "\"Windows\"");
                        request.AddHeader("sec-fetch-dest", "document");
                        request.AddHeader("sec-fetch-mode", "navigate");
                        request.AddHeader("sec-fetch-site", "same-origin");
                        request.AddHeader("sec-fetch-user", "?1");
                        request.AddHeader("upgrade-insecure-requests", "1");
                        request.AddHeader("Cookie", ck + "; cmd_username="+user+"; cmd_password="+pass);
                        var response4 = request.Get("https://2gold.biz/Excel/ExportExcelNotifyPawnNew?TypeProduct=2&GeneralSearch=&Status=1000&StaffId=0");
                        byte[] fileBytes = response4.ToBytes();
                        await Task.Run(() => System.IO.File.WriteAllBytes(filePath, fileBytes));

                    }
                    


                }
                else
                {
                    throw new Exception("Đăng Nhập Thất Bại");
                }
                
            }
            catch (Exception ex)
            {
                
            }
        }
        private async Task<Dictionary<string, int>> GetFileExcel()
        {
            try
            {
                Dictionary<string, int> mergedData = new Dictionary<string, int>();
                List<string> info = new List<string>();
                string user = comboBoxUsername1.Text.Trim();
                string pass = textBoxPass.Text;
                HttpRequest request = new HttpRequest();
                request.UserAgent = xNet.Http.ChromeUserAgent();
                string cookie = request.Get("https://2gold.biz/").Cookies.ToString();
                request.Cookies = new CookieDictionary();
                // gán cookie vào http
                for (int c = 0; c < cookie.Split(';').Length; c++)
                {
                    try
                    {
                        string name = cookie.Split(';')[c].Split('=')[0].Trim();
                        string value = cookie.Split(';')[c].Substring(cookie.Split(';')[c].IndexOf('=') + 1).Trim();
                        if (request.Cookies.ContainsKey(name))
                            request.Cookies.Remove(name);

                        request.Cookies.Add(name, value);
                    }
                    catch (Exception ex) { }
                }
                request.AddHeader("accept", "*/*");
                request.AddHeader("accept-language", "en-US,en;q=0.9,vi;q=0.8");
                request.AddHeader("origin", "https://2gold.biz");
                request.AddHeader("priority", "u=1, i");
                request.AddHeader("referer", "https://2gold.biz/User/Login");
                request.AddHeader("sec-ch-ua", "\"Google Chrome\";v=\"137\", \"Chromium\";v=\"137\", \"Not/A)Brand\";v=\"24\"");
                request.AddHeader("sec-ch-ua-mobile", "?1");
                request.AddHeader("sec-ch-ua-platform", "\"Android\"");
                request.AddHeader("sec-fetch-dest", "empty");
                request.AddHeader("sec-fetch-mode", "cors");
                request.AddHeader("sec-fetch-site", "same-origin");
                request.AddHeader("x-kl-ajax-request", "Ajax_Request");
                request.AddHeader("x-requested-with", "XMLHttpRequest");
                var body = "Username=" + user + "&Password=" + pass;
                string response = request.Post("https://2gold.biz/User/ProcessLogin", body, "application/x-www-form-urlencoded; charset=UTF-8").ToString();
                JObject json = JObject.Parse(response);
                if (json["Message"].ToString().Contains("Đăng nhập thành công"))
                {
                    string UserID = json["Data"]["UserID"].ToString();
                    string ShopID = json["Data"]["ShopID"].ToString();
                    string Token = json["Data"]["Token"].ToString();
                    request.AddHeader("accept", "*/*");
                    request.AddHeader("accept-language", "en-US,en;q=0.9,vi;q=0.8");
                    request.AddHeader("origin", "https://2gold.biz");
                    request.AddHeader("priority", "u=1, i");
                    request.AddHeader("referer", "https://2gold.biz/");
                    request.AddHeader("sec-ch-ua", "\"Google Chrome\";v=\"137\", \"Chromium\";v=\"137\", \"Not/A)Brand\";v=\"24\"");
                    request.AddHeader("sec-ch-ua-mobile", "?0");
                    request.AddHeader("sec-ch-ua-platform", "\"Windows\"");
                    request.AddHeader("sec-fetch-dest", "empty");
                    request.AddHeader("sec-fetch-mode", "cors");
                    request.AddHeader("sec-fetch-site", "same-site");
                    var body2 = "UserID=" + UserID + "&ShopID=" + ShopID + "&Token=" + Token;
                    string response2 = request.Post("https://api.2gold.biz/api/PaymentNotify/GetNotifyCount", body2, "application/x-www-form-urlencoded; charset=UTF-8").ToString();
                    json = JObject.Parse(response2);
                    if (json["Message"].ToString().Contains("Lấy số liệu nhắc nợ thành công"))
                    {
                        int InstallmentTotal = int.Parse(json["Data"]["InstallmentTotal"].ToString());
                        double result = InstallmentTotal / 50.0;
                        int pageCount = (int)Math.Ceiling(result);
                        string firtNum = "";
                        string secondNum = "";
                        for (int j = 0; j < pageCount; j++)
                        {

                            if (j == 0)
                            {
                                firtNum = "1";
                            }
                            else
                            {
                                firtNum = (j + 1).ToString();
                                secondNum = pageCount.ToString();
                            }
                            request.AddHeader("accept", "application/json, text/javascript, */*; q=0.01");
                            request.AddHeader("accept-language", "en-US,en;q=0.9,vi;q=0.8");
                            request.AddHeader("origin", "https://2gold.biz");
                            request.AddHeader("priority", "u=1, i");
                            request.AddHeader("referer", "https://2gold.biz/");
                            request.AddHeader("sec-ch-ua", "\"Google Chrome\";v=\"137\", \"Chromium\";v=\"137\", \"Not/A)Brand\";v=\"24\"");
                            request.AddHeader("sec-ch-ua-mobile", "?0");
                            request.AddHeader("sec-ch-ua-platform", "\"Windows\"");
                            request.AddHeader("sec-fetch-dest", "empty");
                            request.AddHeader("sec-fetch-mode", "cors");
                            request.AddHeader("sec-fetch-site", "same-site");
                            var body1 = "datatable%5Bpagination%5D%5Bpage%5D=" + firtNum + "&datatable%5Bpagination%5D%5Bpages%5D=" + secondNum + "&datatable%5Bpagination%5D%5Bperpage%5D=" + InstallmentTotal + "&datatable%5Bsort%5D%5Bsort%5D=&datatable%5Bsort%5D%5Bfield%5D=&datatable%5Bquery%5D%5BgeneralSearch%5D=&datatable%5Bquery%5D%5BStatus%5D=1000&datatable%5Bquery%5D%5BShopID%5D=" + ShopID + "&datatable%5Bquery%5D%5BTypeProduct%5D=2&datatable%5Bquery%5D%5BUserID%5D=" + UserID + "&datatable%5Bquery%5D%5BToken%5D=" + Token + "&datatable%5Bquery%5D%5BStaffId%5D=0&datatable%5Bquery%5D%5BPageCurrent%5D=1&datatable%5Bquery%5D%5BPerPageCurrent%5D=&datatable%5Bquery%5D%5BcolumnCurrent%5D=&datatable%5Bquery%5D%5BsortCurrent%5D=";
                            string response1 = request.Post("https://api.2gold.biz/api/PaymentNotify/List", body1, "application/x-www-form-urlencoded; charset=UTF-8").ToString();
                            json = JObject.Parse(response1);
                            if (json["data"].Count() > 0)
                            {
                                int countData = json["data"].Count();
                                for (int i = 0; i < countData; i++)
                                {
                                    string maHD = json["data"][i]["StrCodeID"].ToString();
                                    string tenKH = json["data"][i]["CustomerName"].ToString();
                                    string diachi = json["data"][i]["CustomerAddress"].ToString();
                                    string noCu = json["data"][i]["DebitMoney"].ToString();
                                    string tiencandong = json["data"][i]["PaymentNotify"].ToString();
                                    string lydo = json["data"][i]["Note"].ToString();
                                    info.Add(maHD + "|" + tenKH + "|" + tiencandong + "|" + lydo);
                                }

                            }
                        }


                    }



                }
                else
                {
                    throw new Exception("Đăng Nhập Thất Bại");
                }
                return mergedData;
            }
            catch (Exception ex)
            {
                return null;
            } 
            
        }

        private void button10_Click(object sender, EventArgs e)
        {
            string userIcloud = comboBoxUsername1.SelectedItem?.ToString();
            string passIcloud = textBoxPass.Text?.Trim();

            if (string.IsNullOrEmpty(userIcloud) || string.IsNullOrEmpty(passIcloud))
            {
                MessageBox.Show("Vui lòng nhập đầy đủ thông tin acount 2Gold", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            List<string> dsKH = GetUserTomorow2();
            LoadAndSaveData(dsKH);
        }

        private void button11_Click(object sender, EventArgs e)
        {
            this.button5.Invoke(new MethodInvoker(delegate ()
            {
                this.button5.Enabled = false;
            }));
            string searchKey = this.textBoxSearch.Text.Trim();
            bool flag = string.IsNullOrEmpty(searchKey);
            if (flag)
            {
                MessageBox.Show("Vui lòng nhập mã đơn hàng cần tìm!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            else
            {
                bool flag2 = originalData.Count == 0;
                if (flag2)
                {
                    MessageBox.Show("Chưa có dữ liệu để tìm kiếm! Vui lòng load dữ liệu trước.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
                else
                {
                    List<string> list = originalData.Where(delegate (string item)
                    {
                        string[] array = item.Split(new char[]
                        {
                    '|'
                        });
                        bool flag4 = array.Length >= 2;
                        bool result;
                        if (flag4)
                        {
                            string text = array[0].ToLower().Trim();
                            string text2 = array[1].ToLower().Trim();
                            string searchInput = array[8].ToLower().Trim();
                            string value = Regex.Match(text2, "\\d+").Value;
                            string value2 = searchKey.ToLower();
                            string value3 = NormalizeDateSearch(searchKey.ToLower());
                            string value4 = searchKey.ToLower().Replace("-", "").Trim();
                            string value5 = searchKey.ToLower().Replace("bh-", "").Trim();
                            string text3 = text.Replace("-", "").ToLower().Trim();
                            string text4 = text.Replace("bh-", "").ToLower().Trim();
                            string text5 =  NormalizeDateSearch(searchInput);
                            result = (text.Equals(value2, StringComparison.OrdinalIgnoreCase) || text2.Equals(value2, StringComparison.OrdinalIgnoreCase) || value.Equals(value2, StringComparison.OrdinalIgnoreCase) || text5.Equals(value3, StringComparison.OrdinalIgnoreCase) || text3.Equals(value4, StringComparison.OrdinalIgnoreCase) || text4.Equals(value5, StringComparison.OrdinalIgnoreCase));
                        }
                        else
                        {
                            result = false;
                        }
                        return result;
                    }).ToList<string>();
                    bool flag3 = list.Count == 0;
                    if (flag3)
                    {
                        MessageBox.Show("Không tìm thấy đơn hàng có mã: " + searchKey, "Kết quả tìm kiếm", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                    }
                    else
                    {
                        this.DisplayData(list);
                    }
                }
            }

        }
        public static string NormalizeDateSearch(string searchInput)
        {
            bool flag = string.IsNullOrWhiteSpace(searchInput);
            string result;
            if (flag)
            {
                result = searchInput;
            }
            else
            {
                string s = searchInput.Trim();
                string[] array = new string[]
                {
                    "d/M/yyyy",
                    "dd/M/yyyy",
                    "d/MM/yyyy",
                    "dd/MM/yyyy",
                    "d-M-yyyy",
                    "dd-M-yyyy",
                    "d-MM-yyyy",
                    "dd-MM-yyyy",
                    "d/M/yy",
                    "dd/M/yy",
                    "d/MM/yy",
                    "dd/MM/yy",
                    "d/M",
                    "dd/M",
                    "d/MM",
                    "dd/MM",
                    "d-M",
                    "dd-M",
                    "d-MM",
                    "dd-MM"
                };
                DateTime dateTime;
                foreach (string text in array)
                {
                    bool flag2 = DateTime.TryParseExact(s, text, CultureInfo.InvariantCulture, DateTimeStyles.None, out dateTime);
                    if (flag2)
                    {
                        bool flag3 = !text.Contains("y");
                        if (flag3)
                        {
                            dateTime = new DateTime(DateTime.Now.Year, dateTime.Month, dateTime.Day);
                        }
                        else
                        {
                            bool flag4 = text.Contains("yy") && !text.Contains("yyy");
                            if (flag4)
                            {
                                bool flag5 = dateTime.Year < 100;
                                if (flag5)
                                {
                                    dateTime = ((dateTime.Year <= 30) ? new DateTime(2000 + dateTime.Year, dateTime.Month, dateTime.Day) : new DateTime(1900 + dateTime.Year, dateTime.Month, dateTime.Day));
                                }
                            }
                        }
                        return dateTime.ToString("dd/MM/yyyy");
                    }
                }
                bool flag6 = DateTime.TryParse(s, out dateTime);
                if (flag6)
                {
                    result = dateTime.ToString("dd/MM/yyyy");
                }
                else
                {
                    result = searchInput;
                }
            }
            return result;
        }
        private void button12_Click(object sender, EventArgs e)
        {
            this.button5.Invoke(new MethodInvoker(delegate ()
            {
                this.button5.Enabled = true;
            }));
            if (originalData.Count == 0)
            {
                MessageBox.Show("Chưa có dữ liệu gốc để refresh!", "Thông báo",
                               MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Xóa nội dung tìm kiếm
            textBoxSearch.Text = "";

            // Hiển thị lại dữ liệu gốc
            DisplayData(originalData);
            
        }

        private void button14_Click(object sender, EventArgs e)
        {
            Thread t = new Thread(async () =>
            {
                bRunning = true;
                button14.Invoke((MethodInvoker)delegate ()
                {
                    button14.Enabled = false;
                });

                int soluong_toi_da = 79;
                SemaphoreSlim semaphore = new SemaphoreSlim(soluong_toi_da, soluong_toi_da);
                List<Task> tasks = new List<Task>();

                // Lấy giá trị start và end từ numericUpDown
                int startIndex = 0;
                int endIndex = 0;
                bool useFullRange = false;

                // Invoke để lấy giá trị từ UI thread
                numericUpDownStart.Invoke((MethodInvoker)delegate ()
                {
                    startIndex = (int)numericUpDownStart.Value;
                });

                numericUpDownEnd.Invoke((MethodInvoker)delegate ()
                {
                    endIndex = (int)numericUpDownEnd.Value;
                });

                // Kiểm tra nếu cả 2 giá trị đều là 0 (không nhập gì)
                if (startIndex == 0 && endIndex == 0)
                {
                    var result = MessageBox.Show("Bạn chưa nhập giá trị Start và End.\n\nChọn YES để chạy tất cả các dòng\nChọn NO để dừng lại",
                                               "Xác nhận",
                                               MessageBoxButtons.YesNo,
                                               MessageBoxIcon.Question);

                    if (result == DialogResult.Yes)
                    {
                        useFullRange = true;
                        startIndex = 0;
                        endIndex = dataGridView1.Rows.Count - 1;
                    }
                    else
                    {
                        // Người dùng chọn NO, dừng lại
                        button14.Invoke((MethodInvoker)delegate ()
                        {
                            button14.Enabled = true;
                        });
                        return;
                    }
                }
                else
                {
                    // Kiểm tra giá trị hợp lệ
                    if (startIndex < 0) startIndex = 0;
                    if (endIndex >= dataGridView1.Rows.Count) endIndex = dataGridView1.Rows.Count - 1;
                    if (startIndex > endIndex)
                    {
                        // Swap nếu start > end
                        int temp = startIndex;
                        startIndex = endIndex;
                        endIndex = temp;
                    }
                }

                // Chạy từ startIndex đến endIndex
                for (int i = startIndex; i <= endIndex; i++)
                {
                    await semaphore.WaitAsync(); // Chờ có slot trống
                    int currentIndex = i;
                    Task task = Task.Run(async () =>
                    {
                        try
                        {
                            if (bRunning)
                            {
                               runToDie(currentIndex); // QUAN TRỌNG: Sử dụng await
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Error in task {currentIndex}: {ex.Message}");
                        }
                        finally
                        {
                            semaphore.Release(); // Giải phóng slot AFTER task thực sự hoàn thành
                        }
                    });
                    tasks.Add(task);
                    Thread.Sleep(random.Next(60000, 70000)); // delay giữa các task
                }

                await Task.WhenAll(tasks); // Chờ tất cả tasks hoàn thành
                semaphore.Dispose();

                button14.Invoke((MethodInvoker)delegate ()
                {
                    button14.Enabled = true;
                });
            });
            t.Start();
        }

        private void button13_Click(object sender, EventArgs e)
        {
            Form1.shouldMerge = false;
            this.LoadDataByMonGoDB();
            
        }

        private void LoadDataByMonGoDB()
        {
            this.dataGridView1.AllowUserToAddRows = true;
            try
            {
                string text = "";
                string selectedValue = comboBox1.SelectedItem?.ToString();
                bool @checked = this.radioButtonLine1.Checked;
                if (@checked && selectedValue == "None")
                {
                    text = GetDataFromMongoDB1();
                }
                else
                {
                    bool checked2 = this.radioButtonLine2.Checked;
                    if (checked2 && selectedValue == "None")
                    {
                        text = GetDataFromMongoDB2();
                    }
                    else
                    {
                        bool checked3 = radioButtonAll.Checked;
                        if (checked3 && selectedValue == "None")
                        {
                            text = GetDataFromMongoDBEveryDay();
                        }
                        else if (radioButtonQH.Checked && selectedValue == "None")
                        {
                            text = GetDataFromMongoDBByQH();
                        }
                        else
                        {

                        }
                    }
                }
                bool flag = string.IsNullOrWhiteSpace(text) || text == "[]" || text == "{}";
                if (flag)
                {
                    this.LoadEmptyDataGridView();
                    MessageBox.Show("Dữ liệu trống trong MongoDB. Đã tạo bảng trống để nhập dữ liệu.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                }
                else
                {
                    Dictionary<string, Form1.DebtReminderManager.DebtReminderInfo> dictionary = JsonConvert.DeserializeObject<Dictionary<string, Form1.DebtReminderManager.DebtReminderInfo>>(text);
                    bool flag2 = dictionary == null || dictionary.Count == 0;
                    if (flag2)
                    {
                        this.LoadEmptyDataGridView();
                        MessageBox.Show("Không có dữ liệu. Đã tạo bảng trống để nhập dữ liệu.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                    }
                    else
                    {
                        this.LoadDataToGridView(dictionary);
                        this.originalDataNew = dictionary;
                        originalData = new List<string>();
                        int num = 0;
                        foreach (KeyValuePair<string, Form1.DebtReminderManager.DebtReminderInfo> keyValuePair in dictionary)
                        {
                            num++;
                            string item = string.Format("{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}|{8}", new object[]
                            {
                                keyValuePair.Key,
                                keyValuePair.Value.Note,
                                keyValuePair.Value.userIcloud,
                                keyValuePair.Value.PassIcloud,
                                keyValuePair.Value.NamePhoneChange,
                                num - 1,
                                keyValuePair.Value.Line,
                                keyValuePair.Value.date,
                                keyValuePair.Value.TenSo
                            });
                            originalData.Add(item);
                        }
                        bool checked4 = this.radioButtonAll.Checked;
                        if (checked4)
                        {
                            DialogResult dialogResult = MessageBox.Show("Bạn Muốn Lưu Dữ Liệu vào file Excell ", "LƯU DỮ LIỆU HẰNG NGÀY", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                            bool flag3 = dialogResult != DialogResult.Yes;
                            if (!flag3)
                            {
                                string fileName = string.Format("DataEveryDay{0:yyyyMMdd_HHmmss}.xlsx", DateTime.Now);
                                ExportToExcel(dictionary, fileName);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi xử lý dữ liệu: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                this.LoadEmptyDataGridView();
            }
        }

        private string GetDataDeleteFromMongoDB()
        {
            string result;
            try
            {
                string text = "mongodb+srv://banhmichaothuongnhoicloud:8390813asd@cluster0.srtd7rc.mongodb.net/?retryWrites=true&w=majority&appName=Cluster0";
                MongoClientSettings mongoClientSettings = MongoClientSettings.FromConnectionString(text);
                mongoClientSettings.ServerApi = new ServerApi(ServerApiVersion.V1, default(Optional<bool?>), default(Optional<bool?>));
                MongoClient mongoClient = new MongoClient(mongoClientSettings);
                IMongoDatabase database = mongoClient.GetDatabase("duLieuAPP", null);
                IMongoCollection<Form1.DebtReminderManager.dataRungIcloud> collection = database.GetCollection<Form1.DebtReminderManager.dataRungIcloud>("dataRung", null);
                FilterDefinition<Form1.DebtReminderManager.dataRungIcloud> filterDefinition = Builders<Form1.DebtReminderManager.dataRungIcloud>.Filter.Eq<string>((Form1.DebtReminderManager.dataRungIcloud x) => x.Name, "dulieuXoa");
                Form1.DebtReminderManager.dataRungIcloud dataRungIcloud = IFindFluentExtensions.FirstOrDefault<Form1.DebtReminderManager.dataRungIcloud, Form1.DebtReminderManager.dataRungIcloud>(IMongoCollectionExtensions.Find<Form1.DebtReminderManager.dataRungIcloud>(collection, filterDefinition, null), default(CancellationToken));
                result = (((dataRungIcloud != null) ? dataRungIcloud.Data : null) ?? "");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi kết nối MongoDB: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                result = "";
            }
            return result;
        }

        // Token: 0x06000077 RID: 119 RVA: 0x000060CC File Offset: 0x000042CC
        private string GetDataFromMongoDB()
        {
            string result;
            try
            {
                string text = "mongodb+srv://banhmichaothuongnhoicloud:8390813asd@cluster0.srtd7rc.mongodb.net/?retryWrites=true&w=majority&appName=Cluster0";
                MongoClientSettings mongoClientSettings = MongoClientSettings.FromConnectionString(text);
                mongoClientSettings.ServerApi = new ServerApi(ServerApiVersion.V1, default(Optional<bool?>), default(Optional<bool?>));
                MongoClient mongoClient = new MongoClient(mongoClientSettings);
                IMongoDatabase database = mongoClient.GetDatabase("duLieuAPP", null);
                IMongoCollection<Form1.DebtReminderManager.dataRungIcloud> collection = database.GetCollection<Form1.DebtReminderManager.dataRungIcloud>("dataRung", null);
                FilterDefinition<Form1.DebtReminderManager.dataRungIcloud> filterDefinition = Builders<Form1.DebtReminderManager.dataRungIcloud>.Filter.Eq<string>((Form1.DebtReminderManager.dataRungIcloud x) => x.Name, "dulieurung");
                Form1.DebtReminderManager.dataRungIcloud dataRungIcloud = IFindFluentExtensions.FirstOrDefault<Form1.DebtReminderManager.dataRungIcloud, Form1.DebtReminderManager.dataRungIcloud>(IMongoCollectionExtensions.Find<Form1.DebtReminderManager.dataRungIcloud>(collection, filterDefinition, null), default(CancellationToken));
                result = (((dataRungIcloud != null) ? dataRungIcloud.Data : null) ?? "");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi kết nối MongoDB: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                result = "";
            }
            return result;
        }

        // Token: 0x06000078 RID: 120 RVA: 0x000061F0 File Offset: 0x000043F0
        private string GetDataFromMongoDB1()
        {
            string result;
            try
            {
                string text = "mongodb+srv://banhmichaothuongnhoicloud:8390813asd@cluster0.srtd7rc.mongodb.net/?retryWrites=true&w=majority&appName=Cluster0";
                MongoClientSettings mongoClientSettings = MongoClientSettings.FromConnectionString(text);
                mongoClientSettings.ServerApi = new ServerApi(ServerApiVersion.V1, default(Optional<bool?>), default(Optional<bool?>));
                MongoClient mongoClient = new MongoClient(mongoClientSettings);
                IMongoDatabase database = mongoClient.GetDatabase("duLieuAPP", null);
                IMongoCollection<Form1.DebtReminderManager.dataRungIcloud> collection = database.GetCollection<Form1.DebtReminderManager.dataRungIcloud>("dataRung", null);
                FilterDefinition<Form1.DebtReminderManager.dataRungIcloud> filterDefinition = Builders<Form1.DebtReminderManager.dataRungIcloud>.Filter.Eq<string>((Form1.DebtReminderManager.dataRungIcloud x) => x.Name, "Line1");
                Form1.DebtReminderManager.dataRungIcloud dataRungIcloud = IFindFluentExtensions.FirstOrDefault<Form1.DebtReminderManager.dataRungIcloud, Form1.DebtReminderManager.dataRungIcloud>(IMongoCollectionExtensions.Find<Form1.DebtReminderManager.dataRungIcloud>(collection, filterDefinition, null), default(CancellationToken));
                result = (((dataRungIcloud != null) ? dataRungIcloud.Data : null) ?? "");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi kết nối MongoDB: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                result = "";
            }
            return result;
        }

        // Token: 0x06000079 RID: 121 RVA: 0x00006314 File Offset: 0x00004514
        private string GetDataFromMongoDBEveryDay()
        {
            string result;
            try
            {
                string text = "mongodb+srv://banhmichaothuongnhoicloud:8390813asd@cluster0.srtd7rc.mongodb.net/?retryWrites=true&w=majority&appName=Cluster0";
                MongoClientSettings mongoClientSettings = MongoClientSettings.FromConnectionString(text);
                mongoClientSettings.ServerApi = new ServerApi(ServerApiVersion.V1, default(Optional<bool?>), default(Optional<bool?>));
                MongoClient mongoClient = new MongoClient(mongoClientSettings);
                IMongoDatabase database = mongoClient.GetDatabase("duLieuAPP", null);
                IMongoCollection<Form1.DebtReminderManager.dataRungIcloud> collection = database.GetCollection<Form1.DebtReminderManager.dataRungIcloud>("dataRung", null);
                FilterDefinition<Form1.DebtReminderManager.dataRungIcloud> filterDefinition = Builders<Form1.DebtReminderManager.dataRungIcloud>.Filter.Eq<string>((Form1.DebtReminderManager.dataRungIcloud x) => x.Name, "dulieuhangngay");
                Form1.DebtReminderManager.dataRungIcloud dataRungIcloud = IFindFluentExtensions.FirstOrDefault<Form1.DebtReminderManager.dataRungIcloud, Form1.DebtReminderManager.dataRungIcloud>(IMongoCollectionExtensions.Find<Form1.DebtReminderManager.dataRungIcloud>(collection, filterDefinition, null), default(CancellationToken));
                result = (((dataRungIcloud != null) ? dataRungIcloud.Data : null) ?? "");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi kết nối MongoDB: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                result = "";
            }
            return result;
        }

        private string GetDataFromMongoDBByQH()
        {
            string result;
            try
            {
                string text = "mongodb+srv://banhmichaothuongnhoicloud:8390813asd@cluster0.srtd7rc.mongodb.net/?retryWrites=true&w=majority&appName=Cluster0";
                MongoClientSettings mongoClientSettings = MongoClientSettings.FromConnectionString(text);
                mongoClientSettings.ServerApi = new ServerApi(ServerApiVersion.V1, default(Optional<bool?>), default(Optional<bool?>));
                MongoClient mongoClient = new MongoClient(mongoClientSettings);
                IMongoDatabase database = mongoClient.GetDatabase("duLieuAPP", null);
                IMongoCollection<Form1.DebtReminderManager.dataRungIcloud> collection = database.GetCollection<Form1.DebtReminderManager.dataRungIcloud>("dataRung", null);
                FilterDefinition<Form1.DebtReminderManager.dataRungIcloud> filterDefinition = Builders<Form1.DebtReminderManager.dataRungIcloud>.Filter.Eq<string>((Form1.DebtReminderManager.dataRungIcloud x) => x.Name, "dulieuquahan");
                Form1.DebtReminderManager.dataRungIcloud dataRungIcloud = IFindFluentExtensions.FirstOrDefault<Form1.DebtReminderManager.dataRungIcloud, Form1.DebtReminderManager.dataRungIcloud>(IMongoCollectionExtensions.Find<Form1.DebtReminderManager.dataRungIcloud>(collection, filterDefinition, null), default(CancellationToken));
                result = (((dataRungIcloud != null) ? dataRungIcloud.Data : null) ?? "");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi kết nối MongoDB: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                result = "";
            }
            return result;
        }

        // Token: 0x0600007A RID: 122 RVA: 0x00006438 File Offset: 0x00004638
        private string GetDataFromMongoDB2()
        {
            string result;
            try
            {
                string text = "mongodb+srv://banhmichaothuongnhoicloud:8390813asd@cluster0.srtd7rc.mongodb.net/?retryWrites=true&w=majority&appName=Cluster0";
                MongoClientSettings mongoClientSettings = MongoClientSettings.FromConnectionString(text);
                mongoClientSettings.ServerApi = new ServerApi(ServerApiVersion.V1, default(Optional<bool?>), default(Optional<bool?>));
                MongoClient mongoClient = new MongoClient(mongoClientSettings);
                IMongoDatabase database = mongoClient.GetDatabase("duLieuAPP", null);
                IMongoCollection<Form1.DebtReminderManager.dataRungIcloud> collection = database.GetCollection<Form1.DebtReminderManager.dataRungIcloud>("dataRung", null);
                FilterDefinition<Form1.DebtReminderManager.dataRungIcloud> filterDefinition = Builders<Form1.DebtReminderManager.dataRungIcloud>.Filter.Eq<string>((Form1.DebtReminderManager.dataRungIcloud x) => x.Name, "dulieurungLine2");
                Form1.DebtReminderManager.dataRungIcloud dataRungIcloud = IFindFluentExtensions.FirstOrDefault<Form1.DebtReminderManager.dataRungIcloud, Form1.DebtReminderManager.dataRungIcloud>(IMongoCollectionExtensions.Find<Form1.DebtReminderManager.dataRungIcloud>(collection, filterDefinition, null), default(CancellationToken));
                result = (((dataRungIcloud != null) ? dataRungIcloud.Data : null) ?? "");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi kết nối MongoDB: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                result = "";
            }
            return result;
        }
        private void LoadEmptyDataGridView()
        {
            // Xóa dữ liệu cũ
            dataGridView1.Rows.Clear();

            // Thêm một số dòng trống để người dùng có thể nhập dữ liệu
            for (int i = 0; i < 5; i++)
            {
                dataGridView1.Rows.Add();
            }
        }

        // Token: 0x060000D9 RID: 217 RVA: 0x00012D70 File Offset: 0x00010F70
        public static void ExportToExcel(Dictionary<string, Form1.DebtReminderManager.DebtReminderInfo> dataList, string fileName)
        {
            ExcelPackage.LicenseContext = new LicenseContext?(0);
            string folderPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string text = Path.Combine(folderPath, fileName);
            try
            {
                using (ExcelPackage excelPackage = new ExcelPackage())
                {
                    ExcelWorksheet excelWorksheet = excelPackage.Workbook.Worksheets.Add("DebtReminderData");
                    excelWorksheet.Cells[1, 1].Value = "Mã HD";
                    excelWorksheet.Cells[1, 2].Value = "Tên KH";
                    excelWorksheet.Cells[1, 3].Value = "User iCloud";
                    excelWorksheet.Cells[1, 4].Value = "Pass iCloud";
                    excelWorksheet.Cells[1, 5].Value = "Name Phone Change";
                    excelWorksheet.Cells[1, 6].Value = "Note";
                    excelWorksheet.Cells[1, 7].Value = "Line";
                    excelWorksheet.Cells[1, 8].Value = "Date";
                    using (ExcelRange excelRange = excelWorksheet.Cells[1, 1, 1, 9])
                    {
                        excelRange.Style.Font.Bold = true;
                        excelRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        excelRange.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                        excelRange.Style.Border.BorderAround(ExcelBorderStyle.Thin); // ← Sửa ở đây
                    }
                    int num = 2;
                    foreach (KeyValuePair<string, Form1.DebtReminderManager.DebtReminderInfo> keyValuePair in dataList)
                    {
                        excelWorksheet.Cells[num, 1].Value = keyValuePair.Value.MaHD;
                        excelWorksheet.Cells[num, 2].Value = keyValuePair.Value.TenKh;
                        excelWorksheet.Cells[num, 3].Value = keyValuePair.Value.userIcloud;
                        excelWorksheet.Cells[num, 4].Value = keyValuePair.Value.PassIcloud;
                        excelWorksheet.Cells[num, 5].Value = keyValuePair.Value.NamePhoneChange;
                        excelWorksheet.Cells[num, 6].Value = keyValuePair.Value.Note;
                        excelWorksheet.Cells[num, 7].Value = keyValuePair.Value.Line;
                        excelWorksheet.Cells[num, 8].Value = keyValuePair.Value.date;
                        using (ExcelRange excelRange2 = excelWorksheet.Cells[num, 1, num, 9])
                        {
                            excelRange2.Style.Border.BorderAround(ExcelBorderStyle.Thin); // ← Sửa ở đây
                        }
                        num++;
                    }
                    excelWorksheet.Cells.AutoFitColumns();
                    FileInfo fileInfo = new FileInfo(text);
                    excelPackage.SaveAs(fileInfo);
                    Console.WriteLine("File Excel đã được tạo: " + text);
                    Form1.OpenFolderAndSelectFile(text);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Lỗi khi tạo file Excel: " + ex.Message);
            }
        }

        private static void OpenFolderAndSelectFile(string filePath)
        {
            try
            {
                Process.Start("explorer.exe", "/select,\"" + filePath + "\"");
            }
            catch (Exception ex)
            {
                try
                {
                    string directoryName = Path.GetDirectoryName(filePath);
                    Process.Start("explorer.exe", directoryName);
                }
                catch (Exception ex2)
                {
                    Console.WriteLine("Không thể mở thư mục: " + ex2.Message);
                }
            }
        }
        private void LoadDataToGridView(Dictionary<string, Form1.DebtReminderManager.DebtReminderInfo> debtData)
        {
            this.dataGridView1.Rows.Clear();
            int num = 0;
            foreach (KeyValuePair<string, Form1.DebtReminderManager.DebtReminderInfo> keyValuePair in debtData)
            {
                num++;
                Form1.DebtReminderManager.DebtReminderInfo value = keyValuePair.Value;
                string text = value.Note;
                bool flag = !string.IsNullOrEmpty(text) && !string.IsNullOrEmpty(value.date);
                if (flag)
                {
                    bool flag2 = text.ToLower().Contains("tạm tắt");
                    if (flag2)
                    {
                        Regex regex = new Regex("tạm tắt\\s*(\\d+)\\s*ngày", RegexOptions.IgnoreCase);
                        Match match = regex.Match(text);
                        int num2 = 0;
                        bool flag3 = match.Success && int.TryParse(match.Groups[1].Value, out num2);
                        if (flag3)
                        {
                            DateTime dateTime;
                            bool flag4 = DateTime.TryParseExact(value.date, "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out dateTime);
                            if (flag4)
                            {
                                DateTime t = dateTime.AddDays((double)num2);
                                DateTime date = DateTime.Now.Date;
                                bool flag5 = date >= t;
                                if (flag5)
                                {
                                    text = "";
                                }
                            }
                        }
                    }
                }
                int index = this.dataGridView1.Rows.Add();
                this.dataGridView1.Rows[index].Cells[0].Value = value.MaHD;
                this.dataGridView1.Rows[index].Cells[1].Value = value.TenKh;
                this.dataGridView1.Rows[index].Cells[2].Value = value.Line;
                this.dataGridView1.Rows[index].Cells[3].Value = text;
                this.dataGridView1.Rows[index].Cells[9].Value = value.NamePhoneChange;
                this.dataGridView1.Rows[index].Cells[10].Value = value.userIcloud;
                this.dataGridView1.Rows[index].Cells[11].Value = value.PassIcloud;
                this.dataGridView1.Rows[index].Cells[13].Value = num - 1;
                this.dataGridView1.Rows[index].Cells[14].Value = value.date;
                this.dataGridView1.Rows[index].Cells[8].Value = value.TenSo;
            }
        }

        private void button15_Click(object sender, EventArgs e)
        {
            bRunning = false;
            this.button14.Invoke(new MethodInvoker(delegate ()
            {
                this.button14.Enabled = true;
            }));
            this.button20.Invoke(new MethodInvoker(delegate ()
            {
                this.button20.Enabled = true;
            }));
        }

        private void button16_Click(object sender, EventArgs e)
        {
            // Đọc file JSON gốc
            string jsonContent = System.IO.File.ReadAllText("accounts.json");

            // Parse JSON thành Dictionary
            var accounts = JsonConvert.DeserializeObject<Dictionary<string, string>>(jsonContent);

            // Tạo Dictionary mới để lưu password đã giải mã
            var decryptedAccounts = new Dictionary<string, string>();

            // Giải mã từng password
            foreach (var account in accounts)
            {
                string email = account.Key;
                string encryptedPassword = account.Value;
                string decryptedPassword = DecryptPassword(encryptedPassword);

                decryptedAccounts[email] = decryptedPassword;

                // In ra kết quả để kiểm tra
                Console.WriteLine($"{email}: {encryptedPassword} -> {decryptedPassword}");
            }

            // Chuyển đổi về JSON và lưu vào file mới
            string decryptedJson = JsonConvert.SerializeObject(decryptedAccounts, Formatting.Indented);
            System.IO.File.WriteAllText("accounts_decrypted.json", decryptedJson);

            Console.WriteLine("\nFile đã được tạo: accounts_decrypted.json");
        }

        private void button24_Click(object sender, EventArgs e)
        {
            this.dataGridView1.AllowUserToAddRows = true;
            this.radioButtonxoa.Checked = true;
            try
            {
                string dataDeleteFromMongoDB = this.GetDataDeleteFromMongoDB();
                bool flag = string.IsNullOrWhiteSpace(dataDeleteFromMongoDB) || dataDeleteFromMongoDB == "[]" || dataDeleteFromMongoDB == "{}";
                if (flag)
                {
                    this.LoadEmptyDataGridView();
                    MessageBox.Show("Dữ liệu trống trong MongoDB. Đã tạo bảng trống để nhập dữ liệu.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                }
                else
                {
                    Dictionary<string, Form1.DebtReminderManager.DebtReminderInfo> dictionary = JsonConvert.DeserializeObject<Dictionary<string, Form1.DebtReminderManager.DebtReminderInfo>>(dataDeleteFromMongoDB);
                    bool flag2 = dictionary == null || dictionary.Count == 0;
                    if (flag2)
                    {
                        this.LoadEmptyDataGridView();
                        MessageBox.Show("Không có dữ liệu. Đã tạo bảng trống để nhập dữ liệu.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                    }
                    else
                    {
                        this.LoadDataToGridView(dictionary);
                        this.originalDataNew = dictionary;
                        originalData = new List<string>();
                        int num = 0;
                        foreach (KeyValuePair<string, Form1.DebtReminderManager.DebtReminderInfo> keyValuePair in dictionary)
                        {
                            num++;
                            string item = string.Format("{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}|{8}", new object[]
                            {
                                keyValuePair.Key,
                                keyValuePair.Value.Note,
                                keyValuePair.Value.userIcloud,
                                keyValuePair.Value.PassIcloud,
                                keyValuePair.Value.NamePhoneChange,
                                num - 1,
                                keyValuePair.Value.Line,
                                keyValuePair.Value.date,
                                keyValuePair.Value.TenSo
                            });
                            originalData.Add(item);
                        }
                        DialogResult dialogResult = MessageBox.Show("Bạn Muốn Lưu Dữ Liệu vào file Excell ", "LƯU DỮ LIỆU ĐÃ XÓA", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        bool flag3 = dialogResult != DialogResult.Yes;
                        if (!flag3)
                        {
                            string fileName = string.Format("DebtReminderData_{0:yyyyMMdd_HHmmss}.xlsx", DateTime.Now);
                            Form1.ExportToExcel(dictionary, fileName);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi xử lý dữ liệu: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                this.LoadEmptyDataGridView();
            }
        }

        private void richTextBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void button23_Click(object sender, EventArgs e)
        {
            Form1.shouldMerge = true;
            this.button5.Invoke(new MethodInvoker(delegate ()
            {
                this.radioButtonAll.Checked = true;
            }));
            List<string> userTomorowNew = this.GetUserTomorowNew("USR38122602", "SHP19263338");
            this.LoadAndSaveData(userTomorowNew);
        }

        private void button22_Click(object sender, EventArgs e)
        {
            Form1.shouldMerge = true;
            this.button5.Invoke(new MethodInvoker(delegate ()
            {
                this.radioButtonAll.Checked = true;
            }));
            List<string> userTomorowNew = this.GetUserTomorowNew("USR38122602", "SHP10275991");
            this.LoadAndSaveData(userTomorowNew);
        }

        private void button21_Click(object sender, EventArgs e)
        {
            Form1.shouldMerge = true;
            this.button5.Invoke(new MethodInvoker(delegate ()
            {
                this.radioButtonAll.Checked = true;
            }));
            List<string> userTomorowNew = this.GetUserTomorowNew("USR38122602", "SHP93791247");
            this.LoadAndSaveData(userTomorowNew);
        }

        private void button20_Click(object sender, EventArgs e)
        {
            Form1.shouldMerge = true;
            this.button5.Invoke(new MethodInvoker(delegate ()
            {
                this.radioButtonAll.Checked = true;
            }));
            List<string> userTomorowNew = this.GetUserTomorowNew("USR38122602", "SHP01201952");
            this.LoadAndSaveData(userTomorowNew);
        }

        private void button19_Click(object sender, EventArgs e)
        {
            Form1.shouldMerge = true;
            this.button5.Invoke(new MethodInvoker(delegate ()
            {
                this.radioButtonAll.Checked = true;
            }));
            List<string> userTomorowNew = this.GetUserTomorowNew("USR38122602", "SHP78278175");
            this.LoadAndSaveData(userTomorowNew);
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            try
            {
                this.button5.Invoke(new MethodInvoker(delegate ()
                {
                    this.radioButtonQH.Checked = true;
                }));
                checkQuaHan = true;
                List<string> userTomorowNew = this.GetUserQuaHan("USR38122602", "SHP01201952");
                this.LoadAndSaveData(userTomorowNew);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi khi đọc file: {ex.Message}", "Lỗi",
                                MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button9_Click_2(object sender, EventArgs e)
        {
            try
            {
                this.button5.Invoke(new MethodInvoker(delegate ()
                {
                    this.radioButtonQH.Checked = true;
                }));
                checkQuaHan = true;
                List<string> userTomorowNew = this.GetUserQuaHan("USR38122602", "SHP78278175");
                this.LoadAndSaveData(userTomorowNew);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi khi đọc file: {ex.Message}", "Lỗi",
                                MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
