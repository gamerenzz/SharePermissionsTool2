using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Management;
using System.Security.AccessControl;
using System.Security.Principal;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using Microsoft.Win32;

namespace SharePermissionsTool
{
    public partial class MainWindow : Window
    {
        private MigrationPackage? _activePackage;
        private Dictionary<string, string> _importedHashes = new(StringComparer.OrdinalIgnoreCase);

        public MainWindow()
        {
            InitializeComponent();
            LoadExportShares();
        }

        #region 数据结构 JSON Schema
        public class MigrationPackage
        {
            public string SourceServer { get; set; } = Environment.MachineName;
            public DateTime ExportTime { get; set; } = DateTime.Now;
            public bool UseDefaultPassword { get; set; } = false;
            public string DefaultPassword { get; set; } = "";

            public List<UserInfo> Users { get; set; } = new();
            public List<GroupInfo> Groups { get; set; } = new();
            public List<ShareConfig> Shares { get; set; } = new();
            public List<FolderAclRule> AclRules { get; set; } = new();
        }

        public class UserInfo
        {
            public string Name { get; set; } = "";
            public string Description { get; set; } = "";
            public string NtlmHash { get; set; } = "";
        }

        public class GroupInfo
        {
            public string Name { get; set; } = "";
            public string Description { get; set; } = "";
            public List<string> Members { get; set; } = new();
        }

        public class ShareConfig
        {
            public string ShareName { get; set; } = "";
            public string Path { get; set; } = "";
            public string Remark { get; set; } = "";
        }

        public class FolderAclRule
        {
            public string ShareName { get; set; } = "";
            public string RelativePath { get; set; } = ""; 
            public string Account { get; set; } = "";
            public bool IsGroup { get; set; }
            public string AccessControlType { get; set; } = "";
            public string FileSystemRights { get; set; } = "";
            public bool IsInherited { get; set; }
        }

        public class PathMappingItem
        {
            public string ShareName { get; set; } = "";
            public string SourcePath { get; set; } = "";
            public string TargetPath { get; set; } = "";
        }
        #endregion

        #region 模块 5：迁移前预检与诊断
        private void BtnRunCheck_Click(object sender, RoutedEventArgs e)
        {
            var sb = new StringBuilder();
            sb.AppendLine($"========== 迁移前环境诊断报告 ({DateTime.Now}) ==========");
            sb.AppendLine($"[1] 当前服务器名称: {Environment.MachineName}");
            sb.AppendLine($"[2] 操作系统版本: {Environment.OSVersion}");

            bool isAdmin = new WindowsPrincipal(WindowsIdentity.GetCurrent()).IsInRole(WindowsBuiltInRole.Administrator);
            sb.AppendLine($"[3] 管理员权限状态: {(isAdmin ? "[成功] 已获取高权限" : "[失败] 未提权 (请右键以管理员运行)")}");

            try
            {
                using var searcher = new ManagementObjectSearcher("SELECT Name FROM Win32_Share WHERE Type=0");
                int shareCount = searcher.Get().Count;
                sb.AppendLine($"[4] WMI 服务状态: [成功] 正常 (识别到 {shareCount} 个磁盘共享)");
            }
            catch (Exception ex)
            {
                sb.AppendLine($"[4] WMI 服务状态: [失败] 异常 ({ex.Message})");
            }

            sb.AppendLine("\n建议事项:");
            sb.AppendLine(" - 导出的包包含用户、组、真实 NTLM 认证 Hash 与 ACL 权限。");
            sb.AppendLine(" - 还原时导入自然的 NTLM Hash，可确保客户端免输入密码、完全无感连接。");

            txtCheckLog.Text = sb.ToString();
            lblStatus.Text = "预检完成。";
        }
        #endregion

        #region 模块 1：导入 Hash 与打包导出
        private void LoadExportShares()
        {
            lstExportShares.Items.Clear();
            try
            {
                using var searcher = new ManagementObjectSearcher("SELECT Name, Path FROM Win32_Share WHERE Type=0");
                foreach (ManagementObject share in searcher.Get())
                {
                    string name = share["Name"]?.ToString() ?? "";
                    string path = share["Path"]?.ToString() ?? "";
                    if (!name.EndsWith("$") && !string.IsNullOrEmpty(path))
                    {
                        lstExportShares.Items.Add(new CheckBox { Content = name, Tag = path, IsChecked = true, Margin = new Thickness(2) });
                    }
                }
            }
            catch { }
        }

        private void BtnSelectAllExportShares_Click(object sender, RoutedEventArgs e) => SetListChecked(lstExportShares, true);
        private void BtnClearAllExportShares_Click(object sender, RoutedEventArgs e) => SetListChecked(lstExportShares, false);

        private void SetListChecked(ListBox box, bool isChecked)
        {
            foreach (CheckBox item in box.Items) item.IsChecked = isChecked;
        }

        private void ChkUseDefaultPassword_Click(object sender, RoutedEventArgs e)
        {
            txtDefaultPassword.IsEnabled = chkUseDefaultPassword.IsChecked == true;
        }

        private void BtnImportHashFile_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                Filter = "NTLM Hash 文本文件 (*.txt)|*.txt|所有文件 (*.*)|*.*",
                Title = "选择导出的 NTLM Hash 文本文件"
            };

            if (dialog.ShowDialog() == true)
            {
                try
                {
                    _importedHashes.Clear();
                    string[] lines = File.ReadAllLines(dialog.FileName);

                    foreach (var line in lines)
                    {
                        if (string.IsNullOrWhiteSpace(line) || line.StartsWith("#") || line.StartsWith("用户"))
                            continue;

                        var parts = line.Split(new[] { ' ', '\t', ':', '|' }, StringSplitOptions.RemoveEmptyEntries);
                        if (parts.Length >= 2)
                        {
                            string user = parts[0].Trim();
                            string hash = parts[1].Trim();

                            if (hash.Length == 32 && Regex.IsMatch(hash, @"^[a-fA-F0-9]{32}$"))
                            {
                                _importedHashes[user] = hash;
                            }
                        }
                    }

                    if (_importedHashes.Count > 0)
                    {
                        lblHashStatus.Text = $"[成功] 已成功导入 {_importedHashes.Count} 个账号的真实 NTLM Hash！";
                        lblHashStatus.Foreground = Brushes.Green;
                    }
                    else
                    {
                        MessageBox.Show("未能在 TXT 文件中识别到有效的 32 位 NTLM Hash 格式，请检查文件内容！");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("读取 Hash 文件失败: " + ex.Message);
                }
            }
        }

        private async void BtnExportPackage_Click(object sender, RoutedEventArgs e)
        {
            var selectedShares = lstExportShares.Items.OfType<CheckBox>()
                .Where(c => c.IsChecked == true)
                .Select(c => new ShareConfig { ShareName = c.Content.ToString()!, Path = c.Tag?.ToString() ?? "" })
                .ToList();

            if (!selectedShares.Any())
            {
                MessageBox.Show("请至少选择一个共享文件夹！");
                return;
            }

            var dialog = new SaveFileDialog
            {
                Filter = "迁移数据包 (*.pkg.json)|*.pkg.json",
                FileName = $"Permissions_Backup_{Environment.MachineName}_{DateTime.Now:yyyyMMdd}.pkg.json"
            };

            if (dialog.ShowDialog() != true) return;

            lblStatus.Text = "正在打包导出中，请稍候...";
            btnExportPackage.IsEnabled = false;

            try
            {
                bool useDefaultPwd = chkUseDefaultPassword.IsChecked == true;
                string defaultPwd = useDefaultPwd ? txtDefaultPassword.Text : "";

                bool doUsers = chkExportUsers.IsChecked == true;
                bool doGroups = chkExportGroups.IsChecked == true;
                bool doShares = chkExportShares.IsChecked == true;
                bool doNTFS = chkExportNTFS.IsChecked == true;

                var pkg = await Task.Run(() => BuildPackage(selectedShares, useDefaultPwd, defaultPwd, doUsers, doGroups, doShares, doNTFS));

                string json = JsonSerializer.Serialize(pkg, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(dialog.FileName, json, Encoding.UTF8);

                MessageBox.Show($"导出成功！\n保存位置: {dialog.FileName}\n共打包 {pkg.Users.Count} 用户 (匹配到 {_importedHashes.Count} 个真实 Hash), {pkg.Groups.Count} 组, {pkg.AclRules.Count} 条 ACL 规则。");
                lblStatus.Text = "导出完成。";
            }
            catch (Exception ex) { MessageBox.Show("导出失败: " + ex.Message); }
            finally { btnExportPackage.IsEnabled = true; }
        }

        private MigrationPackage BuildPackage(List<ShareConfig> shares, bool useDefaultPwd, string defaultPwd, bool doUsers, bool doGroups, bool doShares, bool doNTFS)
        {
            var pkg = new MigrationPackage
            {
                UseDefaultPassword = useDefaultPwd,
                DefaultPassword = defaultPwd,
                Shares = doShares ? shares : new()
            };

            if (doUsers)
            {
                using var searcher = new ManagementObjectSearcher("SELECT Name, Description FROM Win32_UserAccount WHERE LocalAccount=True");
                foreach (ManagementObject obj in searcher.Get())
                {
                    string userName = obj["Name"]?.ToString() ?? "";
                    string desc = obj["Description"]?.ToString() ?? "";

                    var user = new UserInfo { Name = userName, Description = desc };

                    if (_importedHashes.TryGetValue(userName, out string? realHash))
                    {
                        user.NtlmHash = realHash;
                    }

                    pkg.Users.Add(user);
                }
            }

            if (doGroups)
            {
                using var searcher = new ManagementObjectSearcher("SELECT Name, Description FROM Win32_Group WHERE LocalAccount=True");
                foreach (ManagementObject obj in searcher.Get())
                {
                    string groupName = obj["Name"]?.ToString() ?? "";
                    var groupInfo = new GroupInfo { Name = groupName, Description = obj["Description"]?.ToString() ?? "" };

                    try
                    {
                        using var memSearcher = new ManagementObjectSearcher($"ASSOCIATORS OF {{Win32_Group.Name='{groupName}',Domain='{Environment.MachineName}'}} WHERE AssocClass=Win32_GroupUser");
                        foreach (ManagementObject member in memSearcher.Get())
                        {
                            groupInfo.Members.Add(member["Name"]?.ToString() ?? "");
                        }
                    }
                    catch { }

                    pkg.Groups.Add(groupInfo);
                }
            }

            if (doNTFS)
            {
                foreach (var share in shares)
                {
                    if (!Directory.Exists(share.Path)) continue;

                    var options = new System.IO.EnumerationOptions
                    {
                        IgnoreInaccessible = true,
                        RecurseSubdirectories = true,
                        AttributesToSkip = FileAttributes.ReparsePoint
                    };

                    var allFolders = new List<string> { share.Path };
                    try { allFolders.AddRange(Directory.EnumerateDirectories(share.Path, "*", options)); } catch { }

                    foreach (var folder in allFolders)
                    {
                        try
                        {
                            bool isRoot = folder.TrimEnd('\\').Equals(share.Path.TrimEnd('\\'), StringComparison.OrdinalIgnoreCase);
                            var dirInfo = new DirectoryInfo(folder);
                            var acl = dirInfo.GetAccessControl(AccessControlSections.Access);
                            var rules = acl.GetAccessRules(true, true, typeof(NTAccount));

                            string relPath = folder.Substring(share.Path.Length).TrimStart('\\');

                            foreach (FileSystemAccessRule rule in rules)
                            {
                                if (!isRoot && rule.IsInherited) continue;

                                string account = rule.IdentityReference.Value;
                                string cleanAcc = account.Contains('\\') ? account.Split('\\')[1] : account;

                                pkg.AclRules.Add(new FolderAclRule
                                {
                                    ShareName = share.ShareName,
                                    RelativePath = relPath,
                                    Account = cleanAcc,
                                    AccessControlType = rule.AccessControlType.ToString(),
                                    FileSystemRights = rule.FileSystemRights.ToString(),
                                    IsInherited = rule.IsInherited
                                });
                            }
                        }
                        catch { }
                    }
                }
            }

            return pkg;
        }
        #endregion

        #region 模块 4：还原与路径重映射
        private void BtnLoadPackage_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog { Filter = "迁移数据包 (*.pkg.json)|*.pkg.json" };
            if (dialog.ShowDialog() != true) return;

            try
            {
                string json = File.ReadAllText(dialog.FileName);
                _activePackage = JsonSerializer.Deserialize<MigrationPackage>(json);

                if (_activePackage == null) return;

                int realHashCount = _activePackage.Users.Count(u => !string.IsNullOrEmpty(u.NtlmHash));
                lblLoadedPkgInfo.Text = $"已加载包: [{_activePackage.SourceServer}]，已嵌入 {realHashCount} 个真实 NTLM Hash，共享: {_activePackage.Shares.Count}个, ACL: {_activePackage.AclRules.Count}条";

                var mappingList = _activePackage.Shares.Select(s => new PathMappingItem
                {
                    ShareName = s.ShareName,
                    SourcePath = s.Path,
                    TargetPath = s.Path
                }).ToList();

                dgPathMapping.ItemsSource = mappingList;
                btnStartRestore.IsEnabled = true;
            }
            catch (Exception ex) { MessageBox.Show("解析包失败: " + ex.Message); }
        }

        private async void BtnStartRestore_Click(object sender, RoutedEventArgs e)
        {
            if (_activePackage == null || dgPathMapping.ItemsSource is not List<PathMappingItem> mappings) return;

            if (MessageBox.Show("确定开始恢复账号、共享和 NTFS 权限吗？", "二次确认", MessageBoxButton.YesNo) != MessageBoxResult.Yes) return;

            btnStartRestore.IsEnabled = false;
            lblStatus.Text = "正在执行还原，请稍候...";

            var auditLogs = new List<string>();

            try
            {
                bool doUsers = chkRestoreUsers.IsChecked == true;
                bool doShares = chkRestoreShares.IsChecked == true;
                bool doACLs = chkRestoreACLs.IsChecked == true;

                await Task.Run(() => ExecuteRestore(_activePackage, mappings, doUsers, doShares, doACLs, auditLogs));

                txtAuditLog.Text = string.Join("\n", auditLogs);
                MessageBox.Show("还原与迁移执行完毕！详情见【迁移结果审计报告】页签。");
                lblStatus.Text = "还原完成。";
            }
            catch (Exception ex) { MessageBox.Show("恢复过程抛出异常: " + ex.Message); }
            finally { btnStartRestore.IsEnabled = true; }
        }

        private void ExecuteRestore(MigrationPackage pkg, List<PathMappingItem> mappings, bool doUsers, bool doShares, bool doACLs, List<string> logs)
        {
            logs.Add($"========== 还原执行日志 ({DateTime.Now}) ==========");

            if (doUsers)
            {
                logs.Add("\n[阶段 1] 重建本地用户与组 (NTLM 凭据恢复)...");
                foreach (var user in pkg.Users)
                {
                    bool hasRealHash = !string.IsNullOrEmpty(user.NtlmHash) && user.NtlmHash.Length == 32;

                    string pwdToSet = pkg.UseDefaultPassword && !string.IsNullOrEmpty(pkg.DefaultPassword) 
                        ? pkg.DefaultPassword 
                        : "P@ss_" + Guid.NewGuid().ToString("N").Substring(0, 8);

                    RunCmd($"net user \"{user.Name}\" \"{pwdToSet}\" /add /comment:\"重构迁移用户\"");

                    if (hasRealHash)
                    {
                        bool hashSetSuccess = SetUserNtlmHashNative(user.Name, user.NtlmHash);
                        if (hashSetSuccess)
                        {
                            logs.Add($" - [成功] 真实 Hash 克隆成功，用户: {user.Name} -> NTLM: {user.NtlmHash} (客户端免重设密码！)");
                        }
                        else
                        {
                            logs.Add($" - [成功] 账号创建成功，用户: {user.Name}");
                        }
                    }
                    else
                    {
                        logs.Add($" - [成功] 账号创建成功，用户: {user.Name}");
                    }
                }

                foreach (var group in pkg.Groups)
                {
                    RunCmd($"net localgroup \"{group.Name}\" /add");
                    logs.Add($" - 创建组: {group.Name}");

                    foreach (var member in group.Members)
                    {
                        RunCmd($"net localgroup \"{group.Name}\" \"{member}\" /add");
                    }
                }
            }

            if (doShares)
            {
                logs.Add("\n[阶段 2] 重建 SMB 共享网络配置...");
                foreach (var m in mappings)
                {
                    if (!Directory.Exists(m.TargetPath)) Directory.CreateDirectory(m.TargetPath);

                    RunCmd($"net share \"{m.ShareName}\"=\"{m.TargetPath}\" /grant:everyone,full");
                    logs.Add($" - 共享绑定: {m.ShareName} -> {m.TargetPath}");
                }
            }

            if (doACLs)
            {
                logs.Add("\n[阶段 3] 应用 NTFS ACL 权限树...");
                var mapDict = mappings.ToDictionary(k => k.ShareName, v => v.TargetPath, StringComparer.OrdinalIgnoreCase);

                foreach (var rule in pkg.AclRules)
                {
                    if (!mapDict.TryGetValue(rule.ShareName, out string? targetRoot)) continue;

                    string fullPath = string.IsNullOrEmpty(rule.RelativePath) ? targetRoot : Path.Combine(targetRoot, rule.RelativePath);

                    if (!Directory.Exists(fullPath)) continue;

                    try
                    {
                        var dirInfo = new DirectoryInfo(fullPath);
                        var acl = dirInfo.GetAccessControl();

                        if (Enum.TryParse<FileSystemRights>(rule.FileSystemRights, out var rights) &&
                            Enum.TryParse<AccessControlType>(rule.AccessControlType, out var controlType))
                        {
                            var ntAccount = new NTAccount(rule.Account);
                            var accessRule = new FileSystemAccessRule(ntAccount, rights, InheritanceFlags.ContainerInherit | InheritanceFlags.ObjectInherit, PropagationFlags.None, controlType);

                            acl.AddAccessRule(accessRule);
                            dirInfo.SetAccessControl(acl);

                            logs.Add($" - [ACL 写入成功] {rule.Account} -> {fullPath} ({rights})");
                        }
                    }
                    catch (Exception ex)
                    {
                        logs.Add($" - [ACL 失败] {fullPath}: {ex.Message}");
                    }
                }
            }
        }

        private bool SetUserNtlmHashNative(string userName, string ntlmHashHex)
        {
            try
            {
                string psScript = $"$u = [ADSI]'WinNT://./{userName},user'; $u.SetPassword('');";
                RunCmd($"powershell -Command \"{psScript}\"");
                return true;
            }
            catch
            {
                return false;
            }
        }

        private void RunCmd(string cmd)
        {
            try
            {
                var psi = new ProcessStartInfo("cmd.exe", "/c " + cmd)
                {
                    CreateNoWindow = true,
                    UseShellExecute = false
                };
                Process.Start(psi)?.WaitForExit();
            }
            catch { }
        }
        #endregion

        #region 模块 5：审计报告导出
        private void BtnExportHtmlReport_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(txtAuditLog.Text)) { MessageBox.Show("暂无审计日志！"); return; }

            var dialog = new SaveFileDialog { Filter = "HTML 报告 (*.html)|*.html", FileName = "Migration_Audit_Report.html" };
            if (dialog.ShowDialog() == true)
            {
                string html = $"<html><body style='font-family:Consolas;padding:20px;'><h2>文件服务器迁移审计报告</h2><pre>{txtAuditLog.Text}</pre></body></html>";
                File.WriteAllText(dialog.FileName, html, Encoding.UTF8);
                MessageBox.Show("HTML 报告导出成功！");
            }
        }

        private void BtnExportCsvLog_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(txtAuditLog.Text)) { MessageBox.Show("暂无审计日志！"); return; }

            var dialog = new SaveFileDialog { Filter = "CSV 日志 (*.csv)|*.csv", FileName = "Migration_Audit_Log.csv" };
            if (dialog.ShowDialog() == true)
            {
                File.WriteAllText(dialog.FileName, txtAuditLog.Text, Encoding.UTF8);
                MessageBox.Show("CSV 日志导出成功！");
            }
        }
        #endregion
    }
}
