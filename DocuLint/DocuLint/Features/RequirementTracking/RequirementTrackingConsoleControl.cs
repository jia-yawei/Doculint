using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace DocuLint
{
    internal sealed class RequirementTrackingConsoleControl : UserControl
    {
        private const string SearchPlaceholderText = "模糊搜索名称/ID";

        private readonly Func<Word.Application> applicationAccessor;
        private readonly ComboBox cmbSourceDoc;
        private readonly ComboBox cmbTargetDoc;
        private readonly Button btnBrowseSourceDoc;
        private readonly Button btnBrowseTargetDoc;
        private readonly Button btnLoadSourceData;
        private readonly Button btnLoadTargetData;
        private readonly ListBox lstSourceReqs;
        private readonly CheckedListBox chkListTargetReqs;
        private readonly TextBox txtSearchTarget;
        private readonly Label lblSourceHeader;
        private readonly Label lblTargetHeader;
        private readonly Label lblProgress;

        private readonly Dictionary<string, RequirementTraceMapping> mappingsBySourceId =
            new Dictionary<string, RequirementTraceMapping>(StringComparer.OrdinalIgnoreCase);

        private readonly List<RequirementTrackingDocumentOption> openDocumentOptions =
            new List<RequirementTrackingDocumentOption>();

        private readonly List<string> manuallySelectedDocumentPaths =
            new List<string>();

        private readonly HashSet<string> selectedTargetRequirementIds =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        private readonly List<RequirementItem> visibleTargetRequirements =
            new List<RequirementItem>();

        private RequirementTrackingDocumentSnapshot sourceSnapshot;
        private RequirementTrackingDocumentSnapshot targetSnapshot;
        private bool suppressSourceSelection;
        private bool suppressTargetChecks;
        private bool suppressSearchRefresh;

        internal RequirementTrackingConsoleControl(Func<Word.Application> applicationAccessor)
        {
            this.applicationAccessor = applicationAccessor ?? throw new ArgumentNullException(nameof(applicationAccessor));

            Dock = DockStyle.Fill;
            BackColor = Color.White;
            Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Regular, GraphicsUnit.Point);
            AutoScaleMode = AutoScaleMode.Dpi;

            TableLayoutPanel rootLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 3,
                Padding = new Padding(10)
            };
            rootLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 190f));
            rootLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
            rootLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 34f));

            GroupBox topGroup = new GroupBox
            {
                Dock = DockStyle.Fill,
                Text = "需求追踪控制台",
                Font = new Font("Microsoft YaHei UI", 10.5F, FontStyle.Bold, GraphicsUnit.Point),
                Padding = new Padding(10, 14, 10, 12)
            };

            TableLayoutPanel topLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 3,
                RowCount = 3
            };
            topLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 150f));
            topLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            topLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 72f));
            topLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 38f));
            topLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 38f));
            topLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 38f));

            Label lblSourceDoc = CreateFieldLabel("基准文档");
            Label lblTargetDoc = CreateFieldLabel("目标文档");

            cmbSourceDoc = CreateDocumentComboBox();
            cmbTargetDoc = CreateDocumentComboBox();
            btnBrowseSourceDoc = CreateBrowseButton();
            btnBrowseSourceDoc.Click += BtnBrowseSourceDoc_Click;
            btnBrowseTargetDoc = CreateBrowseButton();
            btnBrowseTargetDoc.Click += BtnBrowseTargetDoc_Click;

            btnLoadSourceData = CreateLoadButton("加载并解析基准文档需求");
            btnLoadSourceData.Dock = DockStyle.None;
            btnLoadSourceData.Width = 256;
            btnLoadSourceData.Anchor = AnchorStyles.None;
            btnLoadSourceData.Click += BtnLoadSourceData_Click;
            btnLoadTargetData = CreateLoadButton("加载并解析目标文档需求");
            btnLoadTargetData.Dock = DockStyle.None;
            btnLoadTargetData.Width = 256;
            btnLoadTargetData.Anchor = AnchorStyles.None;
            btnLoadTargetData.Click += BtnLoadTargetData_Click;

            TableLayoutPanel loadButtonLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 1,
                Margin = new Padding(0)
            };
            loadButtonLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50f));
            loadButtonLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50f));
            loadButtonLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
            loadButtonLayout.Controls.Add(btnLoadSourceData, 0, 0);
            loadButtonLayout.Controls.Add(btnLoadTargetData, 1, 0);

            topLayout.Controls.Add(lblSourceDoc, 0, 0);
            topLayout.Controls.Add(cmbSourceDoc, 1, 0);
            topLayout.Controls.Add(btnBrowseSourceDoc, 2, 0);
            topLayout.Controls.Add(lblTargetDoc, 0, 1);
            topLayout.Controls.Add(cmbTargetDoc, 1, 1);
            topLayout.Controls.Add(btnBrowseTargetDoc, 2, 1);
            topLayout.Controls.Add(loadButtonLayout, 0, 2);
            topLayout.SetColumnSpan(loadButtonLayout, 3);
            topGroup.Controls.Add(topLayout);

            TableLayoutPanel middleLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 2,
                Padding = new Padding(0, 6, 0, 6)
            };
            middleLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 44f));
            middleLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));

            Label lblMappingTitle = new Label
            {
                Dock = DockStyle.Fill,
                Font = new Font("Microsoft YaHei UI", 10.5F, FontStyle.Bold, GraphicsUnit.Point),
                Text = "需求映射与关系建立",
                TextAlign = ContentAlignment.MiddleLeft
            };

            Button btnReverseTrace = CreateLoadButton("反向追踪");
            btnReverseTrace.Dock = DockStyle.None;
            btnReverseTrace.Width = 186;
            btnReverseTrace.Anchor = AnchorStyles.None;
            btnReverseTrace.Margin = new Padding(0);
            btnReverseTrace.Click += BtnReverseTrace_Click;

            TableLayoutPanel mappingHeaderLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2
            };
            mappingHeaderLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            mappingHeaderLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 202f));
            mappingHeaderLayout.Controls.Add(lblMappingTitle, 0, 0);
            mappingHeaderLayout.Controls.Add(btnReverseTrace, 1, 0);

            TableLayoutPanel splitContainer = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 1
            };
            splitContainer.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50f));
            splitContainer.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50f));
            splitContainer.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));

            TableLayoutPanel leftPaneLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 2,
                Padding = new Padding(0)
            };
            leftPaneLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 28f));
            leftPaneLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));

            lblSourceHeader = new Label
            {
                Dock = DockStyle.Fill,
                Font = new Font("Microsoft YaHei UI", 9.5F, FontStyle.Bold, GraphicsUnit.Point),
                Text = "需求列表",
                TextAlign = ContentAlignment.MiddleLeft
            };

            lstSourceReqs = new ListBox
            {
                Dock = DockStyle.Fill,
                IntegralHeight = false,
                HorizontalScrollbar = true
            };
            lstSourceReqs.SelectedIndexChanged += LstSourceReqs_SelectedIndexChanged;

            leftPaneLayout.Controls.Add(lblSourceHeader, 0, 0);
            leftPaneLayout.Controls.Add(lstSourceReqs, 0, 1);
            Panel leftPanePanel = CreateBorderPanel();
            leftPanePanel.Controls.Add(leftPaneLayout);

            TableLayoutPanel rightPaneLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 3
            };
            rightPaneLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 28f));
            rightPaneLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 34f));
            rightPaneLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));

            lblTargetHeader = new Label
            {
                Dock = DockStyle.Fill,
                Font = new Font("Microsoft YaHei UI", 9.5F, FontStyle.Bold, GraphicsUnit.Point),
                Text = "目标需求候选（多选）",
                TextAlign = ContentAlignment.MiddleLeft
            };

            txtSearchTarget = new TextBox
            {
                Dock = DockStyle.Fill,
                ForeColor = Color.Gray,
                Text = SearchPlaceholderText
            };
            txtSearchTarget.Enter += TxtSearchTarget_Enter;
            txtSearchTarget.Leave += TxtSearchTarget_Leave;
            txtSearchTarget.TextChanged += TxtSearchTarget_TextChanged;

            chkListTargetReqs = new CheckedListBox
            {
                Dock = DockStyle.Fill,
                CheckOnClick = true,
                IntegralHeight = false,
                HorizontalScrollbar = true
            };
            chkListTargetReqs.ItemCheck += ChkListTargetReqs_ItemCheck;

            rightPaneLayout.Controls.Add(lblTargetHeader, 0, 0);
            rightPaneLayout.Controls.Add(txtSearchTarget, 0, 1);
            rightPaneLayout.Controls.Add(chkListTargetReqs, 0, 2);
            Panel rightPanePanel = CreateBorderPanel();
            rightPanePanel.Controls.Add(rightPaneLayout);
            splitContainer.Controls.Add(leftPanePanel, 0, 0);
            splitContainer.Controls.Add(rightPanePanel, 1, 0);

            middleLayout.Controls.Add(mappingHeaderLayout, 0, 0);
            middleLayout.Controls.Add(splitContainer, 0, 1);

            Panel bottomPanel = new Panel
            {
                Dock = DockStyle.Fill
            };

            lblProgress = new Label
            {
                Dock = DockStyle.Fill,
                ForeColor = Color.FromArgb(70, 70, 70),
                Text = "追踪进度：0/0",
                TextAlign = ContentAlignment.MiddleLeft
            };

            bottomPanel.Controls.Add(lblProgress);

            rootLayout.Controls.Add(topGroup, 0, 0);
            rootLayout.Controls.Add(middleLayout, 0, 1);
            rootLayout.Controls.Add(bottomPanel, 0, 2);
            Controls.Add(rootLayout);

            RefreshDocumentOptions();
        }

        internal void RefreshDocumentOptions()
        {
            string previousSourceFullName = GetSelectedDocumentFullName(cmbSourceDoc);
            string previousTargetFullName = GetSelectedDocumentFullName(cmbTargetDoc);

            openDocumentOptions.Clear();
            openDocumentOptions.AddRange(BuildDocumentOptions());

            BindDocumentOptions(cmbSourceDoc, previousSourceFullName);
            BindDocumentOptions(cmbTargetDoc, previousTargetFullName);
            ApplyPreferredDocumentSelections(previousSourceFullName, previousTargetFullName);
        }

        private void BtnLoadSourceData_Click(object sender, EventArgs e)
        {
            LoadRequirements(cmbSourceDoc, btnLoadSourceData, "基准文档", true);
        }

        private void BtnLoadTargetData_Click(object sender, EventArgs e)
        {
            LoadRequirements(cmbTargetDoc, btnLoadTargetData, "目标文档", false);
        }

        private void LoadRequirements(ComboBox documentComboBox, Button triggerButton, string documentRole, bool isSource)
        {
            Word.Document document = null;
            bool previousUseWaitCursor = UseWaitCursor;
            string previousButtonText = triggerButton.Text;
            try
            {
                UseWaitCursor = true;
                triggerButton.Enabled = false;
                triggerButton.Text = "正在解析...";
                triggerButton.Refresh();
                ReportLoadProgress("正在准备需求解析...");

                string documentPath = GetSelectedDocumentFullName(documentComboBox);
                if (string.IsNullOrWhiteSpace(documentPath))
                {
                    throw new InvalidOperationException($"请先选择一个{documentRole}。");
                }

                RequirementTrackingDocumentKind documentKind = DetectRequirementTrackingDocumentKind(documentPath);
                if (documentKind == RequirementTrackingDocumentKind.Unknown)
                {
                    throw new InvalidOperationException(
                        $"当前选择的文档不属于需求追踪支持范围。\r\n当前选择：{Path.GetFileName(documentPath)}\r\n\r\n必须选择以下四类文档之一：系统规格说明、需求规格说明、软件设计说明、软件测试说明。");
                }

                document = OpenOrResolveDocument(documentPath);
                if (document == null)
                {
                    throw new InvalidOperationException($"无法打开{documentRole}，请检查文件是否存在或被占用。");
                }
                ReportLoadProgress($"正在解析{documentRole}：{Path.GetFileName(documentPath)}");

                RequirementTrackingDocumentSnapshot loadedSnapshot = RequirementTrackingWordService.CollectRequirements(
                    document,
                    documentKind,
                    message => ReportLoadProgress($"[{Path.GetFileName(documentPath)}] {message}"));

                if ((loadedSnapshot.Requirements?.Count ?? 0) == 0)
                {
                    throw new InvalidOperationException(
                        $"{documentRole}中未识别到需求标识。\r\n当前解析文档：{Path.GetFileName(documentPath)}\r\n文档类型：{GetDocumentKindDisplayName(documentKind)}");
                }

                mappingsBySourceId.Clear();
                selectedTargetRequirementIds.Clear();
                ResetSearchPlaceholder();
                if (isSource)
                {
                    sourceSnapshot = loadedSnapshot;
                }
                else
                {
                    targetSnapshot = loadedSnapshot;
                }

                RenderSourceRequirements(null);
                RenderTargetRequirements();
                UpdateHeaders();
                UpdateProgressLabel();
                ReportLoadProgress($"已提取{documentRole}需求：{loadedSnapshot.Requirements.Count} 项");
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message, "需求追踪控制台");
            }
            finally
            {
                UseWaitCursor = previousUseWaitCursor;
                triggerButton.Text = previousButtonText;
                triggerButton.Enabled = true;
                // 不要释放文档对象的 COM 包装器，否则会导致用户自行打开的 Word 文档后续操作抛出“脱离基础 RCW”的异常。
            }
        }

        private static RequirementTrackingDocumentKind DetectRequirementTrackingDocumentKind(string fullName)
        {
            string fileName = Path.GetFileNameWithoutExtension(fullName) ?? string.Empty;
            if (ContainsDocumentNameToken(fileName, "系统规格说明", "系统/子系统规格说明", "系统子系统规格说明", "SSS"))
            {
                return RequirementTrackingDocumentKind.SystemSpecification;
            }

            if (ContainsDocumentNameToken(fileName, "需求规格说明", "软件需求规格说明", "SRS"))
            {
                return RequirementTrackingDocumentKind.RequirementSpecification;
            }

            if (ContainsDocumentNameToken(fileName, "软件设计说明", "软件设计描述", "SDD", "SDS"))
            {
                return RequirementTrackingDocumentKind.SoftwareDesignDescription;
            }

            if (ContainsDocumentNameToken(fileName, "软件测试说明", "软件测试描述", "测试说明", "STD", "STS"))
            {
                return RequirementTrackingDocumentKind.SoftwareTestDescription;
            }

            return RequirementTrackingDocumentKind.Unknown;
        }

        private static bool ContainsDocumentNameToken(string fileName, params string[] tokens)
        {
            if (string.IsNullOrWhiteSpace(fileName) || tokens == null)
            {
                return false;
            }

            return tokens.Any(token =>
            {
                if (string.IsNullOrWhiteSpace(token))
                {
                    return false;
                }

                if (Regex.IsMatch(token, @"^[A-Za-z0-9]+$"))
                {
                    return Regex.IsMatch(
                        fileName,
                        $@"(^|[^A-Za-z0-9]){Regex.Escape(token)}([^A-Za-z0-9]|$)",
                        RegexOptions.IgnoreCase);
                }

                return fileName.IndexOf(token, StringComparison.OrdinalIgnoreCase) >= 0;
            });
        }

        private static string GetDocumentKindDisplayName(RequirementTrackingDocumentKind documentKind)
        {
            switch (documentKind)
            {
                case RequirementTrackingDocumentKind.SystemSpecification:
                    return "系统规格说明";
                case RequirementTrackingDocumentKind.RequirementSpecification:
                    return "需求规格说明";
                case RequirementTrackingDocumentKind.SoftwareDesignDescription:
                    return "软件设计说明";
                case RequirementTrackingDocumentKind.SoftwareTestDescription:
                    return "软件测试说明";
                default:
                    return "未知文档";
            }
        }

        private void BtnBrowseSourceDoc_Click(object sender, EventArgs e)
        {
            BrowseAndSelectDocument(cmbSourceDoc);
        }

        private void BtnBrowseTargetDoc_Click(object sender, EventArgs e)
        {
            BrowseAndSelectDocument(cmbTargetDoc);
        }

        private void LstSourceReqs_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (suppressSourceSelection)
            {
                return;
            }

            RenderTargetRequirements();
        }

        private void BtnReverseTrace_Click(object sender, EventArgs e)
        {
            string sourceFullName = GetSelectedDocumentFullName(cmbSourceDoc);
            string targetFullName = GetSelectedDocumentFullName(cmbTargetDoc);

            RequirementTrackingDocumentSnapshot swappedSourceSnapshot = targetSnapshot;
            RequirementTrackingDocumentSnapshot swappedTargetSnapshot = sourceSnapshot;

            sourceSnapshot = swappedSourceSnapshot;
            targetSnapshot = swappedTargetSnapshot;

            mappingsBySourceId.Clear();
            selectedTargetRequirementIds.Clear();
            suppressTargetChecks = false;
            ResetSearchPlaceholder();

            if (!string.IsNullOrWhiteSpace(targetFullName))
            {
                SelectDocumentByFullName(cmbSourceDoc, targetFullName);
            }

            if (!string.IsNullOrWhiteSpace(sourceFullName))
            {
                SelectDocumentByFullName(cmbTargetDoc, sourceFullName);
            }

            RenderSourceRequirements(null);
            RenderTargetRequirements();
            UpdateHeaders();
            UpdateProgressLabel();
        }

        private void ChkListTargetReqs_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            if (suppressTargetChecks)
            {
                return;
            }

            RequirementItem currentSourceRequirement = GetSelectedSourceRequirement();
            if (e.Index < 0 || e.Index >= visibleTargetRequirements.Count)
            {
                return;
            }

            RequirementItem targetRequirement = visibleTargetRequirements[e.Index];
            if (!string.IsNullOrWhiteSpace(targetRequirement?.Id))
            {
                if (e.NewValue == CheckState.Checked)
                {
                    selectedTargetRequirementIds.Add(targetRequirement.Id);
                }
                else
                {
                    selectedTargetRequirementIds.Remove(targetRequirement.Id);
                }
            }

            if (currentSourceRequirement != null)
            {
                SetMappingState(
                    currentSourceRequirement.Id,
                    targetRequirement.Id,
                    e.NewValue == CheckState.Checked);
                RenderSourceRequirements(currentSourceRequirement.Id);
            }
            UpdateProgressLabel();
        }

        private void TxtSearchTarget_TextChanged(object sender, EventArgs e)
        {
            if (suppressSearchRefresh)
            {
                return;
            }

            RenderTargetRequirements();
        }

        private void TxtSearchTarget_Enter(object sender, EventArgs e)
        {
            if (!IsSearchPlaceholderActive())
            {
                return;
            }

            suppressSearchRefresh = true;
            txtSearchTarget.Text = string.Empty;
            txtSearchTarget.ForeColor = SystemColors.WindowText;
            suppressSearchRefresh = false;
        }

        private void TxtSearchTarget_Leave(object sender, EventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(txtSearchTarget.Text))
            {
                return;
            }

            ResetSearchPlaceholder();
        }

        private void RenderSourceRequirements(string preferredSourceRequirementId)
        {
            suppressSourceSelection = true;
            lstSourceReqs.BeginUpdate();
            try
            {
                lstSourceReqs.Items.Clear();

                IReadOnlyList<RequirementItem> requirements = sourceSnapshot?.Requirements ?? new List<RequirementItem>();
                foreach (RequirementItem requirement in requirements.Where(item => item != null))
                {
                    requirement.IsMapped = HasMappedTargets(requirement.Id);
                    lstSourceReqs.Items.Add(new SourceRequirementListEntry(requirement));
                }

                if (lstSourceReqs.Items.Count > 0)
                {
                    int selectedIndex = 0;
                    if (!string.IsNullOrWhiteSpace(preferredSourceRequirementId))
                    {
                        for (int i = 0; i < lstSourceReqs.Items.Count; i++)
                        {
                            SourceRequirementListEntry entry = lstSourceReqs.Items[i] as SourceRequirementListEntry;
                            if (string.Equals(entry?.Requirement?.Id, preferredSourceRequirementId, StringComparison.OrdinalIgnoreCase))
                            {
                                selectedIndex = i;
                                break;
                            }
                        }
                    }

                    lstSourceReqs.SelectedIndex = selectedIndex;
                }
            }
            finally
            {
                lstSourceReqs.EndUpdate();
                suppressSourceSelection = false;
            }
        }

        private void RenderTargetRequirements()
        {
            visibleTargetRequirements.Clear();
            suppressTargetChecks = true;
            chkListTargetReqs.BeginUpdate();
            try
            {
                chkListTargetReqs.Items.Clear();

                RequirementItem currentSourceRequirement = GetSelectedSourceRequirement();
                IEnumerable<RequirementItem> filteredTargets = (IEnumerable<RequirementItem>)(targetSnapshot?.Requirements ?? new List<RequirementItem>())
                    .Where(item => item != null)
                    .Where(MatchesTargetFilter);

                foreach (RequirementItem targetRequirement in filteredTargets)
                {
                    int itemIndex = chkListTargetReqs.Items.Add(targetRequirement);
                    visibleTargetRequirements.Add(targetRequirement);
                    bool shouldCheck = !string.IsNullOrWhiteSpace(targetRequirement.Id) &&
                                       (selectedTargetRequirementIds.Contains(targetRequirement.Id) ||
                                        (currentSourceRequirement != null && IsMappedToCurrentSource(currentSourceRequirement.Id, targetRequirement.Id)));
                    chkListTargetReqs.SetItemChecked(itemIndex, shouldCheck);
                }
            }
            finally
            {
                chkListTargetReqs.EndUpdate();
                suppressTargetChecks = false;
            }
        }

        private void UpdateHeaders()
        {
            int targetCount = targetSnapshot?.Requirements?.Count ?? 0;
            lblSourceHeader.Text = "章节号 / 标识 / 名称";
            lblTargetHeader.Text = targetCount > 0 ? $"章节号 / 标识 / 名称（{targetCount}）" : "章节号 / 标识 / 名称";
        }

        private void UpdateProgressLabel()
        {
            int totalRequirements = sourceSnapshot?.Requirements?.Count ?? 0;
            int trackedRequirements = selectedTargetRequirementIds.Count;
            lblProgress.Text = $"已提取需求：{trackedRequirements}/{totalRequirements}";
        }

        private void ReportLoadProgress(string message)
        {
            lblProgress.Text = string.IsNullOrWhiteSpace(message) ? "正在解析需求数据..." : message;
            lblProgress.Refresh();
            Application.DoEvents();
        }

        private void BindDocumentOptions(ComboBox comboBox, string preferredFullName)
        {
            comboBox.BeginUpdate();
            try
            {
                comboBox.Items.Clear();
                foreach (RequirementTrackingDocumentOption option in openDocumentOptions)
                {
                    comboBox.Items.Add(option);
                }
            }
            finally
            {
                comboBox.EndUpdate();
            }

            if (comboBox.Items.Count == 0)
            {
                return;
            }

            if (!string.IsNullOrWhiteSpace(preferredFullName))
            {
                for (int i = 0; i < comboBox.Items.Count; i++)
                {
                    RequirementTrackingDocumentOption option = comboBox.Items[i] as RequirementTrackingDocumentOption;
                    if (option != null && string.Equals(option.FullName, preferredFullName, StringComparison.OrdinalIgnoreCase))
                    {
                        comboBox.SelectedIndex = i;
                        return;
                    }
                }
            }

            if (comboBox.SelectedIndex < 0)
            {
                comboBox.SelectedIndex = 0;
            }
        }

        private IReadOnlyList<RequirementTrackingDocumentOption> BuildDocumentOptions()
        {
            Dictionary<string, RequirementTrackingDocumentOption> optionMap =
                new Dictionary<string, RequirementTrackingDocumentOption>(StringComparer.OrdinalIgnoreCase);
            string activeDocumentFullName = GetActiveDocumentFullName();

            foreach (DocumentGroupDocumentItem document in GetActiveGroupDocuments())
            {
                string fullPath = NormalizeFilePath(document?.FilePath);
                if (string.IsNullOrWhiteSpace(fullPath))
                {
                    continue;
                }

                RequirementTrackingDocumentOption option;
                if (!optionMap.TryGetValue(fullPath, out option))
                {
                    option = new RequirementTrackingDocumentOption
                    {
                        FullName = fullPath,
                        DisplayName = Path.GetFileName(fullPath)
                    };
                    optionMap[fullPath] = option;
                }

                option.IsFromActiveGroup = true;
                if (string.IsNullOrWhiteSpace(option.DisplayName))
                {
                    option.DisplayName = Path.GetFileName(fullPath);
                }
            }

            if (!string.IsNullOrWhiteSpace(activeDocumentFullName))
            {
                RequirementTrackingDocumentOption option;
                if (!optionMap.TryGetValue(activeDocumentFullName, out option))
                {
                    option = new RequirementTrackingDocumentOption
                    {
                        FullName = activeDocumentFullName,
                        DisplayName = Path.GetFileName(activeDocumentFullName)
                    };
                    optionMap[activeDocumentFullName] = option;
                }

                option.IsCurrentlyOpen = true;
            }

            foreach (string manuallySelectedPath in manuallySelectedDocumentPaths)
            {
                string fullPath = NormalizeFilePath(manuallySelectedPath);
                if (string.IsNullOrWhiteSpace(fullPath))
                {
                    continue;
                }

                RequirementTrackingDocumentOption option;
                if (!optionMap.TryGetValue(fullPath, out option))
                {
                    option = new RequirementTrackingDocumentOption
                    {
                        FullName = fullPath,
                        DisplayName = Path.GetFileName(fullPath)
                    };
                    optionMap[fullPath] = option;
                }
            }

            return optionMap.Values
                .OrderByDescending(item => item.IsFromActiveGroup)
                .ThenByDescending(item => item.IsCurrentlyOpen)
                .ThenBy(item => item.DisplayName ?? string.Empty, StringComparer.CurrentCultureIgnoreCase)
                .ToList();
        }

        private void ApplyPreferredDocumentSelections(string previousSourceFullName, string previousTargetFullName)
        {
            if (string.IsNullOrWhiteSpace(previousSourceFullName))
            {
                RequirementTrackingDocumentOption preferredSource = FindPreferredDocument("任务书", "sss", "系统/子系统规格说明");
                if (preferredSource == null && openDocumentOptions.Count > 0)
                {
                    preferredSource = openDocumentOptions[0];
                }

                SelectDocumentOption(cmbSourceDoc, preferredSource);
            }

            if (string.IsNullOrWhiteSpace(previousTargetFullName))
            {
                RequirementTrackingDocumentOption preferredTarget = FindCurrentlyOpenDocument();
                if (preferredTarget == null)
                {
                    preferredTarget = FindPreferredDocument("需求规格说明", "srs");
                }

                if (preferredTarget == null && openDocumentOptions.Count > 1)
                {
                    string sourceFullName = GetSelectedDocumentFullName(cmbSourceDoc);
                    preferredTarget = openDocumentOptions.FirstOrDefault(option =>
                        !string.Equals(option.FullName, sourceFullName, StringComparison.OrdinalIgnoreCase));
                }

                if (preferredTarget == null && openDocumentOptions.Count > 0)
                {
                    preferredTarget = openDocumentOptions[0];
                }

                SelectDocumentOption(cmbTargetDoc, preferredTarget);
            }
        }

        private RequirementTrackingDocumentOption FindPreferredDocument(params string[] keywords)
        {
            return openDocumentOptions
                .OrderByDescending(option => option?.IsFromActiveGroup ?? false)
                .ThenByDescending(option => option?.IsCurrentlyOpen ?? false)
                .FirstOrDefault(option =>
            {
                string text = $"{option?.DisplayName} {option?.FullName}";
                return keywords.Any(keyword =>
                    !string.IsNullOrWhiteSpace(keyword) &&
                    text.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0);
            });
        }

        private RequirementTrackingDocumentOption FindCurrentlyOpenDocument()
        {
            string activeDocumentFullName = GetActiveDocumentFullName();
            if (string.IsNullOrWhiteSpace(activeDocumentFullName))
            {
                return null;
            }

            return openDocumentOptions.FirstOrDefault(option =>
                string.Equals(option?.FullName, activeDocumentFullName, StringComparison.OrdinalIgnoreCase));
        }

        private void SelectDocumentOption(ComboBox comboBox, RequirementTrackingDocumentOption option)
        {
            if (comboBox == null || option == null)
            {
                if (comboBox != null && comboBox.Items.Count > 0 && comboBox.SelectedIndex < 0)
                {
                    comboBox.SelectedIndex = 0;
                }

                return;
            }

            for (int i = 0; i < comboBox.Items.Count; i++)
            {
                RequirementTrackingDocumentOption current = comboBox.Items[i] as RequirementTrackingDocumentOption;
                if (current != null && string.Equals(current.FullName, option.FullName, StringComparison.OrdinalIgnoreCase))
                {
                    comboBox.SelectedIndex = i;
                    return;
                }
            }
        }

        private int FindSourceRequirementIndex(string sourceRequirementId)
        {
            if (string.IsNullOrWhiteSpace(sourceRequirementId))
            {
                return -1;
            }

            for (int i = 0; i < lstSourceReqs.Items.Count; i++)
            {
                SourceRequirementListEntry entry = lstSourceReqs.Items[i] as SourceRequirementListEntry;
                if (entry?.Requirement == null)
                {
                    continue;
                }

                if (string.Equals(entry.Requirement.Id, sourceRequirementId, StringComparison.OrdinalIgnoreCase))
                {
                    return i;
                }
            }

            return -1;
        }

        private RequirementItem GetSelectedSourceRequirement()
        {
            SourceRequirementListEntry entry = lstSourceReqs.SelectedItem as SourceRequirementListEntry;
            return entry?.Requirement;
        }

        private bool MatchesTargetFilter(RequirementItem requirement)
        {
            if (requirement == null)
            {
                return false;
            }

            string searchText = GetEffectiveSearchText();
            if (string.IsNullOrWhiteSpace(searchText))
            {
                return true;
            }

            return requirement.Id?.IndexOf(searchText, StringComparison.OrdinalIgnoreCase) >= 0 ||
                   requirement.Name?.IndexOf(searchText, StringComparison.OrdinalIgnoreCase) >= 0 ||
                   requirement.SectionNumber?.IndexOf(searchText, StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private void SetMappingState(string sourceRequirementId, string targetRequirementId, bool isMapped)
        {
            if (string.IsNullOrWhiteSpace(sourceRequirementId) || string.IsNullOrWhiteSpace(targetRequirementId))
            {
                return;
            }

            RequirementTraceMapping mapping;
            if (!mappingsBySourceId.TryGetValue(sourceRequirementId, out mapping))
            {
                if (!isMapped)
                {
                    return;
                }

                mapping = new RequirementTraceMapping
                {
                    SourceRequirementId = sourceRequirementId
                };
                mappingsBySourceId[sourceRequirementId] = mapping;
            }

            bool alreadyMapped = mapping.TargetRequirementIds.Any(item =>
                string.Equals(item, targetRequirementId, StringComparison.OrdinalIgnoreCase));

            if (isMapped && !alreadyMapped)
            {
                mapping.TargetRequirementIds.Add(targetRequirementId);
            }
            else if (!isMapped && alreadyMapped)
            {
                mapping.TargetRequirementIds = mapping.TargetRequirementIds
                    .Where(item => !string.Equals(item, targetRequirementId, StringComparison.OrdinalIgnoreCase))
                    .ToList();
            }

            if (mapping.TargetRequirementIds.Count == 0)
            {
                mappingsBySourceId.Remove(sourceRequirementId);
            }
        }

        private bool IsMappedToCurrentSource(string sourceRequirementId, string targetRequirementId)
        {
            if (string.IsNullOrWhiteSpace(sourceRequirementId) || string.IsNullOrWhiteSpace(targetRequirementId))
            {
                return false;
            }

            RequirementTraceMapping mapping;
            if (!mappingsBySourceId.TryGetValue(sourceRequirementId, out mapping))
            {
                return false;
            }

            return mapping.TargetRequirementIds.Any(item =>
                string.Equals(item, targetRequirementId, StringComparison.OrdinalIgnoreCase));
        }

        private bool HasMappedTargets(string sourceRequirementId)
        {
            if (string.IsNullOrWhiteSpace(sourceRequirementId))
            {
                return false;
            }

            RequirementTraceMapping mapping;
            if (!mappingsBySourceId.TryGetValue(sourceRequirementId, out mapping))
            {
                return false;
            }

            return mapping.TargetRequirementIds.Any(item => !string.IsNullOrWhiteSpace(item));
        }

        private string GetSelectedDocumentFullName(ComboBox comboBox)
        {
            RequirementTrackingDocumentOption option = comboBox?.SelectedItem as RequirementTrackingDocumentOption;
            return option?.FullName ?? string.Empty;
        }

        private string GetEffectiveSearchText()
        {
            if (IsSearchPlaceholderActive())
            {
                return string.Empty;
            }

            return (txtSearchTarget.Text ?? string.Empty).Trim();
        }

        private bool IsSearchPlaceholderActive()
        {
            return string.Equals(txtSearchTarget.Text, SearchPlaceholderText, StringComparison.Ordinal);
        }

        private void ResetSearchPlaceholder()
        {
            suppressSearchRefresh = true;
            txtSearchTarget.Text = SearchPlaceholderText;
            txtSearchTarget.ForeColor = Color.Gray;
            suppressSearchRefresh = false;
        }

        private Word.Application GetApplication()
        {
            return applicationAccessor?.Invoke();
        }

        private string GetActiveDocumentFullName()
        {
            try
            {
                Word.Document activeDocument = GetApplication()?.ActiveDocument;
                return NormalizeFilePath(activeDocument?.FullName);
            }
            catch
            {
                return string.Empty;
            }
        }

        private void BrowseAndSelectDocument(ComboBox targetComboBox)
        {
            using (OpenFileDialog dialog = new OpenFileDialog())
            {
                dialog.Title = "选择 Word 文档";
                dialog.Filter = "文档文件|*.docx;*.doc;*.docm;*.wps;*.wpt|Word 文档|*.docx;*.doc;*.docm|WPS 文档|*.wps;*.wpt|所有文件|*.*";
                dialog.FilterIndex = 1;
                dialog.Multiselect = false;
                dialog.CheckFileExists = true;
                dialog.CheckPathExists = true;
                dialog.RestoreDirectory = true;

                DialogResult result = ShowFileDialog(dialog);
                if (result != DialogResult.OK || string.IsNullOrWhiteSpace(dialog.FileName))
                {
                    return;
                }

                string selectedPath = NormalizeFilePath(dialog.FileName);
                if (string.IsNullOrWhiteSpace(selectedPath))
                {
                    return;
                }

                EnsureDocumentOptionExists(selectedPath);
                openDocumentOptions.Clear();
                openDocumentOptions.AddRange(BuildDocumentOptions());
                EnsureDocumentOptionExists(selectedPath);
                BindDocumentOptions(targetComboBox, selectedPath);
                SelectDocumentByFullName(targetComboBox, selectedPath);
            }
        }

        private DialogResult ShowFileDialog(FileDialog dialog)
        {
            if (dialog == null)
            {
                return DialogResult.Cancel;
            }

            try
            {
                return dialog.ShowDialog(this);
            }
            catch
            {
                return dialog.ShowDialog();
            }
        }

        private Word.Document OpenOrResolveDocument(string fullName)
        {
            if (string.IsNullOrWhiteSpace(fullName))
            {
                return null;
            }

            Word.Document matchedDocument = ResolveOpenDocument(fullName);
            if (matchedDocument != null)
            {
                return matchedDocument;
            }

            Word.Application app = GetApplication();
            if (app == null)
            {
                return null;
            }

            try
            {
                return app.Documents.Open(fullName, ReadOnly: true, Visible: false);
            }
            catch
            {
                return null;
            }
        }

        private Word.Document ResolveOpenDocument(string fullName)
        {
            if (string.IsNullOrWhiteSpace(fullName))
            {
                return null;
            }

            Word.Application app = GetApplication();
            if (app == null)
            {
                return null;
            }

            Word.Documents documents = null;
            Word.Document matchedDocument = null;
            try
            {
                documents = app.Documents;
                if (documents == null)
                {
                    return null;
                }

                int count = documents.Count;
                for (int i = 1; i <= count; i++)
                {
                    Word.Document doc = null;
                    try
                    {
                        doc = documents[i];
                        string currentFullName = string.Empty;
                        try
                        {
                            currentFullName = doc?.FullName ?? string.Empty;
                        }
                        catch
                        {
                            currentFullName = string.Empty;
                        }

                        if (!string.Equals(currentFullName, fullName, StringComparison.OrdinalIgnoreCase))
                        {
                            continue;
                        }

                        matchedDocument = doc;
                        doc = null;
                        break;
                    }
                    finally
                    {
                        ReleaseComObject(doc);
                    }
                }
            }
            finally
            {
                ReleaseComObject(documents);
            }

            return matchedDocument;
        }

        private IEnumerable<DocumentGroupDocumentItem> GetActiveGroupDocuments()
        {
            DocumentGroupStore store = new DocumentGroupStore();
            DocumentGroupCatalog catalog = store.Load();
            store.RefreshDocumentMetadata(catalog);
            DocumentGroupItem activeGroup = catalog.GetActiveGroup();
            return activeGroup?.Documents ?? Enumerable.Empty<DocumentGroupDocumentItem>();
        }

        private void EnsureDocumentOptionExists(string fullPath)
        {
            if (string.IsNullOrWhiteSpace(fullPath))
            {
                return;
            }

            if (!manuallySelectedDocumentPaths.Any(item => string.Equals(item, fullPath, StringComparison.OrdinalIgnoreCase)))
            {
                manuallySelectedDocumentPaths.Add(fullPath);
            }

            if (openDocumentOptions.Any(item => string.Equals(item.FullName, fullPath, StringComparison.OrdinalIgnoreCase)))
            {
                return;
            }

            openDocumentOptions.Add(new RequirementTrackingDocumentOption
            {
                FullName = fullPath,
                DisplayName = Path.GetFileName(fullPath)
            });
        }

        private void SelectDocumentByFullName(ComboBox comboBox, string fullPath)
        {
            if (comboBox == null || string.IsNullOrWhiteSpace(fullPath))
            {
                return;
            }

            for (int i = 0; i < comboBox.Items.Count; i++)
            {
                RequirementTrackingDocumentOption option = comboBox.Items[i] as RequirementTrackingDocumentOption;
                if (option == null)
                {
                    continue;
                }

                if (string.Equals(option.FullName, fullPath, StringComparison.OrdinalIgnoreCase))
                {
                    comboBox.SelectedIndex = i;
                    return;
                }
            }
        }

        private static string NormalizeFilePath(string filePath)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(filePath))
                {
                    return string.Empty;
                }

                return Path.GetFullPath(filePath.Trim());
            }
            catch
            {
                return string.Empty;
            }
        }

        private static Label CreateFieldLabel(string text)
        {
            return new Label
            {
                Dock = DockStyle.Fill,
                AutoEllipsis = false,
                Font = new Font("Microsoft YaHei UI", 8.5F, FontStyle.Regular, GraphicsUnit.Point),
                Text = text,
                TextAlign = ContentAlignment.MiddleLeft
            };
        }

        private static ComboBox CreateDocumentComboBox()
        {
            return new ComboBox
            {
                Dock = DockStyle.Fill,
                Font = new Font("Microsoft YaHei UI", 8.5F, FontStyle.Regular, GraphicsUnit.Point),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
        }

        private static Button CreateBrowseButton()
        {
            Button button = new Button
            {
                Dock = DockStyle.Fill,
                Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Regular, GraphicsUnit.Point),
                Text = "浏览..."
            };
            ApplySecondaryButtonStyle(button);
            return button;
        }

        private static Button CreateLoadButton(string text)
        {
            Button button = new Button
            {
                Dock = DockStyle.Fill,
                Font = new Font("Microsoft YaHei", 10F, FontStyle.Bold, GraphicsUnit.Point),
                Text = text,
                Height = 38,
                Margin = new Padding(3, 3, 3, 1)
            };
            ApplyPrimaryButtonStyle(button);
            return button;
        }

        private static void ApplyPrimaryButtonStyle(Button button)
        {
            if (button == null)
            {
                return;
            }

            button.UseVisualStyleBackColor = false;
            button.BackColor = Color.FromArgb(42, 122, 226);
            button.ForeColor = Color.White;
            button.FlatStyle = FlatStyle.Flat;
            button.FlatAppearance.BorderSize = 0;
            button.FlatAppearance.MouseOverBackColor = Color.FromArgb(58, 136, 236);
            button.FlatAppearance.MouseDownBackColor = Color.FromArgb(33, 105, 199);
            button.Cursor = Cursors.Hand;
        }

        private static void ApplySecondaryButtonStyle(Button button)
        {
            if (button == null)
            {
                return;
            }

            button.UseVisualStyleBackColor = false;
            button.BackColor = Color.FromArgb(235, 243, 255);
            button.ForeColor = Color.FromArgb(36, 89, 171);
            button.FlatStyle = FlatStyle.Flat;
            button.FlatAppearance.BorderColor = Color.FromArgb(176, 203, 240);
            button.FlatAppearance.BorderSize = 1;
            button.FlatAppearance.MouseOverBackColor = Color.FromArgb(223, 236, 255);
            button.FlatAppearance.MouseDownBackColor = Color.FromArgb(209, 227, 253);
            button.Cursor = Cursors.Hand;
        }

        private static Panel CreateBorderPanel()
        {
            return new Panel
            {
                Dock = DockStyle.Fill,
                BorderStyle = BorderStyle.FixedSingle,
                Padding = new Padding(8)
            };
        }

        private static void ReleaseComObject(object comObject)
        {
            if (comObject == null || !Marshal.IsComObject(comObject))
            {
                return;
            }

            try
            {
                Marshal.ReleaseComObject(comObject);
            }
            catch
            {
            }
        }

        private sealed class SourceRequirementListEntry
        {
            internal SourceRequirementListEntry(RequirementItem requirement)
            {
                Requirement = requirement;
            }

            internal RequirementItem Requirement { get; }

            public override string ToString()
            {
                if (Requirement == null) return string.Empty;
                return Requirement.DisplayText;
            }
        }

        private sealed class NativeWindowWrapper : IWin32Window
        {
            internal NativeWindowWrapper(IntPtr handle)
            {
                Handle = handle;
            }

            public IntPtr Handle { get; }
        }
    }
}
