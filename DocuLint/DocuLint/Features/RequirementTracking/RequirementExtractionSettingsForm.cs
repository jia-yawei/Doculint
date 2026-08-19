using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace DocuLint
{
    internal sealed class RequirementExtractionSettingsForm : Form
    {
        private readonly CheckBox systemSpecificationTemplateCheckBox;
        private readonly CheckBox requirementSpecificationTemplateCheckBox;
        private readonly CheckBox softwareDesignTemplateCheckBox;
        private readonly CheckBox useDefaultTemplatesCheckBox;
        private readonly CheckBox useCustomTemplatesCheckBox;
        private readonly TextBox templatesTextBox;
        private readonly FlowLayoutPanel templatesPanel;
        private readonly Label templateHintLabel;
        private readonly CheckBox limitPagesCheckBox;
        private readonly CheckBox scanFieldResultsCheckBox;
        private readonly NumericUpDown startPageInput;
        private readonly NumericUpDown endPageInput;
        private readonly NumericUpDown sectionLevelInput;
        private readonly CheckBox preserveForwardMappingsCheckBox;
        private bool updatingPresetTemplateSelection;

        internal RequirementExtractionSettingsForm(
            DocumentMarkerDocumentType currentDocumentType,
            DocumentMarkerDocumentType selectedTemplateType,
            bool useCustomTemplates,
            IEnumerable<string> templates,
            bool limitPages,
            int startPage,
            int endPage,
            int sectionTruncationLevel,
            bool scanFieldResults,
            bool preserveForwardMappingsWhenReverseTracing)
        {
            Text = "提取设置";
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            AutoScaleMode = AutoScaleMode.Dpi;
            AutoScaleDimensions = new SizeF(96F, 96F);
            ClientSize = new Size(960, 616);
            Font = new Font("Microsoft YaHei UI", 9F);

            TabControl settingsTabs = new TabControl
            {
                Dock = DockStyle.Fill,
                HotTrack = true,
                Padding = new Point(12, 5)
            };
            TabPage extractionSettingsTab = new TabPage("提取设置")
            {
                BackColor = Color.FromArgb(245, 245, 245),
                Padding = Padding.Empty
            };
            TabPage trackingSettingsTab = new TabPage("追踪设置")
            {
                BackColor = Color.White,
                Padding = Padding.Empty
            };

            TableLayoutPanel trackingLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 1,
                Padding = new Padding(20)
            };
            preserveForwardMappingsCheckBox = new CheckBox
            {
                AutoSize = true,
                Text = "反向追踪时保留正向追踪关系",
                Checked = preserveForwardMappingsWhenReverseTracing,
                Margin = Padding.Empty
            };
            trackingLayout.Controls.Add(preserveForwardMappingsCheckBox, 0, 0);
            trackingSettingsTab.Controls.Add(trackingLayout);

            TableLayoutPanel layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                AutoScroll = true,
                ColumnCount = 1,
                RowCount = 9,
                Padding = new Padding(20)
            };
            for (int index = 0; index < 8; index++)
            {
                layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            }
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            useDefaultTemplatesCheckBox = new CheckBox
            {
                AutoSize = true,
                Text = "使用默认规则",
                Font = new Font("Microsoft YaHei UI", 10F, FontStyle.Bold),
                Margin = new Padding(0, 0, 0, 6)
            };
            FlowLayoutPanel presetTemplatePanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                Margin = new Padding(0, 0, 0, 10)
            };
            systemSpecificationTemplateCheckBox = CreatePresetTemplateCheckBox(
                "任务书/系统规格说明（SSS、SDTD）",
                DocumentMarkerDocumentType.SystemSpecification);
            requirementSpecificationTemplateCheckBox = CreatePresetTemplateCheckBox(
                "需求规格说明（SRS）",
                DocumentMarkerDocumentType.RequirementSpecification);
            softwareDesignTemplateCheckBox = CreatePresetTemplateCheckBox(
                "设计说明（SDS、SDD）",
                DocumentMarkerDocumentType.SoftwareDesign);
            presetTemplatePanel.Controls.Add(systemSpecificationTemplateCheckBox);
            presetTemplatePanel.Controls.Add(requirementSpecificationTemplateCheckBox);
            presetTemplatePanel.Controls.Add(softwareDesignTemplateCheckBox);
            SelectPresetTemplate(selectedTemplateType == DocumentMarkerDocumentType.Unknown
                ? NormalizePresetTemplateType(currentDocumentType)
                : NormalizePresetTemplateType(selectedTemplateType));

            useCustomTemplatesCheckBox = new CheckBox
            {
                AutoSize = true,
                Text = "使用自定义规则",
                Checked = useCustomTemplates,
                Font = new Font("Microsoft YaHei UI", 10F, FontStyle.Bold),
                Margin = new Padding(0, 0, 0, 6)
            };
            useDefaultTemplatesCheckBox.Checked = !useCustomTemplates;
            templatesPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                Margin = new Padding(0, 0, 0, 4)
            };
            IReadOnlyList<string> normalizedTemplates = NormalizeTemplates(templates);
            templatesTextBox = new TextBox
            {
                Width = 900,
                Height = 30,
                Margin = new Padding(0, 0, 0, 6),
                Text = string.Join("; ", normalizedTemplates)
            };
            templatesPanel.Controls.Add(templatesTextBox);
            templateHintLabel = new Label
            {
                AutoSize = true,
                Text = "# 表示一个连字符之间的任意段，例如：Q2460-2153-SRS-#-#。最多 3 种，格式之间用分号分隔。",
                ForeColor = Color.FromArgb(96, 104, 116),
                Margin = new Padding(0, 0, 0, 14)
            };
            useDefaultTemplatesCheckBox.CheckedChanged += TemplateModeCheckBox_CheckedChanged;
            useCustomTemplatesCheckBox.CheckedChanged += TemplateModeCheckBox_CheckedChanged;
            UpdateCustomTemplateState();

            FlowLayoutPanel pageRow = CreateSettingsRow();
            limitPagesCheckBox = new CheckBox
            {
                AutoSize = true,
                Text = "限定页码范围",
                Checked = limitPages,
                Margin = new Padding(0, 6, 14, 0)
            };
            startPageInput = CreatePageInput(startPage);
            endPageInput = CreatePageInput(endPage);
            pageRow.Controls.Add(limitPagesCheckBox);
            pageRow.Controls.Add(startPageInput);
            pageRow.Controls.Add(CreateRowLabel("至"));
            pageRow.Controls.Add(endPageInput);
            pageRow.Controls.Add(CreateRowLabel("页"));
            limitPagesCheckBox.CheckedChanged += (_, __) => UpdatePageInputs();
            UpdatePageInputs();

            scanFieldResultsCheckBox = new CheckBox
            {
                AutoSize = true,
                Text = "扫描域结果（可能增加提取耗时）",
                Checked = scanFieldResults,
                Margin = new Padding(0, 0, 0, 14)
            };

            FlowLayoutPanel sectionRow = CreateSettingsRow();
            Label sectionLabel = CreateRowLabel("章节号截断级别");
            sectionLabel.Margin = new Padding(0, 6, 14, 0);
            sectionLevelInput = new NumericUpDown
            {
                Width = 96,
                Height = 32,
                Minimum = 0,
                Maximum = 9,
                Value = Math.Max(0, Math.Min(9, sectionTruncationLevel)),
                Margin = new Padding(0, 0, 12, 0)
            };
            Label sectionHint = CreateRowLabel("0 表示不截断");
            sectionHint.ForeColor = Color.FromArgb(96, 104, 116);
            sectionRow.Controls.Add(sectionLabel);
            sectionRow.Controls.Add(sectionLevelInput);
            sectionRow.Controls.Add(sectionHint);

            FlowLayoutPanel buttons = new FlowLayoutPanel
            {
                Dock = DockStyle.Bottom,
                FlowDirection = FlowDirection.RightToLeft,
                WrapContents = false,
                AutoSize = true,
                Margin = new Padding(0, 20, 0, 0)
            };
            Button confirmButton = CreateDialogButton("确定", DialogResult.OK);
            Button cancelButton = CreateDialogButton("取消", DialogResult.Cancel);
            buttons.Controls.Add(confirmButton);
            buttons.Controls.Add(cancelButton);

            layout.Controls.Add(useDefaultTemplatesCheckBox, 0, 0);
            layout.Controls.Add(presetTemplatePanel, 0, 1);
            layout.Controls.Add(useCustomTemplatesCheckBox, 0, 2);
            layout.Controls.Add(templatesPanel, 0, 3);
            layout.Controls.Add(templateHintLabel, 0, 4);
            layout.Controls.Add(pageRow, 0, 5);
            layout.Controls.Add(scanFieldResultsCheckBox, 0, 6);
            layout.Controls.Add(sectionRow, 0, 7);
            layout.Controls.Add(buttons, 0, 8);
            extractionSettingsTab.Controls.Add(layout);
            settingsTabs.TabPages.Add(extractionSettingsTab);
            settingsTabs.TabPages.Add(trackingSettingsTab);
            Controls.Add(settingsTabs);

            AcceptButton = confirmButton;
            CancelButton = cancelButton;
            FormClosing += RequirementExtractionSettingsForm_FormClosing;
        }

        internal DocumentMarkerDocumentType SelectedPresetTemplateType
        {
            get
            {
                if (systemSpecificationTemplateCheckBox.Checked)
                {
                    return DocumentMarkerDocumentType.SystemSpecification;
                }

                if (softwareDesignTemplateCheckBox.Checked)
                {
                    return DocumentMarkerDocumentType.SoftwareDesign;
                }

                return DocumentMarkerDocumentType.RequirementSpecification;
            }
        }

        internal bool UseCustomTemplates => useCustomTemplatesCheckBox.Checked;

        internal IReadOnlyList<string> Templates => ParseTemplates(templatesTextBox.Text);

        internal bool LimitPages => limitPagesCheckBox.Checked;

        internal bool ScanFieldResults => scanFieldResultsCheckBox.Checked;

        internal int StartPage => (int)startPageInput.Value;

        internal int EndPage => (int)endPageInput.Value;

        internal int SectionTruncationLevel => (int)sectionLevelInput.Value;

        internal bool PreserveForwardMappingsWhenReverseTracing =>
            preserveForwardMappingsCheckBox.Checked;

        private void UpdateCustomTemplateState()
        {
            bool enabled = useCustomTemplatesCheckBox.Checked;
            systemSpecificationTemplateCheckBox.Enabled = !enabled;
            requirementSpecificationTemplateCheckBox.Enabled = !enabled;
            softwareDesignTemplateCheckBox.Enabled = !enabled;
            templatesPanel.Enabled = enabled;
            templateHintLabel.Enabled = enabled;
            templatesTextBox.BackColor = enabled ? Color.White : Color.FromArgb(242, 244, 247);
        }

        private void TemplateModeCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (updatingPresetTemplateSelection)
            {
                return;
            }

            CheckBox changed = sender as CheckBox;
            if (changed == null)
            {
                return;
            }

            updatingPresetTemplateSelection = true;
            try
            {
                if (changed.Checked)
                {
                    if (ReferenceEquals(changed, useDefaultTemplatesCheckBox))
                    {
                        useCustomTemplatesCheckBox.Checked = false;
                    }
                    else
                    {
                        useDefaultTemplatesCheckBox.Checked = false;
                    }
                }
                else if (!useDefaultTemplatesCheckBox.Checked && !useCustomTemplatesCheckBox.Checked)
                {
                    changed.Checked = true;
                }
            }
            finally
            {
                updatingPresetTemplateSelection = false;
            }

            UpdateCustomTemplateState();
        }

        private void UpdatePageInputs()
        {
            startPageInput.Enabled = limitPagesCheckBox.Checked;
            endPageInput.Enabled = limitPagesCheckBox.Checked;
        }

        private void RequirementExtractionSettingsForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (DialogResult != DialogResult.OK)
            {
                return;
            }

            if (UseCustomTemplates && CountTemplateEntries(templatesTextBox.Text) > 3)
            {
                MessageBox.Show(this, "最多只能设置 3 条模板通配规则。", "提取设置");
                e.Cancel = true;
                return;
            }

            if (LimitPages && StartPage > EndPage)
            {
                MessageBox.Show(this, "起始页不能大于结束页。", "提取设置");
                e.Cancel = true;
            }
        }

        private void SelectPresetTemplate(DocumentMarkerDocumentType documentType)
        {
            updatingPresetTemplateSelection = true;
            try
            {
                systemSpecificationTemplateCheckBox.Checked = documentType == DocumentMarkerDocumentType.SystemSpecification;
                requirementSpecificationTemplateCheckBox.Checked = documentType == DocumentMarkerDocumentType.RequirementSpecification;
                softwareDesignTemplateCheckBox.Checked = documentType == DocumentMarkerDocumentType.SoftwareDesign;
            }
            finally
            {
                updatingPresetTemplateSelection = false;
            }
        }

        private CheckBox CreatePresetTemplateCheckBox(string text, DocumentMarkerDocumentType documentType)
        {
            CheckBox checkBox = new CheckBox
            {
                AutoSize = true,
                Text = text,
                Tag = documentType,
                Margin = new Padding(0, 0, 0, 5)
            };
            checkBox.CheckedChanged += PresetTemplateCheckBox_CheckedChanged;
            return checkBox;
        }

        private void PresetTemplateCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (updatingPresetTemplateSelection)
            {
                return;
            }

            CheckBox changed = sender as CheckBox;
            if (changed == null)
            {
                return;
            }

            updatingPresetTemplateSelection = true;
            try
            {
                if (changed.Checked)
                {
                    foreach (CheckBox checkBox in new[]
                    {
                        systemSpecificationTemplateCheckBox,
                        requirementSpecificationTemplateCheckBox,
                        softwareDesignTemplateCheckBox
                    })
                    {
                        if (!ReferenceEquals(checkBox, changed))
                        {
                            checkBox.Checked = false;
                        }
                    }
                }
                else if (!systemSpecificationTemplateCheckBox.Checked &&
                         !requirementSpecificationTemplateCheckBox.Checked &&
                         !softwareDesignTemplateCheckBox.Checked)
                {
                    changed.Checked = true;
                }
            }
            finally
            {
                updatingPresetTemplateSelection = false;
            }
        }

        private static DocumentMarkerDocumentType NormalizePresetTemplateType(DocumentMarkerDocumentType documentType)
        {
            switch (documentType)
            {
                case DocumentMarkerDocumentType.SystemSpecification:
                case DocumentMarkerDocumentType.RequirementSpecification:
                case DocumentMarkerDocumentType.SoftwareDesign:
                    return documentType;
                default:
                    return DocumentMarkerDocumentType.RequirementSpecification;
            }
        }

        private static FlowLayoutPanel CreateSettingsRow()
        {
            return new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                Margin = new Padding(0, 0, 0, 14)
            };
        }

        private static Label CreateRowLabel(string text)
        {
            return new Label
            {
                AutoSize = true,
                Text = text,
                TextAlign = ContentAlignment.MiddleLeft,
                Margin = new Padding(0, 6, 12, 0)
            };
        }

        private static NumericUpDown CreatePageInput(int value)
        {
            return new NumericUpDown
            {
                Width = 112,
                Height = 32,
                Minimum = 1,
                Maximum = 9999,
                Value = Math.Max(1, Math.Min(9999, value <= 0 ? 1 : value)),
                Margin = new Padding(0, 0, 12, 0)
            };
        }

        private static Button CreateDialogButton(string text, DialogResult dialogResult)
        {
            return new Button
            {
                Text = text,
                DialogResult = dialogResult,
                MinimumSize = new Size(88, 36),
                FlatStyle = FlatStyle.System,
                Margin = new Padding(8, 0, 0, 0)
            };
        }

        private static IReadOnlyList<string> NormalizeTemplates(IEnumerable<string> templates)
        {
            return (templates ?? Enumerable.Empty<string>())
                .Select(value => (value ?? string.Empty).Trim())
                .Where(value => value.Length > 0)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .Take(3)
                .ToList();
        }

        private static IReadOnlyList<string> ParseTemplates(string value)
        {
            return NormalizeTemplates((value ?? string.Empty)
                .Split(new[] { ';', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries));
        }

        private static int CountTemplateEntries(string value)
        {
            return (value ?? string.Empty)
                .Split(new[] { ';', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(item => item.Trim())
                .Where(item => item.Length > 0)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .Count();
        }

    }
}
