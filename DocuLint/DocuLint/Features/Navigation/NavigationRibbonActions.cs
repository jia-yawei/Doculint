using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using Word = Microsoft.Office.Interop.Word;

namespace DocuLint
{
    public partial class Ribbon1
    {
        private void button9_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Word.Application app = Globals.ThisAddIn.Application;
                Word.Document doc = app?.ActiveDocument;
                if (doc == null)
                {
                    MessageBox.Show("当前没有活动文档。", "文档不加班");
                    return;
                }

                Globals.ThisAddIn.ShowBookmarkListPane(doc, CollectBookmarkEntries(doc));
            }
            catch (Exception ex)
            {
                MessageBox.Show($"打开书签窗格失败: {ex.Message}", "文档不加班");
            }
        }

        private void button10_Click(object sender, RibbonControlEventArgs e)
        {
            btnListCaptions_Click(sender, e);
        }

        private void button11_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Word.Application app = Globals.ThisAddIn.Application;
                Word.Document doc = app?.ActiveDocument;
                if (doc == null)
                {
                    MessageBox.Show("当前没有活动文档。", "文档不加班");
                    return;
                }

                DocumentMarkerCollectionResult markerResult = DocumentMarkerService.CollectMarkers(doc);
                Globals.ThisAddIn.ShowMarkerListPane(doc, markerResult.Entries, markerResult.DocumentType);

                if (markerResult.Entries.Count == 0)
                {
                    MessageBox.Show("未识别到当前文档中的标识。", "文档不加班");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"打开标识窗格失败: {ex.Message}", "文档不加班");
            }
        }

        internal static List<NavigationPaneEntry> CollectBookmarkEntries(Word.Document doc)
        {
            List<NavigationPaneEntry> entries = new List<NavigationPaneEntry>();
            if (doc == null)
            {
                return entries;
            }

            try
            {
                foreach (Word.Bookmark bookmark in doc.Bookmarks)
                {
                    if (bookmark?.Range == null)
                    {
                        continue;
                    }

                    string name = (bookmark.Name ?? string.Empty).Trim();
                    if (string.IsNullOrWhiteSpace(name) || name.StartsWith("\\", StringComparison.Ordinal))
                    {
                        continue;
                    }

                    entries.Add(new NavigationPaneEntry
                    {
                        Start = bookmark.Range.Start,
                        Text = name
                    });
                }
            }
            catch
            {
            }

            return entries
                .OrderBy(entry => entry.Start)
                .GroupBy(entry => new { entry.Start, entry.Text })
                .Select(group => group.First())
                .ToList();
        }
    }
}
