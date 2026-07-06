using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Xml.Linq;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;

namespace DocuLint
{
    internal sealed class DocumentBasicInfoStore
    {
        private const string ProductCodePropertyName = "DocuLint_ProductCode";
        private const string ProductNamePropertyName = "DocuLint_ProductName";
        private const string RequirementPrefixPropertyName = "DocuLint_RequirementPrefix";
        private const string CodeLineCountPropertyName = "DocuLint_CodeLineCount";
        private const string PayloadPropertyName = "DocuLint_BasicInfoPayload";

        public DocumentBasicInfo Load(Word.Document doc)
        {
            if (doc == null)
            {
                return new DocumentBasicInfo();
            }

            object properties = GetCustomDocumentProperties(doc);
            DocumentBasicInfo payloadInfo = LoadFromPayload(properties);
            if (payloadInfo.HasAnyValue())
            {
                return payloadInfo;
            }

            return LoadLegacyFields(properties);
        }

        public void Save(Word.Document doc, DocumentBasicInfo info)
        {
            if (doc == null)
            {
                throw new InvalidOperationException("当前没有可配置的文档。");
            }

            DocumentBasicInfo safeInfo = Normalize(info);
            object properties = GetCustomDocumentProperties(doc);

            SetPropertyValue(properties, PayloadPropertyName, SerializeToPayload(safeInfo));
            SetLegacyMirrorValues(properties, safeInfo);
        }

        private static DocumentBasicInfo LoadLegacyFields(object properties)
        {
            List<DocumentBasicInfoField> fields = new List<DocumentBasicInfoField>
            {
                new DocumentBasicInfoField { Name = "产品标识", Value = GetPropertyValue(properties, ProductCodePropertyName) },
                new DocumentBasicInfoField { Name = "产品名称", Value = GetPropertyValue(properties, ProductNamePropertyName) },
                new DocumentBasicInfoField { Name = "需求前缀", Value = GetPropertyValue(properties, RequirementPrefixPropertyName) },
                new DocumentBasicInfoField { Name = "代码行数", Value = GetPropertyValue(properties, CodeLineCountPropertyName) }
            };

            return new DocumentBasicInfo
            {
                Fields = fields
                    .Where(field => field != null &&
                        (!string.IsNullOrWhiteSpace(field.Name) || !string.IsNullOrWhiteSpace(field.Value)))
                    .ToList()
            };
        }

        private static DocumentBasicInfo LoadFromPayload(object properties)
        {
            string payload = GetPropertyValue(properties, PayloadPropertyName);
            if (string.IsNullOrWhiteSpace(payload))
            {
                return new DocumentBasicInfo();
            }

            try
            {
                XDocument document = XDocument.Parse(payload);
                List<DocumentBasicInfoField> fields = document.Root?
                    .Elements("field")
                    .Select(element => new DocumentBasicInfoField
                    {
                        Name = (string)element.Attribute("name") ?? string.Empty,
                        Value = element.Value ?? string.Empty
                    })
                    .Where(field => !string.IsNullOrWhiteSpace(field.Name) || !string.IsNullOrWhiteSpace(field.Value))
                    .ToList()
                    ?? new List<DocumentBasicInfoField>();

                return new DocumentBasicInfo { Fields = fields };
            }
            catch
            {
                return new DocumentBasicInfo();
            }
        }

        private static string SerializeToPayload(DocumentBasicInfo info)
        {
            DocumentBasicInfo safeInfo = Normalize(info);
            XElement root = new XElement("basicInfo",
                safeInfo.Fields.Select(field =>
                    new XElement("field",
                        new XAttribute("name", field.Name ?? string.Empty),
                        field.Value ?? string.Empty)));

            return root.ToString(SaveOptions.DisableFormatting);
        }

        private static void SetLegacyMirrorValues(object properties, DocumentBasicInfo info)
        {
            string productCode = GetFieldValue(info, "产品标识");
            string productName = GetFieldValue(info, "产品名称");
            string requirementPrefix = GetFieldValue(info, "需求前缀");
            string codeLineCount = GetFieldValue(info, "代码行数");

            SetPropertyValue(properties, ProductCodePropertyName, productCode);
            SetPropertyValue(properties, ProductNamePropertyName, productName);
            SetPropertyValue(properties, RequirementPrefixPropertyName, requirementPrefix);
            SetPropertyValue(properties, CodeLineCountPropertyName, codeLineCount);
        }

        private static string GetFieldValue(DocumentBasicInfo info, string fieldName)
        {
            return Normalize(info).Fields
                .FirstOrDefault(field => string.Equals(field.Name, fieldName, StringComparison.CurrentCultureIgnoreCase))
                ?.Value ?? string.Empty;
        }

        private static DocumentBasicInfo Normalize(DocumentBasicInfo info)
        {
            DocumentBasicInfo safeInfo = info ?? new DocumentBasicInfo();
            return new DocumentBasicInfo
            {
                Fields = (safeInfo.Fields ?? new List<DocumentBasicInfoField>())
                    .Where(field => field != null &&
                        (!string.IsNullOrWhiteSpace(field.Name) || !string.IsNullOrWhiteSpace(field.Value)))
                    .Select(field => new DocumentBasicInfoField
                    {
                        Name = (field.Name ?? string.Empty).Trim(),
                        Value = (field.Value ?? string.Empty).Trim()
                    })
                    .ToList()
            };
        }

        private static object GetCustomDocumentProperties(Word.Document doc)
        {
            return doc.CustomDocumentProperties;
        }

        private static string GetPropertyValue(object properties, string propertyName)
        {
            if (properties == null || string.IsNullOrWhiteSpace(propertyName))
            {
                return string.Empty;
            }

            try
            {
                object property = properties.GetType().InvokeMember(
                    "Item",
                    BindingFlags.GetProperty,
                    null,
                    properties,
                    new object[] { propertyName });

                if (property == null)
                {
                    return string.Empty;
                }

                object value = property.GetType().InvokeMember(
                    "Value",
                    BindingFlags.GetProperty,
                    null,
                    property,
                    null);

                return Convert.ToString(value) ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private static void SetPropertyValue(object properties, string propertyName, string value)
        {
            if (properties == null || string.IsNullOrWhiteSpace(propertyName))
            {
                return;
            }

            string safeValue = (value ?? string.Empty).Trim();

            try
            {
                object property = properties.GetType().InvokeMember(
                    "Item",
                    BindingFlags.GetProperty,
                    null,
                    properties,
                    new object[] { propertyName });

                if (property != null)
                {
                    property.GetType().InvokeMember(
                        "Value",
                        BindingFlags.SetProperty,
                        null,
                        property,
                        new object[] { safeValue });
                    return;
                }
            }
            catch
            {
            }

            properties.GetType().InvokeMember(
                "Add",
                BindingFlags.InvokeMethod,
                null,
                properties,
                new object[]
                {
                    propertyName,
                    false,
                    Office.MsoDocProperties.msoPropertyTypeString,
                    safeValue
                });
        }
    }
}
