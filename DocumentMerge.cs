using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2010.Word;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using PeakSystem.BusinessService;
using System;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;

namespace PeakSystem.DocumentMerge.Utility
{
    public class DocumentMerge
    {
        public DocumentMerge()
        {

        }

        private PeakSystem.BusinessService.FormTemplates _FormTemplatesBS = null;

        /// <summary>
        /// Read the bookmarks from the document and return the datatable
        /// Created date: 11/20/2018(Kanwaldeep Singh) 
        /// <userstory>586</userstory>
        /// </summary>
        /// <param name="FileName"></param>
        /// <returns></returns>
        public DataTable ReadWordDocumentBookMarks(string FileName)
        {
            DataTable _DataTableBookMarks = null;
            DocumentMerge FormTemplatesObject = null;
            System.Data.DataTable DataTableFormFields = null;
            _FormTemplatesBS = new PeakSystem.BusinessService.FormTemplates();
            if (FileName != null)
            {
                FormTemplatesObject = new DocumentMerge();
                DataTableFormFields = _FormTemplatesBS.GetFormFieldDetails(-1);
                DataRow[] DataRowFormFields;
                string FieldName = string.Empty;

                using (WordprocessingDocument document = WordprocessingDocument.Open(FileName, true))
                {
                    document.ChangeDocumentType(WordprocessingDocumentType.Document);
                    MainDocumentPart mainPart = document.MainDocumentPart;
                    var fields = mainPart.Document.Body.Descendants<FormFieldData>();
                    if (_DataTableBookMarks == null)
                    {
                        _DataTableBookMarks = new DataTable();
                        _DataTableBookMarks.Columns.Add("Bookmark Name");
                        _DataTableBookMarks.Columns.Add("Will merge with...");
                        _DataTableBookMarks.Columns.Add("FormEntityId");
                    }
                    DataRow DataRowBookMark;
                    foreach (var field in fields)
                    {
                        FieldName = ((FormFieldName)field.FirstChild).Val.InnerText;
                        DataRowFormFields = DataTableFormFields.Select("FormFieldName='" + FieldName + "'");
                        DataRowBookMark = _DataTableBookMarks.NewRow();
                        DataRowBookMark["Bookmark Name"] = FieldName;

                        if (DataTableFormFields.Rows.Count > 0 && DataRowFormFields.Length > 0)
                        {
                            DataRowBookMark["Will merge with..."] = DataRowFormFields[0].ItemArray[2].ToString();
                            DataRowBookMark["FormEntityId"] = DataRowFormFields[0]["FormEntityId"];
                        }
                        _DataTableBookMarks.Rows.Add(DataRowBookMark);
                    }

                    /// <summary>
                    /// Preview the content controls. 
                    /// </summary>
                    /// <userstory>1033</userstory>
                    #region Preview ContentControls
                    var sdtbFields = mainPart.Document.Body.Descendants<SdtAlias>();
                    foreach (var sdtb in sdtbFields)
                    {
                        FieldName = sdtb.Parent.ChildElements.OfType<Tag>().ElementAt(0).Val.InnerText;

                        DataRowFormFields = DataTableFormFields.Select("FormFieldName='" + FieldName + "'");
                        DataRowBookMark = _DataTableBookMarks.NewRow();
                        DataRowBookMark["Bookmark Name"] = FieldName;

                        if (DataTableFormFields.Rows.Count > 0 && DataRowFormFields.Length > 0)
                        {
                            DataRowBookMark["Will merge with..."] = DataRowFormFields[0].ItemArray[2].ToString();
                            DataRowBookMark["FormEntityId"] = DataRowFormFields[0]["FormEntityId"];
                        }
                        _DataTableBookMarks.Rows.Add(DataRowBookMark);
                    }
                    #endregion

                }

            }
            return _DataTableBookMarks;
        }

        /// <summary>
        /// Method merge the values in the document
        /// </summary>
        /// Created date: 11/20/2018(Kanwaldeep Singh) 
        /// <userstory>586</userstory>
        /// <param name="WordTemplateId"></param>
        /// <param name="FileName"></param>
        /// <param name="NewFullPath"></param>
        /// <param name="DefaultCompanyId"></param>
        /// <param name="TemplatePath"></param>
        /// <param name="XML"></param>
        public void ReadWordDocumentBookMarksAndMergeValues(Int32 WordTemplateId, ref string FileName, ref string NewFullPath, Int32 DefaultCompanyId, string TemplatePath, string XML)
        {
            Client _ClientsObject = null;
            try
            {
                _FormTemplatesBS = new PeakSystem.BusinessService.FormTemplates();
                _ClientsObject = new Client();
                DataTable DataTableWordTemplates = _ClientsObject.GetDataSource("WordTemplates", WordTemplateId, DefaultCompanyId, DefaultCompanyId, false);

                if (DataTableWordTemplates.Rows.Count == 1)
                {
                    FileName = DataTableWordTemplates.Rows[0]["TemplateFileName"].ToString();
                    NewFullPath = TemplatePath + "Temp/" + FileName;
                    if (DataTableWordTemplates.Rows[0]["TemplateType"].ToString().Equals("W"))
                    {
                        CreateWordDocumentFile(FileName, NewFullPath, XML);
                    }
                    else
                    {
                        _FormTemplatesBS.CreatePDFFile(FileName, NewFullPath, XML);
                    }
                }

            }
            catch (Exception Ex)
            {

                throw Ex;
            }
        }

        /// <summary>
        /// Merged the word document
        /// </summary>
        /// <userstory>586</userstory>
        /// <param name="fileName"></param>
        /// <param name="targetPath"></param>
        /// <param name="xmlData"></param>
        private void CreateWordDocumentFile(string fileName, string targetPath, string xmlData)
        {
            _FormTemplatesBS = new PeakSystem.BusinessService.FormTemplates();
            DataTable _DataTableFormFields = null;
            _DataTableFormFields = _FormTemplatesBS.GetFormFieldDetails(-1);
            DataRow[] DataRowFormFields;
            string FieldName = string.Empty;

            File.Copy(ConfigurationManager.AppSettings["TemplatesPath"] + fileName.ToString(), targetPath, true);
            using (WordprocessingDocument document = WordprocessingDocument.Open(targetPath, true))
            {
                document.ChangeDocumentType(DocumentFormat.OpenXml.WordprocessingDocumentType.Document);
                MainDocumentPart mainPart = document.MainDocumentPart;
                var fields = mainPart.Document.Body.Descendants<FormFieldData>();
                foreach (var field in fields)
                {
                    FieldName = ((FormFieldName)field.FirstChild).Val.InnerText;
                    DataRowFormFields = _DataTableFormFields.Select("FormFieldName='" + FieldName + "'");
                    if (_DataTableFormFields.Rows.Count > 0 && DataRowFormFields.Length > 0)
                    {
                        string FieldValue = _FormTemplatesBS.GetBookMarkValue(Convert.ToInt32(HttpContext.Current.Session["SelectedClient"]), Convert.ToInt32(DataRowFormFields[0]["FormQueryId"]), Convert.ToString(DataRowFormFields[0]["FormQueryFieldName"]), xmlData);

                        string fieldType = field.LastChild.GetType().Name;
                        switch (fieldType)
                        {
                            case "DropDownListFormField":
                                if (((FormFieldName)field.FirstChild).Val.InnerText.Equals(FieldName))
                                {
                                    DropDownListFormField dropdownList = field.Descendants<DropDownListFormField>().First();
                                    if (dropdownList != null)
                                    {
                                        int index = -1;
                                        var dropdownOptions = dropdownList.ToList();

                                        if (dropdownOptions != null)
                                        {
                                            index = dropdownOptions.FindIndex(child => (child as ListEntryFormField).Val == FieldValue);

                                            if (index > -1)
                                            {
                                                DefaultDropDownListItemIndex defaultDropDownListItemIndex1;
                                                defaultDropDownListItemIndex1 = new DefaultDropDownListItemIndex();
                                                defaultDropDownListItemIndex1.Val = index;
                                                dropdownList.DefaultDropDownListItemIndex = defaultDropDownListItemIndex1;
                                            }
                                        }
                                    }
                                }
                                break;
                            case "TextInput":
                                if (((FormFieldName)field.FirstChild).Val.InnerText.Equals(FieldName))
                                {
                                    TextInput text = field.Descendants<TextInput>().First();
                                    SetFormFieldValue(text, FieldValue);
                                }
                                break;
                            case "CheckBox":
                                if (((FormFieldName)field.FirstChild).Val.InnerText.Equals(FieldName))
                                {
                                    CheckBox checkBox = field.Descendants<CheckBox>().First();
                                  
                                    var checkBoxValue = checkBox.ChildElements.LastOrDefault();

                                    if (checkBoxValue != null)
                                        (checkBoxValue as DocumentFormat.OpenXml.Office2013.Word.OnOffType).Val = FieldValue == "Y" ? true : false;
                                }
                                break;
                        }
                    }
                }

                /// <summary>
                /// Merged the content controls in word 
                /// </summary>
                /// <userstory>1033</userstory>
                #region SetContentControls
                var sdtbFields = mainPart.Document.Body.Descendants<SdtAlias>();
                foreach (var sdtb in sdtbFields)
                {
                    FieldName = sdtb.Parent.ChildElements.OfType<Tag>().ElementAt(0).Val.InnerText;
                    DataRowFormFields = _DataTableFormFields.Select("FormFieldName='" + FieldName + "'");
                    if (_DataTableFormFields.Rows.Count > 0 && DataRowFormFields.Length > 0)
                    {
                        string FieldValue = _FormTemplatesBS.GetBookMarkValue(Convert.ToInt32(HttpContext.Current.Session["SelectedClient"]), Convert.ToInt32(DataRowFormFields[0]["FormQueryId"]), Convert.ToString(DataRowFormFields[0]["FormQueryFieldName"]), xmlData);

                        if (sdtb.Parent != null)
                        {
                            string fieldType = sdtb.Parent.LastChild.GetType().Name;

                            switch (fieldType)
                            {
                                case "SdtContentDropDownList":
                                case "SdtContentComboBox":
                                    if (sdtb.Parent.ChildElements.OfType<Tag>().ElementAt(0).Val.InnerText.Equals(FieldName))
                                    {
                                        SdtContentDropDownList ddl = sdtb.Parent.Descendants<SdtContentDropDownList>().FirstOrDefault();
                                        if (ddl != null)
                                        {
                                            int index = -1;
                                            var dropdownOptions = ddl.ToList();
                                            if (dropdownOptions != null)
                                            {
                                                var selectedValue = FieldValue;
                                                index = dropdownOptions.FindIndex(child => (child as ListItem).Value == selectedValue);
                                                if (index > -1)
                                                {
                                                    Text T = sdtb.Parent.Parent.Descendants<Text>().FirstOrDefault();
                                                    T.Text = selectedValue;
                                                }
                                            }
                                        }
                                    }
                                    break;
                                case "SdtPlaceholder":
                                case "SdtContentText":
                                case "ShowingPlaceholder":
                                    if (sdtb.Parent.ChildElements.OfType<Tag>().ElementAt(0).Val.InnerText.Equals(FieldName))
                                    {
                                        var text = sdtb.Parent.Parent.Descendants<Text>().FirstOrDefault();
                                        if (text != null)
                                        {
                                            text.Text = FieldValue;
                                        }
                                    }
                                    break;
                                case "SdtContentCheckBox":
                                    if (sdtb.Parent.ChildElements.OfType<Tag>().ElementAt(0).Val.InnerText.Equals(FieldName))
                                    {
                                        var chk = sdtb.Parent.Descendants<SdtContentCheckBox>().FirstOrDefault();
                                        if (chk != null)
                                        {
                                            if (FieldValue == "Y")
                                            {
                                                chk.Checked.Val = OnOffValues.One;
                                                Text T = sdtb.Parent.Parent.Descendants<Text>().FirstOrDefault();
                                                T.Text = "☒";
                                            }
                                            else
                                            {
                                                chk.Checked.Val = OnOffValues.Zero;
                                                Text T = sdtb.Parent.Parent.Descendants<Text>().FirstOrDefault();
                                                T.Text = "☐";
                                            }
                                        }
                                    }
                                    break;
                            }
                        }
                    }
                }

                #endregion

                document.Save();
                document.Close();
            }
        }

        /// <summary>
        /// Method read the boomarks and show the merged values.
        /// </summary>
        /// Created date: 11/20/2018(Kanwaldeep Singh) 
        /// <userstory>586</userstory>
        /// <param name="WordTemplateId"></param>
        /// <param name="DefaultCompanyId"></param>
        /// <param name="TemplatePath"></param>
        /// <param name="XML"></param>
        /// <returns></returns>
        public DataTable ReadWordDocumentBookMarksAndShowValues(Int32 WordTemplateId, Int32 DefaultCompanyId, string TemplatePath, string XML)
        {
            DataTable DataTableBookMarks = null;
            string FileName = string.Empty;
            System.Data.DataTable DataTableFormFields = null;
            Client _ClientsObject = null;
            _FormTemplatesBS = new PeakSystem.BusinessService.FormTemplates();

            try
            {
                _ClientsObject = new Client();
                DataTable DataTableWordTemplates = _ClientsObject.GetDataSource("WordTemplates", WordTemplateId, DefaultCompanyId, DefaultCompanyId, false);
                if (DataTableWordTemplates.Rows.Count == 1)
                {
                    FileName = DataTableWordTemplates.Rows[0]["TemplateFileName"].ToString();
                    DataTableFormFields = _FormTemplatesBS.GetFormFieldDetails(-1);
                    DataRow[] DataRowFormFields;
                    using (WordprocessingDocument document = WordprocessingDocument.Open(TemplatePath + FileName.ToString(), true))
                    {
                        document.ChangeDocumentType(WordprocessingDocumentType.Document);
                        MainDocumentPart mainPart = document.MainDocumentPart;
                        var fields = mainPart.Document.Body.Descendants<FormFieldData>();

                        if (DataTableBookMarks == null)
                        {
                            DataTableBookMarks = new DataTable();
                            DataTableBookMarks.Columns.Add("Bookmark Name");
                            DataTableBookMarks.Columns.Add("Merged Value");
                        }
                        DataRow DataRowBookMark;
                        string FieldName = string.Empty;
                        HttpContext.Current.Session["DataSetFormFieldValue"] = null;
                        foreach (var field in fields)
                        {
                            FieldName = ((FormFieldName)field.FirstChild).Val.InnerText;

                            DataRowFormFields = DataTableFormFields.Select("FormFieldName='" + FieldName + "'");
                            DataRowBookMark = DataTableBookMarks.NewRow();
                            DataRowBookMark["Bookmark Name"] = FieldName;

                            if (DataTableFormFields.Rows.Count > 0 && DataRowFormFields.Length > 0)
                            {
                                DataRowBookMark["Merged Value"] = _FormTemplatesBS.GetBookMarkValue(Convert.ToInt32(HttpContext.Current.Session["SelectedClient"]), Convert.ToInt32(DataRowFormFields[0]["FormQueryId"]), Convert.ToString(DataRowFormFields[0]["FormQueryFieldName"]), XML);
                            }
                            DataTableBookMarks.Rows.Add(DataRowBookMark);
                        }

                        /// <summary>
                        /// Preview the content controls. 
                        /// </summary>
                        /// <userstory>1033</userstory>
                        #region Preview ContentControls
                        var sdtbFields = mainPart.Document.Body.Descendants<SdtAlias>();
                        foreach (var sdtb in sdtbFields)
                        {
                            FieldName = sdtb.Parent.ChildElements.OfType<Tag>().ElementAt(0).Val.InnerText;

                            DataRowFormFields = DataTableFormFields.Select("FormFieldName='" + FieldName + "'");
                            DataRowBookMark = DataTableBookMarks.NewRow();
                            DataRowBookMark["Bookmark Name"] = FieldName;

                            if (DataTableFormFields.Rows.Count > 0 && DataRowFormFields.Length > 0)
                            {
                                DataRowBookMark["Merged Value"] = _FormTemplatesBS.GetBookMarkValue(Convert.ToInt32(HttpContext.Current.Session["SelectedClient"]), Convert.ToInt32(DataRowFormFields[0]["FormQueryId"]), Convert.ToString(DataRowFormFields[0]["FormQueryFieldName"]), XML);
                            }
                            DataTableBookMarks.Rows.Add(DataRowBookMark);
                        }
                        #endregion
                        document.Close();
                    }

                }
            }
            catch (Exception Ex)
            {
                throw Ex;
            }
            return DataTableBookMarks;
        }

        /// <summary>
        /// Set the value of the textbox in the document
        /// </summary>
        /// Created date: 11/20/2018(Kanwaldeep Singh) 
        /// <userstory>586</userstory>
        /// <param name="textInput"></param>
        /// <param name="value"></param>
        private static void SetFormFieldValue(TextInput textInput, string value)
        {
            if (value == null) // Reset formfield using default if set.
            {
                if (textInput.DefaultTextBoxFormFieldString != null && textInput.DefaultTextBoxFormFieldString.Val.HasValue)
                    value = textInput.DefaultTextBoxFormFieldString.Val.Value;
            }

            // Enforce max length.
            short maxLength = 0; // Unlimited
            if (textInput.MaxLength != null && textInput.MaxLength.Val.HasValue)
                maxLength = textInput.MaxLength.Val.Value;
            if (value != null && maxLength > 0 && value.Length > maxLength)
                value = value.Substring(0, maxLength);

            // Not enforcing TextBoxFormFieldType (read documentation...).
            // Just note that the Word instance may modify the value of a formfield when user leave it based on TextBoxFormFieldType and Format.
            // A curious example:
            // Type Number, format "# ##0,00".
            // Set value to "2016 was the warmest year ever, at least since 1999.".
            // Open the document and select the field then tab out of it.
            // Value now is "2 016 tht,tt" (the logic behind this escapes me).

            // Format value. (Only able to handle formfields with textboxformfieldtype regular.)
            //if (textInput.TextBoxFormFieldType != null
            //&& textInput.TextBoxFormFieldType.Val.HasValue
            ////&& textInput.TextBoxFormFieldType.Val.Value != TextBoxFormFieldValues.Regular
            //)
            //    throw new ApplicationException("SetFormField: Unsupported textboxformfieldtype, only regular is handled.\r\n" + textInput.Parent.OuterXml);
            if (!string.IsNullOrWhiteSpace(value)
            && textInput.Format != null
            && textInput.Format.Val.HasValue)
            {
                switch (textInput.Format.Val.Value)
                {
                    case "Uppercase":
                        value = value.ToUpperInvariant();
                        break;
                    case "Lowercase":
                        value = value.ToLowerInvariant();
                        break;
                    case "First capital":
                        value = value[0].ToString().ToUpperInvariant() + value.Substring(1);
                        break;
                    case "Title case":
                        value = System.Globalization.CultureInfo.InvariantCulture.TextInfo.ToTitleCase(value);
                        break;
                    default: // ignoring any other values (not supposed to be any)
                        break;
                }
            }

            // Find run containing "separate" fieldchar.
            Run rTextInput = textInput.Ancestors<Run>().FirstOrDefault();
            if (rTextInput == null) throw new ApplicationException("SetFormField: Did not find run containing textinput.\r\n" + textInput.Parent.OuterXml);
            Run rSeparate = rTextInput.ElementsAfter().FirstOrDefault(ru =>
               ru.GetType() == typeof(Run)
               && ru.Elements<FieldChar>().FirstOrDefault(fc =>
                  fc.FieldCharType == FieldCharValues.Separate)
                  != null) as Run;
            if (rSeparate == null) throw new ApplicationException("SetFormField: Did not find run containing separate.\r\n" + textInput.Parent.OuterXml);

            // Find run containg "end" fieldchar.
            Run rEnd = rTextInput.ElementsAfter().FirstOrDefault(ru =>
               ru.GetType() == typeof(Run)
               && ru.Elements<FieldChar>().FirstOrDefault(fc =>
                  fc.FieldCharType == FieldCharValues.End)
                  != null) as Run;
            if (rEnd == null) // Formfield value contains paragraph(s)
            {
                Paragraph p = rSeparate.Parent as Paragraph;
                Paragraph pEnd = p.ElementsAfter().FirstOrDefault(pa =>
                pa.GetType() == typeof(Paragraph)
                && pa.Elements<Run>().FirstOrDefault(ru =>
                   ru.Elements<FieldChar>().FirstOrDefault(fc =>
                      fc.FieldCharType == FieldCharValues.End)
                      != null)
                   != null) as Paragraph;
                if (pEnd == null) throw new ApplicationException("SetFormField: Did not find paragraph containing end.\r\n" + textInput.Parent.OuterXml);
                rEnd = pEnd.Elements<Run>().FirstOrDefault(ru =>
                   ru.Elements<FieldChar>().FirstOrDefault(fc =>
                      fc.FieldCharType == FieldCharValues.End)
                      != null);
            }

            // Remove any existing value.

            Run rFirst = rSeparate.NextSibling<Run>();
            if (rFirst == null || rFirst == rEnd)
            {
                RunProperties rPr = rTextInput.GetFirstChild<RunProperties>();
                if (rPr != null) rPr = rPr.CloneNode(true) as RunProperties;
                rFirst = rSeparate.InsertAfterSelf<Run>(new Run(new[] { rPr }));
            }
            rFirst.RemoveAllChildren<Text>();

            Run r = rFirst.NextSibling<Run>();
            while (r != rEnd)
            {
                if (r != null)
                {
                    r.Remove();
                    r = rFirst.NextSibling<Run>();
                }
                else // next paragraph
                {
                    Paragraph p = rFirst.Parent.NextSibling<Paragraph>();
                    if (p == null) throw new ApplicationException("SetFormField: Did not find next paragraph prior to or containing end.\r\n" + textInput.Parent.OuterXml);
                    r = p.GetFirstChild<Run>();
                    if (r == null)
                    {
                        // No runs left in paragraph, move other content to end of paragraph containing "separate" fieldchar.
                        p.Remove();
                        while (p.FirstChild != null)
                        {
                            OpenXmlElement oxe = p.FirstChild;
                            oxe.Remove();
                            if (oxe.GetType() == typeof(ParagraphProperties)) continue;
                            rSeparate.Parent.AppendChild(oxe);
                        }
                    }
                }
            }
            if (rEnd.Parent != rSeparate.Parent)
            {
                // Merge paragraph containing "end" fieldchar with paragraph containing "separate" fieldchar.
                Paragraph p = rEnd.Parent as Paragraph;
                p.Remove();
                while (p.FirstChild != null)
                {
                    OpenXmlElement oxe = p.FirstChild;
                    oxe.Remove();
                    if (oxe.GetType() == typeof(ParagraphProperties)) continue;
                    rSeparate.Parent.AppendChild(oxe);
                }
            }

            // Set new value.

            if (value != null)
            {
                // Word API use \v internally for newline and \r for para. We treat \v, \r\n, and \n as newline (Break).
                string[] lines = value.Replace("\r\n", "\n").Split(new char[] { '\v', '\n', '\r' });
                string line = lines[0];
                Text text = rFirst.AppendChild<Text>(new Text(line));
                if (line.StartsWith(" ") || line.EndsWith(" ")) text.SetAttribute(new OpenXmlAttribute("xml:space", null, "preserve"));
                for (int i = 1; i < lines.Length; i++)
                {
                    rFirst.AppendChild<Break>(new Break());
                    line = lines[i];
                    text = rFirst.AppendChild<Text>(new Text(lines[i]));
                    if (line.StartsWith(" ") || line.EndsWith(" ")) text.SetAttribute(new OpenXmlAttribute("xml:space", null, "preserve"));
                }
            }
            else
            { // An empty formfield of type textinput got char 8194 times 5 or maxlength if maxlength is in the range 1 to 4.
                short length = maxLength;
                if (length == 0 || length > 5) length = 5;
                rFirst.AppendChild(new Text(((char)8194).ToString()));
                r = rFirst;
                for (int i = 1; i < length; i++) r = r.InsertAfterSelf<Run>(r.CloneNode(true) as Run);
            }
        }

      
    }
}