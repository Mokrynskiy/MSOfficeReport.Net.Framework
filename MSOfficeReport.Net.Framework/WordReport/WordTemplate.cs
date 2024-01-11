using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXmlPowerTools;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Table = DocumentFormat.OpenXml.Wordprocessing.Table;
using TableRow = DocumentFormat.OpenXml.Wordprocessing.TableRow;

namespace MSOfficeReport.Net.Framework.WordReport
{
    public sealed class WordTemplate
    {
        private string _path;
        private Dictionary<string, Object> _variables;
        private MemoryStream _ms;
        private WordprocessingDocument _document;
        private readonly Regex _regex = new Regex("\\{\\{.*?\\}\\}");
        private readonly Regex _tagRegex = new Regex("[\\{]{2}[a-zA-Z.]+[\\}]{2}");
        private readonly Regex _itemRegex = new Regex("Item");
        /// <summary>
        /// Отчет в формате MS Word
        /// </summary>
        /// <param name="path">Путь к шаблону Word</param>
        public WordTemplate(string path)
        {
            _path = path;
            _variables = new Dictionary<string, Object>();
        }
        
        /// <summary>
        /// Добавление переменной для заполнения шаблона
        /// </summary>
        /// <param name="name">Псевдоним</param>
        /// <param name="data">Объект данных</param>
        public void AddVariable(string name, object data)
        {
            _variables.Add(name, data);
        }

        /// <summary>
        /// Формирование отчета в формате MS Word на основе шаблона
        /// </summary>
        public void Generate()
        {
            try
            {
                byte[] byteArrey = File.ReadAllBytes(_path);
                MemoryStream ms = new MemoryStream();
                ms.Write(byteArrey, 0, byteArrey.Length);
                _document = WordprocessingDocument.Open(ms, true);
                SimplifyMarkupSettings settings = new SimplifyMarkupSettings
                {
                    RemoveComments = true,
                    RemoveContentControls = true,
                    RemoveEndAndFootNotes = true,
                    RemoveFieldCodes = false,
                    RemoveLastRenderedPageBreak = true,
                    RemovePermissions = true,
                    RemoveProof = true,
                    RemoveRsidInfo = true,
                    RemoveSmartTags = true,
                    RemoveSoftHyphens = true,
                    ReplaceTabsWithSpaces = true
                };
                MarkupSimplifier.SimplifyMarkup(_document, settings);
                _document.CleanRun();                
                var body = _document.MainDocumentPart.Document.Body;
                foreach (var txt in body.Descendants<Text>())
                {
                    if (_regex.IsMatch(txt.Text))
                    {
                        //string result = txt.Text;
                        foreach (var match in _tagRegex.Matches(txt.Text))
                        {
                            var t = match.ToString();
                            string[] splittedTag = t.Replace("{", "").Replace("}","").Split('.');
                            if(splittedTag.Length == 2)
                            {
                                var key = splittedTag[0];

                                if (key == "Item") continue;

                                var varName = splittedTag[1];
                                var variable = (from v in _variables where v.Key == key select v.Value).FirstOrDefault();
                                if (variable != null)
                                {
                                    Type type = variable.GetType();
                                    string value = type.GetProperty(varName)?.GetValue(variable)?.ToString();
                                    if (value != null)
                                    {
                                        if (value.Contains("rtf1"))
                                        {
                                            var altChunkId = varName;
                                            AlternativeFormatImportPart chunk = _document.MainDocumentPart.AddAlternativeFormatImportPart(AlternativeFormatImportPartType.Rtf, altChunkId);
                                            using (MemoryStream mst = new MemoryStream(Encoding.ASCII.GetBytes(value)))
                                            {
                                                chunk.FeedData(mst);                                                
                                            }
                                            AltChunk altChunk = new AltChunk();
                                            altChunk.Id = altChunkId;
                                            txt.InsertAfterSelf(altChunk);
                                            txt.Text = txt.Text.Replace(match.ToString(), "");
                                        }
                                        else txt.Text = txt.Text.Replace(match.ToString(), value);
                                    }
                                    else
                                    {
                                        txt.Text = txt.Text.Replace(match.ToString(), $"[{match.ToString()} - значение было null]");
                                    }
                                }
                            }
                        }
                    }
                }
                IEnumerable<TableProperties> tableProperties = body.Descendants<TableProperties>();
                foreach (var tableProperty in tableProperties)
                {
                    var tableCaption = tableProperty.TableCaption?.Val?.ToString();
                    if (tableCaption != null)
                    {
                        var variable = from v in _variables where v.Key == tableCaption select v;
                        if (variable.Any())
                        {
                            var data = variable.FirstOrDefault().Value;
                            var typ = data.GetType();
                            if (data != null) 
                            {
                                var listData = (IList)data;
                                var table = (Table)tableProperty.Parent;
                                TableRow clonedRow = null;
                                foreach (var row in table.Descendants<TableRow>())
                                {                                    
                                    var txt = from t in row.Descendants<Text>() where _regex.IsMatch(t.Text) select t;
                                    if (txt.Any())
                                    {
                                        if (_itemRegex.IsMatch(txt.FirstOrDefault().Text)) 
                                        {
                                            clonedRow = row;
                                            break;
                                        }
                                    }
                                }
                                if (clonedRow != null)
                                {
                                    var lastRowIndex = listData.Count -1;
                                    for (int i = 1; i < listData.Count; i++)
                                    {
                                        TableRow rowCopy = (TableRow)clonedRow.CloneNode(true);
                                        var rowData = listData[lastRowIndex];
                                        clonedRow.InsertAfterSelf(rowCopy);
                                        foreach (var txt in rowCopy.Descendants<Text>())
                                        {
                                            foreach (var match in _tagRegex.Matches(txt.Text))
                                            {
                                                var t = match.ToString();
                                                string[] splittedTag = t.Replace("{", "").Replace("}", "").Split('.');
                                                if (splittedTag.Count() == 2)
                                                {
                                                    var key = splittedTag[0];
                                                    if (key != "Item") continue;
                                                    var varName = splittedTag[1];

                                                    var type = rowData.GetType();
                                                    string value = type.GetProperty(varName)?.GetValue(rowData)?.ToString();
                                                    if (value != null) 
                                                        txt.Text = txt.Text.Replace(match.ToString(), value);
                                                    else
                                                        txt.Text = txt.Text.Replace(match.ToString(), $"[{match.ToString()} - значение было null]");
                                                }
                                            }
                                        }
                                        lastRowIndex--;
                                    }
                                    foreach (var txt in clonedRow.Descendants<Text>())
                                    {
                                        var rowData = listData[lastRowIndex];
                                        foreach (var match in _tagRegex.Matches(txt.Text))
                                        {
                                            var t = match.ToString();
                                            string[] splittedTag = t.Replace("{", "").Replace("}", "").Split('.');
                                            if (splittedTag.Count() == 2)
                                            {
                                                var key = splittedTag[0];
                                                if (key != "Item") continue;
                                                var varName = splittedTag[1];
                                                var type = rowData.GetType();
                                                string value = type.GetProperty(varName)?.GetValue(rowData)?.ToString();
                                                if (value != null)
                                                    txt.Text = txt.Text.Replace(match.ToString(), value);
                                                else
                                                    txt.Text = txt.Text.Replace(match.ToString(), $"[{match.ToString()} - значение было null]");
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _document.Dispose();                
                throw ex;
            }
        }

        /// <summary>
        /// Сохранение отчета
        /// </summary>
        /// <param name="outputFilePath">Путь сохранения</param>
        /// <returns></returns>
        public string SaveAs(string outputFilePath)
        {
            try
            {
                var doc = _document.Clone(outputFilePath);
                _document.Dispose();
                doc.Dispose();
                return outputFilePath;
            }
            catch (Exception ex)
            {
                _document.Dispose();                
                throw ex;
            }
        }
    }
}
