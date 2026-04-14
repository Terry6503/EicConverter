using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Linq;
using System.Xml;
using System.Text.RegularExpressions;
using System.Collections.Generic;

namespace EicConverter
{
    class Program
    {
        // 核心解密金鑰 (由 gobalvar.js 提取)
        private const string G_RIGHT_CODE = "267807C2DBAD8A9D03FCEB1D5E26408C0ED2EB56D4B5B784C428F3FD5B5028D0ACF267B6BDAFA8233FAC880DE62ADAAE1C9148065F3A73B395B02D24EBF6F99D";

        [STAThread]
        static void Main(string[] args)
        {
            Console.WriteLine("==================================================");
            Console.WriteLine("    公文轉換工具 (加密解鎖版) - x86             ");
            Console.WriteLine("==================================================");

            string currentDir = AppDomain.CurrentDomain.BaseDirectory;
            
            // 檢查核心元件是否存在 (根目錄 or bin 目錄)
            string binPath = Path.Combine(currentDir, "bin");
            bool hasHdf = File.Exists(Path.Combine(currentDir, "HDF32u.dll")) || File.Exists(Path.Combine(binPath, "HDF32u.dll"));
            bool hasGd = File.Exists(Path.Combine(currentDir, "GDLocal.dll")) || File.Exists(Path.Combine(binPath, "GDLocal.dll"));

            if (!hasHdf && !hasGd)
            {
                Console.WriteLine("警告: 找不到核心 DLL 元件 (HDF32u 或 GDLocal)，預定功能可能受限。");
            }

            Console.WriteLine("[執行] 掃描目錄: " + currentDir + "\n");

            string[] directories = Directory.GetDirectories(currentDir);
            if (directories.Length == 0)
            {
                Console.WriteLine("沒有找到任何子資料夾。");
                Console.ReadLine();
                return;
            }

            foreach (var dir in directories)
            {
                string folderOnlyName = Path.GetFileName(dir);
                if (folderOnlyName.Equals("crosscad", StringComparison.OrdinalIgnoreCase)) continue;
                if (folderOnlyName.Equals(".gemini", StringComparison.OrdinalIgnoreCase)) continue;
                if (folderOnlyName.Equals("eic", StringComparison.OrdinalIgnoreCase)) continue;

                try
                {
                    ProcessDirectory(dir, currentDir);
                }
                catch (Exception ex)
                {
                    Console.WriteLine(string.Format("[錯誤] 處理資料夾 {0} 時發生錯誤: {1}", Path.GetFileName(dir), ex.Message));
                }
            }

            Console.WriteLine("\n處理完成，請按任意鍵結束...");
            Console.ReadLine();
        }

        static void ProcessDirectory(string dirPath, string outputDir)
        {
            string folderName = Path.GetFileName(dirPath);
            string[] sdiFiles = Directory.GetFiles(dirPath, "*.sdi");

            if (sdiFiles.Length == 0) return;

            DirectoryInfo dirInfo = new DirectoryInfo(dirPath);
            string datePrefix = dirInfo.CreationTime.ToString("yyMMdd");

            object gdLocal = null;
            Type gdType = null;
            try 
            {
                // 1. 初始化 GDLocal 並解鎖 (整個資料夾共用一次)
                Guid gdClsid = new Guid("6635511E-BE10-11D4-8ACA-0080C8F15AE5");
                gdType = Type.GetTypeFromCLSID(gdClsid);
                if (gdType != null) 
                {
                    gdLocal = Activator.CreateInstance(gdType);
                    try { gdType.InvokeMember("Init", BindingFlags.InvokeMethod, null, gdLocal, new object[] { new DummyWindow(), G_RIGHT_CODE }); } catch { }
                    
                    // 2. 預掃描：找出資料夾內公文的「共用日期」
                    foreach (var sdi in sdiFiles) 
                    {
                        XmlDocument xmlDoc = GetXmlDocument(sdi, gdLocal, gdType);
                        if (xmlDoc != null) 
                        {
                            string foundDate = GetXmlDatePrefix(xmlDoc);
                            if (!string.IsNullOrEmpty(foundDate)) 
                            {
                                datePrefix = foundDate;
                                break; 
                            }
                        }
                    }
                }
            } 
            catch (Exception ex) { Console.WriteLine("       -> [初始化警告] " + ex.Message); }

            // 3. 執行轉換
            if (sdiFiles.Length == 1)
            {
                string targetSdi = sdiFiles[0];
                Console.WriteLine(string.Format("[找到公文] {0} -> 來源: {1}", folderName, Path.GetFileName(targetSdi)));
                ConvertViaCom(targetSdi, outputDir, folderName, datePrefix, gdLocal, gdType);
            }
            else
            {
                Console.WriteLine(string.Format("[找到公文] {0} -> 包含 {1} 個檔案，將進行同步日期處置 (Prefix: {2})", folderName, sdiFiles.Length, datePrefix));
                foreach (var sdiFile in sdiFiles)
                {
                    ConvertViaCom(sdiFile, outputDir, folderName, datePrefix, gdLocal, gdType, true);
                }
            }

            // 4. 釋放資源
            if (gdLocal != null) Marshal.ReleaseComObject(gdLocal);
        }

        static void ConvertViaCom(string sdiFile, string outputDir, string folderName, string fallbackPrefix, object gdLocal, Type gdType, bool multiMode = false)
        {
            string outTxt = Path.Combine(outputDir, string.Format("{0}_{1}.txt", fallbackPrefix, folderName));
            string outRtf = Path.Combine(outputDir, string.Format("{0}_{1}.rtf", fallbackPrefix, folderName));

            object dddViewObj = null;
            try
            {
                // Note: gdLocal is now passed from caller

                // 2. 啟動 DddView 引擎
                Guid clsid = new Guid("4747F455-5DB3-4F94-B15E-220C06F4776B");
                Type dddViewType = Type.GetTypeFromCLSID(clsid);

                if (dddViewType != null)
                {
                    dddViewObj = Activator.CreateInstance(dddViewType);
                    try { dddViewType.InvokeMember("DddViewCreate", BindingFlags.InvokeMethod, null, dddViewObj, null); } catch { }

                    try {
                        Console.WriteLine("       -> 正在執行 DddView.load...");
                        dddViewType.InvokeMember("load", BindingFlags.InvokeMethod, null, dddViewObj, new object[] { sdiFile });
                        
                        // 調用 preview() 以初始化內部緩衝區與頁面資訊 (此為瀏覽器行為模擬)
                        try { dddViewType.InvokeMember("preview", BindingFlags.InvokeMethod, null, dddViewObj, null); } catch { }

                        // 優先嘗試直接提取文字 (供備援使用)
                        string sContent = "";
                        try {
                            object result = dddViewType.InvokeMember("GetTxtTextFrDdd", BindingFlags.InvokeMethod, null, dddViewObj, null);
                            if (result != null) sContent = result.ToString();
                        } catch { }

                        if (!File.Exists(outTxt) || new FileInfo(outTxt).Length == 0) {
                            if (!File.Exists(outTxt)) File.WriteAllText(outTxt, "");
                            string shortOutTxt = GetShortPath(outTxt);
                            try { dddViewType.InvokeMember("GetTxtFileFrDdd", BindingFlags.InvokeMethod, null, dddViewObj, new object[] { shortOutTxt }); } 
                            catch { dddViewType.InvokeMember("docSaveTo", BindingFlags.InvokeMethod, null, dddViewObj, new object[] { shortOutTxt }); }
                        }

                        if (File.Exists(outTxt) && new FileInfo(outTxt).Length > 0 && string.IsNullOrEmpty(sContent)) {
                            sContent = File.ReadAllText(outTxt);
                        }

                        // 4. 結構化 RTF 建構 (語義化 XML 模式)
                        try {
                            if (File.Exists(outRtf)) File.Delete(outRtf);
                            Console.WriteLine("       -> 正在執行語義化轉檔 (XML 映射模式) -> " + Path.GetFileName(outRtf));
                            
                            // 獲取 XML 作為語義來源
                            XmlDocument xmlDoc = GetXmlDocument(sdiFile, gdLocal, gdType);
                            if (xmlDoc != null) {
                                string prefix = fallbackPrefix; // 強制使用已同步的日期前綴
                                
                                string suffix = "";
                                if (multiMode)
                                {
                                    string docType = "未知";
                                    XmlNode typeNode = GetNodeByLocalName(xmlDoc, "文別");
                                    if (typeNode != null && typeNode.Attributes["名稱"] != null) docType = typeNode.Attributes["名稱"].Value;
                                    
                                    if (docType.Contains("簽")) suffix = "_簽";
                                    else if (docType.Contains("函")) suffix = "_函";
                                    else if (docType.Contains("稿")) suffix = "_稿";
                                    else if (docType.Contains("說明")) suffix = "_說明";
                                    else suffix = "_" + docType;
                                }

                                string finalOutTxt = Path.Combine(outputDir, string.Format("{0}_{1}{2}.txt", prefix, folderName, suffix));
                                string finalOutRtf = Path.Combine(outputDir, string.Format("{0}_{1}{2}.rtf", prefix, folderName, suffix));

                                // 遷移或重新命名 TXT
                                if (outTxt != finalOutTxt)
                                {
                                    if (File.Exists(outTxt))
                                    {
                                        if (File.Exists(finalOutTxt)) File.Delete(finalOutTxt);
                                        File.Move(outTxt, finalOutTxt);
                                    }
                                    outTxt = finalOutTxt;
                                }
                                outRtf = finalOutRtf;
                                string finalOutDi = Path.Combine(outputDir, string.Format("{0}_{1}{2}.di", prefix, folderName, suffix));

                                Console.WriteLine("       -> 正在執行語義化轉檔 (XML 映射模式) -> " + Path.GetFileName(outRtf));
                                SaveSemanticRtf(outRtf, xmlDoc, sdiFile);
                                SaveSemanticTxt(outTxt, xmlDoc, sdiFile);
                                SaveDi(finalOutDi, xmlDoc, sdiFile);
                                
                                if (File.Exists(outTxt) && new FileInfo(outTxt).Length > 0) {
                                    Console.WriteLine("       -> 語義化轉檔完成 (Suffix: " + suffix + ")");
                                }
                            } else {
                                // 備援：如果抓不到 XML，則使用舊的文字解析邏輯
                                SaveStructuredRtf(outRtf, sContent, sdiFile);
                                Console.WriteLine("       -> 警告: 無法獲取 XML，已使用文字解析備援模式。");
                            }

                        } catch (Exception exRtf) {
                            Console.WriteLine("       -> RTF 語義化轉檔異常: " + exRtf.Message);
                        }

                        if (File.Exists(outTxt) && new FileInfo(outTxt).Length > 0) return; 
                    } catch (Exception exLoad) {
                        Console.WriteLine("       -> 引擎 A 異常 (切換引擎 B): " + exLoad.Message);
                    }
                }

                // 3. 備援引擎 B
                Console.WriteLine("       -> 嘗試使用引擎 B (手動解碼模式)...");
                TryManualRead(sdiFile, outTxt, gdLocal, gdType);
            }
            catch (Exception ex)
            {
                string msg = ex.Message;
                if (ex.InnerException != null) msg += " -> " + ex.InnerException.Message;
                Console.WriteLine(string.Format("       -> 處理錯誤: {0}", msg));
            }
            finally
            {
                if (dddViewObj != null) Marshal.ReleaseComObject(dddViewObj);
            }
        }

        [DllImport("kernel32.dll", CharSet = CharSet.Auto)]
        static extern int GetShortPathName(string lpszLongPath, System.Text.StringBuilder lpszShortPath, int cchBuffer);

        static string GetShortPath(string longPath)
        {
            try {
                string dir = Path.GetDirectoryName(longPath);
                if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);
                System.Text.StringBuilder sb = new System.Text.StringBuilder(255);
                int result = GetShortPathName(longPath, sb, sb.Capacity);
                return result > 0 ? sb.ToString() : longPath;
            } catch { return longPath; }
        }

        static void TryManualRead(string sdiFile, string outTxt, object gdLocal, Type gdType)
        {
            if (gdLocal == null || gdType == null) {
                Console.WriteLine("       -> 元件未就緒。");
                return;
            }

            try {
                Type xmlType = Type.GetTypeFromProgID("Msxml2.DOMDocument.6.0");
                if (xmlType == null) xmlType = Type.GetTypeFromProgID("Msxml2.DOMDocument");
                
                if (xmlType != null) {
                    object xmlDoc = Activator.CreateInstance(xmlType);
                    Console.WriteLine("       -> 正在透過 LoadXML 解密公文...");
                    bool success = (bool)gdType.InvokeMember("LoadXML", BindingFlags.InvokeMethod, null, gdLocal, new object[] { xmlDoc, sdiFile });
                    
                    if (success) {
                        string xmlContent = (string)xmlType.InvokeMember("xml", BindingFlags.GetProperty, null, xmlDoc, null);
                        if (!string.IsNullOrEmpty(xmlContent)) {
                            // Debug: Dump raw XML
                            File.WriteAllText(Path.Combine(Path.GetDirectoryName(outTxt), "raw_structure.xml"), xmlContent, System.Text.Encoding.UTF8);
                            
                            string extractedText = ExtractTextFromXml(xmlContent);
                            File.WriteAllText(outTxt, extractedText, System.Text.Encoding.UTF8);
                            Console.WriteLine("       -> 透過引擎 B (GDLocal 解密) 提取文字完成。");
                            return;
                        }
                    }
                }

                gdType.InvokeMember("Load", BindingFlags.InvokeMethod, null, gdLocal, new object[] { sdiFile });
                using (StreamWriter sw = new StreamWriter(outTxt)) {
                    while (true) {
                        bool isEof = (bool)gdType.InvokeMember("IsEOF", BindingFlags.InvokeMethod, null, gdLocal, null);
                        if (isEof) break;
                        string line = (string)gdType.InvokeMember("ReadLn", BindingFlags.InvokeMethod, null, gdLocal, null);
                        sw.WriteLine(line);
                    }
                }
                Console.WriteLine("       -> 透過引擎 B (流讀取) 提取完成。");
            } catch (Exception ex) {
                Console.WriteLine("       -> 引擎 B 失敗: " + ex.Message);
            }
        }

        static string ExtractTextFromXml(string xmlContent)
        {
            System.Text.StringBuilder sb = new System.Text.StringBuilder();
            try {
                string content = xmlContent.Replace("<P>", "\r\n").Replace("</P>", "\r\n")
                                           .Replace("<BR>", "\r\n")
                                           .Replace("<Paragraph>", "\r\n")
                                           .Replace("<Text>", "")
                                           .Replace("</Text>", "");

                bool inTag = false;
                foreach (char c in content) {
                    if (c == '<') inTag = true;
                    else if (c == '>') inTag = false;
                    else if (!inTag) sb.Append(c);
                }
            } catch { return xmlContent; }
            return sb.ToString().Replace("\r\n\r\n\r\n", "\r\n\r\n").Trim();
        }

        static string GetXmlDatePrefix(XmlDocument xmlDoc)
        {
            try {
                // 優先從 <年月日> 提取 (產製時間)
                XmlNode dateNode = GetNodeByLocalName(xmlDoc, "年月日");
                if (dateNode == null) dateNode = GetNodeByLocalName(xmlDoc, "發文日期");
                
                if (dateNode != null && dateNode.Attributes["年"] != null) {
                    string sYear = dateNode.Attributes["年"].Value;
                    string sMonth = dateNode.Attributes["月"] != null ? dateNode.Attributes["月"].Value : "1";
                    string sDay = dateNode.Attributes["日"] != null ? dateNode.Attributes["日"].Value : "1";
                    
                    int year = int.Parse(sYear);
                    int month = int.Parse(sMonth);
                    int day = int.Parse(sDay);
                    
                    // 格式化為 yyMMdd (西元後兩位)
                    DateTime dt = new DateTime(year > 1911 ? year : year + 1911, month, day);
                    return dt.ToString("yyMMdd");
                }
            } catch { }
            return null;
        }

        // --- 結構化轉檔輔助功能 ---

        static string GetMetadataValue(string xmlPath, string tagName)
        {
            try {
                if (!File.Exists(xmlPath)) return "";
                XmlDocument doc = new XmlDocument();
                doc.Load(xmlPath);
                XmlNode node = doc.SelectSingleNode("//" + tagName);
                return node != null ? node.InnerText : "";
            } catch { return ""; }
        }

        static XmlDocument GetXmlDocument(string sdiFile, object gdLocal, Type gdType)
        {
            if (gdLocal == null || gdType == null) return null;
            try {
                Type xmlType = Type.GetTypeFromProgID("Msxml2.DOMDocument.6.0");
                if (xmlType == null) xmlType = Type.GetTypeFromProgID("Msxml2.DOMDocument");
                if (xmlType != null) {
                    object xmlDocObj = Activator.CreateInstance(xmlType);
                    bool success = (bool)gdType.InvokeMember("LoadXML", BindingFlags.InvokeMethod, null, gdLocal, new object[] { xmlDocObj, sdiFile });
                    if (success) {
                        string xml = (string)xmlType.InvokeMember("xml", BindingFlags.GetProperty, null, xmlDocObj, null);
                        XmlDocument doc = new XmlDocument();
                        doc.LoadXml(xml);
                        return doc;
                    }
                }
            } catch { }
            return null;
        }

        static string FormatRocDate(string dateStr)
        {
            if (string.IsNullOrEmpty(dateStr)) return "";
            string[] parts = dateStr.Split('/');
            if (parts.Length == 3) {
                return string.Format("中華民國{0}年{1}月{2}日", parts[0], parts[1].TrimStart('0'), parts[2].TrimStart('0'));
            }
            return dateStr;
        }

        static string EncodeRtf(string text)
        {
            if (string.IsNullOrEmpty(text)) return "";
            System.Text.StringBuilder sb = new System.Text.StringBuilder();
            foreach (char c in text)
            {
                if (c < 128)
                {
                    if (c == '\\' || c == '{' || c == '}') sb.Append("\\" + c);
                    else sb.Append(c);
                }
                else
                {
                    short uniCode = (short)c;
                    sb.Append("\\u" + uniCode + "?");
                }
            }
            return sb.ToString();
        }

        static void SaveSemanticRtf(string outRtf, XmlDocument xmlDoc, string sdiFile)
        {
            string docType = "未知";
            string styleName = "";
            XmlNode typeNode = GetNodeByLocalName(xmlDoc, "文別");
            if (typeNode != null) {
                if (typeNode.Attributes["名稱"] != null) docType = typeNode.Attributes["名稱"].Value;
                if (typeNode.Attributes["樣式檔"] != null) styleName = typeNode.Attributes["樣式檔"].Value;
            }
            
            if (docType.Equals("便簽")) {
                SaveNoteRtf(outRtf, xmlDoc, sdiFile);
                return;
            }
            if (docType.Equals("開會通知單")) {
                SaveMeetingNoticeRtf(outRtf, xmlDoc, sdiFile);
                return;
            }
            if (docType.Equals("會勘通知單")) {
                SaveSiteInspectionNoticeRtf(outRtf, xmlDoc, sdiFile);
                return;
            }
            
            bool isExplanation = docType.Contains("說明") || 
                                 styleName.Contains("說明") || 
                                 sdiFile.Contains("說明") || 
                                 docType.IndexOf("A4", StringComparison.OrdinalIgnoreCase) >= 0;

            if (isExplanation) {
                SaveExplanationRtf(outRtf, xmlDoc, sdiFile);
                return;
            }

            if (docType.Equals("簽")) {
                SaveMemorandumRtf(outRtf, xmlDoc, sdiFile);
                return;
            }

            if (!docType.Equals("函")) {
                Console.WriteLine("       -> [提示] 檢測到文別為「" + docType + "」，將嘗試套用通用轉檔邏輯。");
            }

            // 1. 提取中繼資料
            XmlNode dateNode = GetNodeByLocalName(xmlDoc, "發文日期");
            string rocDate = "";
            if (dateNode != null && dateNode.Attributes["年"] != null) {
                try {
                    int wYear = int.Parse(dateNode.Attributes["年"].Value);
                    rocDate = string.Format("中華民國{0}年{1}月{2}日", (wYear > 1911 ? wYear - 1911 : wYear), dateNode.Attributes["月"].Value, dateNode.Attributes["日"].Value);
                } catch { }
            }
            
            XmlNode unitNode = GetNodeByLocalName(xmlDoc, "全銜");
            string unitName = (unitNode != null) ? unitNode.InnerText : "高雄市政府地政局大寮地政事務所";
            
            string rIdYear = GetNodeTextByLocalName(xmlDoc, "年度");
            string rIdWord = GetNodeTextByLocalName(xmlDoc, "字");
            string rIdNum = GetNodeTextByLocalName(xmlDoc, "流水號");
            string fullId = string.Format("{0}{1}第{2}號", rIdYear, rIdWord, rIdNum);
            
            string address = GetNodeTextByLocalName(xmlDoc, "地址");
            if (string.IsNullOrEmpty(address)) address = "83157高雄市大寮區仁勇路69號";
            
            string phone = GetNodeTextByLocalName(xmlDoc, "聯絡方式", "電話");
            string fax = GetNodeTextByLocalName(xmlDoc, "聯絡方式", "傳真");
            string email = GetNodeTextByLocalName(xmlDoc, "聯絡方式", "電子信箱");
            string contactMan = GetNodeTextByLocalName(xmlDoc, "聯絡方式", "承辦人");

            // 2. 準備 RTF 標頭
            string rtfHeader = @"{\rtf1\ansi\ansicpg950\deff0\deflang1033\deflangfe1028" +
                                @"{\fonttbl{\f0\fnil\fprq1\fcharset136 \'bc\'d0\'b7\'a2\'c5'e9;}}" +
                                @"{\colortbl ;\red0\green0\blue255;}" +
                                @"\paperw12240\paperh15840\margl1080\margr1080\margt1080\margb1080\gutter0" +
                                @"\viewkind4\uc1\pard\lang1028\f0\fs40 ";

            System.Text.StringBuilder sb = new System.Text.StringBuilder();
            sb.Append(rtfHeader);
            
            sb.Append(@"\pard\qc\fs44\b " + EncodeRtf(unitName) + @"  " + EncodeRtf(docType) + @"\b0\par ");
            sb.Append(@"\pard\qr\fs20 " + EncodeRtf("地址：") + EncodeRtf(address) + @"\par ");
            if (!string.IsNullOrEmpty(phone)) sb.Append(EncodeRtf("電話：") + EncodeRtf(phone) + @"\par ");
            if (!string.IsNullOrEmpty(fax)) sb.Append(EncodeRtf("傳真：") + EncodeRtf(fax) + @"\par ");
            if (!string.IsNullOrEmpty(email)) sb.Append(EncodeRtf("電子信箱：") + EncodeRtf(email) + @"\par ");
            if (!string.IsNullOrEmpty(contactMan)) sb.Append(EncodeRtf("承辦人：") + EncodeRtf(contactMan) + @"\par ");

            sb.Append(@"\pard\ql\fs24\par " + EncodeRtf("受文者：") + @"\par ");
            sb.Append(EncodeRtf("發文日期：") + EncodeRtf(rocDate) + @"\par ");
            sb.Append(EncodeRtf("發文字號：") + EncodeRtf(fullId) + @"\par ");
            
            XmlNode speedNode = GetNodeByLocalName(xmlDoc, "速別");
            string speed = (speedNode != null && speedNode.Attributes["代碼"] != null) ? speedNode.Attributes["代碼"].Value : "普通件";
            sb.Append(EncodeRtf("速別：") + EncodeRtf(speed) + @"\par ");
            
            XmlNode secretNode = GetNodeByLocalName(xmlDoc, "密等及解密條件或保密期限");
            string secret = (secretNode != null && secretNode.Attributes["代碼"] != null) ? secretNode.Attributes["代碼"].Value : "";
            sb.Append(EncodeRtf("密等及解密條件或保密期限：") + EncodeRtf(secret) + @"\par ");
            sb.Append(EncodeRtf("附件：") + @"\par ");
            
            string subject = GetNodeTextByLocalName(xmlDoc, "主旨");
            if (string.IsNullOrEmpty(subject)) {
                 XmlNode sNode = GetNodeByLocalName(xmlDoc, "主旨");
                 if (sNode != null) subject = sNode.InnerText.Trim();
            }
            sb.Append(@"\par\pard\fi-800\li800\fs32\b " + EncodeRtf("主旨：") + @"\b0 " + EncodeRtf(subject) + @"\par ");
            
            XmlNodeList pNodes = GetNodesByLocalName(xmlDoc, "條列");
            if (pNodes != null && pNodes.Count > 0) {
                sb.Append(@"\pard\fi-800\li800\fs32\b " + EncodeRtf("說明：") + @"\b0\par ");
                foreach (XmlNode p in pNodes) {
                    string prefix = (p.Attributes["序號"] != null) ? p.Attributes["序號"].Value : "";
                    sb.Append(@"\fi-560\li800\fs32 " + EncodeRtf(prefix) + EncodeRtf(p.InnerText.Trim()) + @"\par ");
                }
            }
            
            XmlNodeList originalNodes = GetNodesByLocalName(xmlDoc, "正本");
            if (originalNodes != null) {
                foreach (XmlNode n in originalNodes) {
                    XmlNode nameNode = GetNodeByLocalName(n, "顯示名稱");
                    if (nameNode != null) sb.Append(@"\par\pard\fi-720\li720\fs24\b " + EncodeRtf("正本：") + @"\b0 " + EncodeRtf(nameNode.InnerText) + @"\par ");
                }
            }
            
            XmlNodeList copyNodes = GetNodesByLocalName(xmlDoc, "副本");
            if (copyNodes != null) {
                foreach (XmlNode n in copyNodes) {
                    XmlNode nameNode = GetNodeByLocalName(n, "顯示名稱");
                    if (nameNode != null) sb.Append(@"\pard\fi-720\li720\fs24\b " + EncodeRtf("副本：") + @"\b0 " + EncodeRtf(nameNode.InnerText) + @"\par ");
                }
            }
            
            string responsibly = GetNodeTextByLocalName(xmlDoc, "分層負責");
            if (!string.IsNullOrEmpty(responsibly)) {
                sb.Append(@"\par\pard\fs24 " + EncodeRtf(responsibly) + @"\par ");
            }
            
            string draftMethod = GetNodeTextByLocalName(xmlDoc, "擬辦方式");
            if (!string.IsNullOrEmpty(draftMethod)) {
                sb.Append(@"\pard\fs24 " + EncodeRtf(draftMethod) + @"\par ");
            }
            
            sb.Append(@"}");
            File.WriteAllText(outRtf, sb.ToString(), System.Text.Encoding.ASCII);
        }

        static void SaveSemanticTxt(string outTxt, XmlDocument xmlDoc, string sdiFile)
        {
            string docType = "未知";
            string styleName = "";
            XmlNode typeNode = GetNodeByLocalName(xmlDoc, "文別");
            if (typeNode != null) {
                if (typeNode.Attributes["名稱"] != null) docType = typeNode.Attributes["名稱"].Value;
                if (typeNode.Attributes["樣式檔"] != null) styleName = typeNode.Attributes["樣式檔"].Value;
            }

            if (docType.Equals("便簽")) {
                SaveNoteTxt(outTxt, xmlDoc, sdiFile);
                return;
            }
            if (docType.Equals("開會通知單")) {
                SaveMeetingNoticeTxt(outTxt, xmlDoc, sdiFile);
                return;
            }
            if (docType.Equals("會勘通知單")) {
                SaveSiteInspectionNoticeTxt(outTxt, xmlDoc, sdiFile);
                return;
            }
            
            bool isExplanation = docType.Contains("說明") || 
                                 styleName.Contains("說明") || 
                                 sdiFile.Contains("說明") || 
                                 docType.IndexOf("A4", StringComparison.OrdinalIgnoreCase) >= 0;

            if (isExplanation) {
                SaveExplanationTxt(outTxt, xmlDoc, sdiFile);
                return;
            }

            if (docType.Equals("簽")) {
                SaveMemorandumTxt(outTxt, xmlDoc, sdiFile);
                return;
            }

            System.Text.StringBuilder sb = new System.Text.StringBuilder();
            
            string unitName = GetNodeTextByLocalName(xmlDoc, "全銜");
            if (string.IsNullOrEmpty(unitName)) unitName = "高雄市政府地政局大寮地政事務所";
            
            string rIdYear = GetNodeTextByLocalName(xmlDoc, "年度");
            string rIdWord = GetNodeTextByLocalName(xmlDoc, "字");
            string rIdNum = GetNodeTextByLocalName(xmlDoc, "流水號");
            string fullId = string.Format("{0}{1}第{2}號", rIdYear, rIdWord, rIdNum);
            
            string address = GetNodeTextByLocalName(xmlDoc, "地址");
            if (string.IsNullOrEmpty(address)) address = "83157高雄市大寮區仁勇路69號";
            
            string dept = GetNodeTextByLocalName(xmlDoc, "單位");
            if (string.IsNullOrEmpty(dept)) dept = "測量課";
            string contactMan = GetNodeTextByLocalName(xmlDoc, "聯絡方式", "承辦人");
            string phone = GetNodeTextByLocalName(xmlDoc, "聯絡方式", "電話");
            string fax = GetNodeTextByLocalName(xmlDoc, "聯絡方式", "傳真");
            string email = GetNodeTextByLocalName(xmlDoc, "聯絡方式", "電子信箱");
            
            string contactStr = string.Format("承辦單位({0})、承辦人({1})、電話({2})、傳真({3})、電子信箱({4})", dept, contactMan, phone, fax, email);

            XmlNode dateNode = GetNodeByLocalName(xmlDoc, "發文日期");
            string rocDate = "";
            if (dateNode != null && dateNode.Attributes["年"] != null) {
                try {
                    int wYear = int.Parse(dateNode.Attributes["年"].Value);
                    rocDate = string.Format("中華民國{0}年{1}月{2}日", (wYear > 1911 ? wYear - 1911 : wYear), dateNode.Attributes["月"].Value, dateNode.Attributes["日"].Value);
                } catch { }
            }

            string speed = "普通件";
            XmlNode sNode = GetNodeByLocalName(xmlDoc, "速別");
            if (sNode != null && sNode.Attributes["代碼"] != null) speed = sNode.Attributes["代碼"].Value;
            
            string secret = "";
            XmlNode secNode = GetNodeByLocalName(xmlDoc, "密等及解密條件或保密期限");
            if (secNode != null && secNode.Attributes["代碼"] != null) secret = secNode.Attributes["代碼"].Value;

            sb.AppendLine("【" + unitName + "】");
            sb.AppendLine("【內部編號】(" + Path.GetFileNameWithoutExtension(sdiFile) + ")");
            sb.AppendLine("發文字號：" + fullId);
            sb.AppendLine("地址：" + address);
            sb.AppendLine("聯絡方式：" + contactStr);
            sb.AppendLine("發文日期：" + rocDate);
            sb.AppendLine("速別：" + speed);
            sb.AppendLine("密等及解密條件或保密期限：" + secret);
            
            List<string> originals = new List<string>();
            XmlNodeList oNodes = GetNodesByLocalName(xmlDoc, "正本");
            if (oNodes != null) {
                foreach (XmlNode n in oNodes) {
                    XmlNode nameNode = GetNodeByLocalName(n, "顯示名稱");
                    if (nameNode != null) originals.Add(nameNode.InnerText.Trim());
                }
            }
            sb.AppendLine("正本：" + string.Join("、", originals.ToArray()));
            
            List<string> copies = new List<string>();
            XmlNodeList cNodes = GetNodesByLocalName(xmlDoc, "副本");
            if (cNodes != null) {
                foreach (XmlNode n in cNodes) {
                    XmlNode nameNode = GetNodeByLocalName(n, "顯示名稱");
                    if (nameNode != null) copies.Add(nameNode.InnerText.Trim());
                }
            }
            sb.AppendLine("副本：" + string.Join("、", copies.ToArray()));
            
            string subject = GetNodeTextByLocalName(xmlDoc, "主旨");
            sb.AppendLine("主旨：" + subject);
            sb.AppendLine("說明：");
            
            XmlNodeList pNodes = GetNodesByLocalName(xmlDoc, "條列");
            if (pNodes != null) {
                foreach (XmlNode p in pNodes) {
                    string prefix = (p.Attributes["序號"] != null) ? p.Attributes["序號"].Value : "";
                    sb.AppendLine("\t" + prefix + p.InnerText.Trim());
                }
            }
            
            sb.AppendLine("");
            File.WriteAllText(outTxt, sb.ToString(), System.Text.Encoding.Unicode);
        }

        static XmlNode GetNodeByLocalName(XmlNode parent, string localName)
        {
            return parent.SelectSingleNode(".//*[local-name()='" + localName + "']");
        }

        static XmlNodeList GetNodesByLocalName(XmlNode parent, string localName)
        {
            return parent.SelectNodes(".//*[local-name()='" + localName + "']");
        }

        static string GetNodeTextByLocalName(XmlNode parent, string localName, string attrFilter = null)
        {
            XmlNodeList nodes = parent.SelectNodes(".//*[local-name()='" + localName + "']");
            if (nodes == null) return "";
            foreach (XmlNode node in nodes) {
                if (attrFilter == null) return node.InnerText.Trim();
                foreach (XmlAttribute attr in node.Attributes) {
                    if (attr.Value == attrFilter) return node.InnerText.Trim();
                }
            }
            return "";
        }

        static void SaveStructuredRtf(string outRtf, string sContent, string sdiFile)
        {
            string dir = Path.GetDirectoryName(sdiFile);
            string hdPath = Path.Combine(dir, "hdnote.xml");
            
            string rDate = GetMetadataValue(hdPath, "R_DATE");
            string rocDate = FormatRocDate(rDate);
            string rId = GetMetadataValue(hdPath, "R_ID");

            string rtfHeader = @"{\rtf1\ansi\ansicpg950\deff0\deflang1033\deflangfe1028" +
                                @"{\fonttbl{\f0\fscript\fprq1\fcharset136 \'bc\'d0\'b7\'a2\'c5'e9;}}" +
                                @"{\colortbl ;\red0\green0\blue255;}" +
                                @"\paperw12240\paperh15840\margl1080\margr1080\margt1080\margb1080\gutter0" +
                                @"\viewkind4\uc1\pard\lang1028\f0\fs40 ";

            string[] lines = sContent.Split(new string[] { "\r\n", "\n" }, StringSplitOptions.None);
            List<string> subjectLines = new List<string>();
            List<string> descLines = new List<string>();
            List<string> footerLines = new List<string>();
            List<string> headerLines = new List<string>();
            
            int state = 0;
            foreach (string line in lines) {
                string trimmed = line.Trim();
                if (string.IsNullOrEmpty(trimmed)) continue;
                if (state == 0 && (trimmed.StartsWith("有關") || trimmed.Contains("圖乙案"))) state = 1;
                else if (state == 1 && (trimmed.StartsWith("依據") || trimmed.Contains("辦理") || trimmed.StartsWith("說明"))) state = 2;
                else if (state == 2 && (trimmed.StartsWith("正本") || trimmed.StartsWith("副本") || trimmed.StartsWith("臺灣") || trimmed.Contains("律師"))) state = 3;
                switch (state) {
                    case 0: headerLines.Add(trimmed); break;
                    case 1: subjectLines.Add(trimmed); break;
                    case 2:
                        string descContent = trimmed;
                        if (descContent.StartsWith("說明：")) descContent = descContent.Substring(3).Trim();
                        if (!string.IsNullOrEmpty(descContent)) descLines.Add(descContent);
                        break;
                    case 3: footerLines.Add(trimmed); break;
                }
            }
            if (subjectLines.Count == 0 && descLines.Count == 0) descLines.AddRange(lines.Select(l => l.Trim()).Where(l => !string.IsNullOrEmpty(l)));

            System.Text.StringBuilder sb = new System.Text.StringBuilder();
            sb.Append(rtfHeader);
            sb.Append(@"\pard\qc\fs40 " + EncodeRtf("高雄市政府地政局大寮地政事務所") + @"  " + EncodeRtf("函") + @"\par\n");
            sb.Append(@"\pard\li5600\fs24 " + EncodeRtf("地址：83157高雄市大寮區仁勇路69號") + @"\par\n");
            sb.Append(@"\pard\fs24 " + EncodeRtf("發文日期：") + EncodeRtf(rocDate) + @"\par\n");
            sb.Append(@"\pard\fs24 " + EncodeRtf("發文字號：") + EncodeRtf(rId) + @"\par\n");
            sb.Append(@"\pard\fs24 " + EncodeRtf("速別：普通件") + @"\par\n");
            sb.Append(@"\par\pard\fi-960\li960\fs32 " + EncodeRtf("主旨：") + EncodeRtf(string.Join("", subjectLines.ToArray())) + @"\par\n");
            if (descLines.Count > 0) {
                sb.Append(@"\pard\fi-960\li960\fs32 " + EncodeRtf("說明：") + @"\par\n");
                string[] chineseNums = { "一", "二", "三", "四", "五", "六", "七", "八", "九", "十" };
                for (int i = 0; i < descLines.Count; i++) {
                    string prefix = (i < chineseNums.Length) ? chineseNums[i] + "、" : (i + 1) + "、";
                    sb.Append(@"\fi-640\li960\fs32 " + EncodeRtf(prefix) + EncodeRtf(descLines[i]) + @"\par\n");
                }
            }
            if (footerLines.Count > 0) {
                sb.Append(@"\par\pard\fi-720\li720\fs24 ");
                bool firstFooter = true;
                foreach (string fLine in footerLines) {
                    if (!firstFooter) sb.Append(@"\par\n");
                    string displayLine = fLine;
                    if (!displayLine.Contains("：") && !displayLine.Contains(":")) {
                        if (displayLine.Contains("法院")) displayLine = "正本：" + displayLine;
                        else if (displayLine.Contains("律師") || displayLine.Contains("課")) displayLine = "副本：" + displayLine;
                        else displayLine = "附註/簽署：" + displayLine;
                    }
                    sb.Append(EncodeRtf(displayLine));
                    firstFooter = false;
                }
                sb.Append(@"\par\n");
            }
            sb.Append(@"}");
            File.WriteAllText(outRtf, sb.ToString(), System.Text.Encoding.ASCII);
        }

        static void SaveNoteRtf(string outRtf, XmlDocument xmlDoc, string sdiFile)
        {
            XmlNode dateNode = GetNodeByLocalName(xmlDoc, "發文日期");
            string rocDate = "";
            if (dateNode != null && dateNode.Attributes["年"] != null) {
                try {
                    int wYear = int.Parse(dateNode.Attributes["年"].Value);
                    rocDate = string.Format("{0}年{1}月{2}日", (wYear > 1911 ? wYear - 1911 : wYear), dateNode.Attributes["月"].Value, dateNode.Attributes["日"].Value);
                } catch { }
            }
            string dept = GetNodeTextByLocalName(xmlDoc, "承辦單位");
            if (string.IsNullOrEmpty(dept)) dept = GetNodeTextByLocalName(xmlDoc, "聯絡方式", "承辦單位");

            string rtfHeader = @"{\rtf1\ansi\ansicpg950\deff0\deflang1033\deflangfe1028" +
                                @"{\fonttbl{\f0\fnil\fprq1\fcharset136 \'bc\'d0\'b7\'a2\'c5'e9;}}" +
                                @"{\colortbl ;\red0\green0\blue255;}" +
                                @"\paperw12240\paperh15840\margl1080\margr1080\margt1080\margb1080\gutter0" +
                                @"\viewkind4\uc1\pard\lang1028\f0 ";

            System.Text.StringBuilder sb = new System.Text.StringBuilder();
            sb.Append(rtfHeader);
            sb.Append(@"\trowd\trgaph0\trrh340\clvertalb\cellx1000\clvertalc\cellx6600\clvertalb\cellx9600");
            sb.Append(@"\pard\intbl\qc\fs60 " + EncodeRtf("簽") + @"\cell ");
            sb.Append(@"\pard\intbl\fs40 " + EncodeRtf("於 ") + EncodeRtf(dept) + @"\cell ");
            sb.Append(@"\pard\intbl\qr\fs24 " + EncodeRtf("日期：") + EncodeRtf(rocDate) + @"\cell ");
            sb.Append(@"\row\pard\par ");

            string subject = GetNodeTextByLocalName(xmlDoc, "主旨");
            if (!string.IsNullOrEmpty(subject)) sb.Append(@"\pard\fi0\li0\fs32 " + EncodeRtf(subject) + @"\par ");

            sb.Append(@"\pard\fi0\li0\fs32 " + EncodeRtf("擬：") + @"\par ");
            XmlNodeList pNodes = GetNodesByLocalName(xmlDoc, "條列");
            if (pNodes != null && pNodes.Count > 0) {
                foreach (XmlNode p in pNodes) {
                    string prefix = (p.Attributes["序號"] != null) ? p.Attributes["序號"].Value : "";
                    sb.Append(@"\fi-640\li960\fs32 " + EncodeRtf(prefix) + EncodeRtf(p.InnerText.Trim()) + @"\par ");
                }
            } else {
                string desc = GetNodeTextByLocalName(xmlDoc, "說明");
                if (!string.IsNullOrEmpty(desc)) sb.Append(@"\fi-640\li960\fs32 " + EncodeRtf(desc) + @"\par ");
            }
            sb.Append(@"}");
            File.WriteAllText(outRtf, sb.ToString(), System.Text.Encoding.ASCII);
            Console.WriteLine("       -> [便簽] RTF 轉換完成。");
        }

        static void SaveNoteTxt(string outTxt, XmlDocument xmlDoc, string sdiFile)
        {
            System.Text.StringBuilder sb = new System.Text.StringBuilder();
            string dept = GetNodeTextByLocalName(xmlDoc, "承辦單位");
            if (string.IsNullOrEmpty(dept)) dept = GetNodeTextByLocalName(xmlDoc, "聯絡方式", "承辦單位");
            XmlNode dateNode = GetNodeByLocalName(xmlDoc, "發文日期");
            string rocDate = "";
            if (dateNode != null && dateNode.Attributes["年"] != null) {
                try {
                    int wYear = int.Parse(dateNode.Attributes["年"].Value);
                    rocDate = string.Format("{0}年{1}月{2}日", (wYear > 1911 ? wYear - 1911 : wYear), dateNode.Attributes["月"].Value, dateNode.Attributes["日"].Value);
                } catch { }
            }
            sb.AppendLine("【便簽】");
            sb.AppendLine("於：" + dept);
            sb.AppendLine("日期：" + rocDate);
            string subject = GetNodeTextByLocalName(xmlDoc, "主旨");
            if (!string.IsNullOrEmpty(subject)) sb.AppendLine(subject);
            sb.AppendLine("擬：");
            XmlNodeList pNodes = GetNodesByLocalName(xmlDoc, "條列");
            if (pNodes != null && pNodes.Count > 0) {
                foreach (XmlNode p in pNodes) {
                    string prefix = (p.Attributes["序號"] != null) ? p.Attributes["序號"].Value : "";
                    sb.AppendLine("\t" + prefix + p.InnerText.Trim());
                }
            } else {
                string desc = GetNodeTextByLocalName(xmlDoc, "說明");
                if (!string.IsNullOrEmpty(desc)) sb.AppendLine("\t" + desc);
            }
            File.WriteAllText(outTxt, sb.ToString(), System.Text.Encoding.Unicode);
            Console.WriteLine("       -> [便簽] TXT 轉換完成。");
        }

        static void SaveMeetingNoticeRtf(string outRtf, XmlDocument xmlDoc, string sdiFile)
        {
            XmlNode unitNode = GetNodeByLocalName(xmlDoc, "全銜");
            string unitName = (unitNode != null) ? unitNode.InnerText : "高雄市政府地政局大寮地政事務所";
            XmlNode typeNode = GetNodeByLocalName(xmlDoc, "文別");
            string docType = (typeNode != null && typeNode.Attributes["名稱"] != null) ? typeNode.Attributes["名稱"].Value : "開會通知單";

            XmlNode dateNode = GetNodeByLocalName(xmlDoc, "發文日期");
            string rocDate = "";
            if (dateNode != null && dateNode.Attributes["年"] != null) {
                try {
                    int wYear = int.Parse(dateNode.Attributes["年"].Value);
                    rocDate = string.Format("中華民國{0}年{1}月{2}日", (wYear > 1911 ? wYear - 1911 : wYear), dateNode.Attributes["月"].Value, dateNode.Attributes["日"].Value);
                } catch { }
            }
            string rIdYear = GetNodeTextByLocalName(xmlDoc, "年度");
            string rIdWord = GetNodeTextByLocalName(xmlDoc, "字");
            string rIdNum = GetNodeTextByLocalName(xmlDoc, "流水號");
            string fullId = string.Format("{0}{1}第{2}號", rIdYear, rIdWord, rIdNum);
            string speed = GetNodeTextByLocalName(xmlDoc, "速別");
            if (string.IsNullOrEmpty(speed)) {
                XmlNode sNode = GetNodeByLocalName(xmlDoc, "速別");
                speed = (sNode != null && sNode.Attributes["代碼"] != null) ? sNode.Attributes["代碼"].Value : "普通件";
            }
            string secret = GetNodeTextByLocalName(xmlDoc, "密等及解密條件或保密期限");
            if (string.IsNullOrEmpty(secret)) {
                XmlNode secNode = GetNodeByLocalName(xmlDoc, "密等及解密條件或保密期限");
                secret = (secNode != null && secNode.Attributes["代碼"] != null) ? secNode.Attributes["代碼"].Value : "";
            }
            string appendix = GetNodeTextByLocalName(xmlDoc, "附件");
            string reason = GetNodeTextByLocalName(xmlDoc, "開會事由");
            if (string.IsNullOrEmpty(reason)) reason = GetNodeTextByLocalName(xmlDoc, "主旨");
            string location = GetNodeTextByLocalName(xmlDoc, "開會地點");
            string chair = GetNodeTextByLocalName(xmlDoc, "主持人");
            string contactMan = GetNodeTextByLocalName(xmlDoc, "聯絡人");
            if (string.IsNullOrEmpty(contactMan)) contactMan = GetNodeTextByLocalName(xmlDoc, "聯絡方式", "承辦人");
            string phone = GetNodeTextByLocalName(xmlDoc, "電話");
            if (string.IsNullOrEmpty(phone)) phone = GetNodeTextByLocalName(xmlDoc, "聯絡方式", "電話");
            string contactStr = string.Format("{0} {1}", contactMan, phone).Trim();

            string rtfHeader = @"{\rtf1\ansi\ansicpg950\deff0\deflang1033\deflangfe1028" +
                                @"{\fonttbl{\f0\fnil\fprq1\fcharset136 \'bc\'d0\'b7\'a2\'c5'e9;}}" +
                                @"{\colortbl ;\red0\green0\blue255;}" +
                                @"\paperw12240\paperh15840\margl1080\margr1080\margt1080\margb1080\gutter0" +
                                @"\viewkind4\uc1\pard\lang1028\f0 ";

            System.Text.StringBuilder sb = new System.Text.StringBuilder();
            sb.Append(rtfHeader);
            sb.Append(@"\pard\qc\fs40\b " + EncodeRtf(unitName) + EncodeRtf("　") + EncodeRtf(docType) + @"\b0\par ");
            sb.Append(@"\par\pard\fs24 " + EncodeRtf("發文日期：") + EncodeRtf(rocDate) + @"\par ");
            sb.Append(EncodeRtf("發文字號：") + EncodeRtf(fullId) + @"\par ");
            sb.Append(EncodeRtf("速別：") + EncodeRtf(speed) + @"\par ");
            sb.Append(EncodeRtf("密等及解密條件或保密期限：") + EncodeRtf(secret) + @"\par ");
            sb.Append(EncodeRtf("附件：") + EncodeRtf(appendix) + @"\par ");
            sb.Append(@"\par\pard\fs32 " + EncodeRtf("開會事由：") + EncodeRtf(reason) + @"\par ");
            sb.Append(EncodeRtf("開會地點：") + EncodeRtf(location) + @"\par ");
            sb.Append(@"\fi-1280\li1280 " + EncodeRtf("主持人：") + EncodeRtf(chair) + @"\par ");
            sb.Append(@"\fi-2240\li2240 " + EncodeRtf("聯絡人及電話：") + EncodeRtf(contactStr) + @"\par ");

            List<string> attendees = new List<string>();
            XmlNodeList attendeeNodes = GetNodesByLocalName(xmlDoc, "出席者");
            if (attendeeNodes != null && attendeeNodes.Count > 0) foreach (XmlNode n in attendeeNodes) attendees.Add(n.InnerText.Trim());
            sb.Append(@"\par\pard\fi-960\li960\fs24 " + EncodeRtf("出席者：") + EncodeRtf(string.Join("、", attendees.ToArray())) + @"\par ");

            List<string> copies = new List<string>();
            XmlNodeList copyNodes = GetNodesByLocalName(xmlDoc, "副本");
            if (copyNodes != null) foreach (XmlNode n in copyNodes) {
                XmlNode nameNode = GetNodeByLocalName(n, "顯示名稱");
                if (nameNode != null) copies.Add(nameNode.InnerText.Trim());
                else copies.Add(n.InnerText.Trim());
            }
            sb.Append(@"\pard\fi-720\li720\fs24 " + EncodeRtf("副本：") + EncodeRtf(string.Join("、", copies.ToArray())) + @"\par ");

            XmlNodeList pNodes = GetNodesByLocalName(xmlDoc, "條列");
            if (pNodes != null && pNodes.Count > 0) {
                sb.Append(@"\par\pard\fi0\li0\fs32 " + EncodeRtf("備註：") + @"\par ");
                foreach (XmlNode p in pNodes) {
                    string prefix = (p.Attributes["序號"] != null) ? p.Attributes["序號"].Value : "";
                    sb.Append(@"\fi-640\li960\fs32 " + EncodeRtf(prefix) + EncodeRtf(p.InnerText.Trim()) + @"\par ");
                }
            } else {
                string remark = GetNodeTextByLocalName(xmlDoc, "備註");
                if (!string.IsNullOrEmpty(remark)) sb.Append(@"\par\pard\fi0\li0\fs32 " + EncodeRtf("備註：") + EncodeRtf(remark) + @"\par ");
            }

            sb.Append(@"}");
            File.WriteAllText(outRtf, sb.ToString(), System.Text.Encoding.ASCII);
            Console.WriteLine("       -> [開會通知單] RTF 轉換完成。");
        }

        static void SaveMeetingNoticeTxt(string outTxt, XmlDocument xmlDoc, string sdiFile)
        {
            System.Text.StringBuilder sb = new System.Text.StringBuilder();
            string unitName = GetNodeTextByLocalName(xmlDoc, "全銜");
            if (string.IsNullOrEmpty(unitName)) unitName = "高雄市政府地政局大寮地政事務所";
            XmlNode dateNode = GetNodeByLocalName(xmlDoc, "發文日期");
            string rocDate = "";
            if (dateNode != null && dateNode.Attributes["年"] != null) {
                try {
                    int wYear = int.Parse(dateNode.Attributes["年"].Value);
                    rocDate = string.Format("中華民國{0}年{1}月{2}日", (wYear > 1911 ? wYear - 1911 : wYear), dateNode.Attributes["月"].Value, dateNode.Attributes["日"].Value);
                } catch { }
            }
            string rIdYear = GetNodeTextByLocalName(xmlDoc, "年度");
            string rIdWord = GetNodeTextByLocalName(xmlDoc, "字");
            string rIdNum = GetNodeTextByLocalName(xmlDoc, "流水號");
            string fullId = string.Format("{0}{1}第{2}號", rIdYear, rIdWord, rIdNum);

            sb.AppendLine("【開會通知單】");
            sb.AppendLine("機關：" + unitName);
            sb.AppendLine("發文日期：" + rocDate);
            sb.AppendLine("發文字號：" + fullId);
            sb.AppendLine("開會事由：" + (GetNodeTextByLocalName(xmlDoc, "開會事由") ?? GetNodeTextByLocalName(xmlDoc, "主旨")));
            sb.AppendLine("開會地點：" + GetNodeTextByLocalName(xmlDoc, "開會地點"));
            sb.AppendLine("主持人：" + GetNodeTextByLocalName(xmlDoc, "主持人"));
            string contactMan = GetNodeTextByLocalName(xmlDoc, "聯絡人") ?? GetNodeTextByLocalName(xmlDoc, "聯絡方式", "承辦人");
            string phone = GetNodeTextByLocalName(xmlDoc, "電話") ?? GetNodeTextByLocalName(xmlDoc, "聯絡方式", "電話");
            sb.AppendLine("聯絡人及電話：" + contactMan + " " + phone);
            sb.AppendLine("出席者：" + string.Join("、", GetNodesByLocalName(xmlDoc, "出席者").Cast<XmlNode>().Select(n => n.InnerText.Trim()).ToArray()));
            sb.AppendLine("備註：");

            XmlNodeList pNodes = GetNodesByLocalName(xmlDoc, "條列");
            if (pNodes != null && pNodes.Count > 0) {
                foreach (XmlNode p in pNodes) {
                    string prefix = (p.Attributes["序號"] != null) ? p.Attributes["序號"].Value : "";
                    sb.AppendLine("\t" + prefix + p.InnerText.Trim());
                }
            } else {
                string remark = GetNodeTextByLocalName(xmlDoc, "備註");
                if (!string.IsNullOrEmpty(remark)) sb.AppendLine("\t" + remark);
            }

            File.WriteAllText(outTxt, sb.ToString(), System.Text.Encoding.Unicode);
            Console.WriteLine("       -> [開會通知單] TXT 轉換完成。");
        }

        static void SaveSiteInspectionNoticeRtf(string outRtf, XmlDocument xmlDoc, string sdiFile)
        {
            XmlNode unitNode = GetNodeByLocalName(xmlDoc, "全銜");
            string unitName = (unitNode != null) ? unitNode.InnerText : "高雄市政府地政局大寮地政事務所";
            XmlNode typeNode = GetNodeByLocalName(xmlDoc, "文別");
            string docType = (typeNode != null && typeNode.Attributes["名稱"] != null) ? typeNode.Attributes["名稱"].Value : "會勘通知單";
            XmlNode dateNode = GetNodeByLocalName(xmlDoc, "發文日期");
            string rocDate = "";
            if (dateNode != null && dateNode.Attributes["年"] != null) {
                try {
                    int wYear = int.Parse(dateNode.Attributes["年"].Value);
                    rocDate = string.Format("中華民國{0}年{1}月{2}日", (wYear > 1911 ? wYear - 1911 : wYear), dateNode.Attributes["月"].Value, dateNode.Attributes["日"].Value);
                } catch { }
            }
            string rIdYear = GetNodeTextByLocalName(xmlDoc, "年度");
            string rIdWord = GetNodeTextByLocalName(xmlDoc, "字");
            string rIdNum = GetNodeTextByLocalName(xmlDoc, "流水號");
            string fullId = string.Format("{0}{1}第{2}號", rIdYear, rIdWord, rIdNum);
            string address = GetNodeTextByLocalName(xmlDoc, "地址");
            if (string.IsNullOrEmpty(address)) address = "83157高雄市大寮區仁勇路69號";
            string speed = GetNodeTextByLocalName(xmlDoc, "速別");
            if (string.IsNullOrEmpty(speed)) {
                XmlNode sNode = GetNodeByLocalName(xmlDoc, "速別");
                speed = (sNode != null && sNode.Attributes["代碼"] != null) ? sNode.Attributes["代碼"].Value : "普通件";
            }
            string secret = GetNodeTextByLocalName(xmlDoc, "密等及解密條件或保密期限");
            if (string.IsNullOrEmpty(secret)) {
                XmlNode secNode = GetNodeByLocalName(xmlDoc, "密等及解密條件或保密期限");
                secret = (secNode != null && secNode.Attributes["代碼"] != null) ? secNode.Attributes["代碼"].Value : "";
            }
            string appendix = GetNodeTextByLocalName(xmlDoc, "附件");
            string reason = GetNodeTextByLocalName(xmlDoc, "會勘事由");
            if (string.IsNullOrEmpty(reason)) reason = GetNodeTextByLocalName(xmlDoc, "主旨");
            string time = GetNodeTextByLocalName(xmlDoc, "會勘時間");
            string location = GetNodeTextByLocalName(xmlDoc, "會勘地點");
            string chair = GetNodeTextByLocalName(xmlDoc, "主持人");
            string contactMan = GetNodeTextByLocalName(xmlDoc, "聯絡人");
            if (string.IsNullOrEmpty(contactMan)) contactMan = GetNodeTextByLocalName(xmlDoc, "聯絡方式", "承辦人");
            string phone = GetNodeTextByLocalName(xmlDoc, "電話");
            if (string.IsNullOrEmpty(phone)) phone = GetNodeTextByLocalName(xmlDoc, "聯絡方式", "電話");
            string contactStr = string.Format("{0} {1}", contactMan, phone).Trim();

            string rtfHeader = @"{\rtf1\ansi\ansicpg950\deff0\deflang1033\deflangfe1028" +
                                @"{\fonttbl{\f0\fnil\fprq1\fcharset136 \'bc\'d0\'b7\'a2\'c5'e9;}}" +
                                @"{\colortbl ;\red0\green0\blue255;}" +
                                @"\paperw12240\paperh15840\margl1080\margr1080\margt1080\margb1080\gutter0" +
                                @"\viewkind4\uc1\pard\lang1028\f0 ";

            System.Text.StringBuilder sb = new System.Text.StringBuilder();
            sb.Append(rtfHeader);
            sb.Append(@"\pard\qc\fs40\b " + EncodeRtf(unitName) + EncodeRtf("　") + EncodeRtf(docType) + @"\b0\par ");
            sb.Append(@"\par\pard\li5600\fs24 " + EncodeRtf("地址：") + EncodeRtf(address) + @"\par ");
            sb.Append(@"\pard\fs24 " + EncodeRtf("發文日期：") + EncodeRtf(rocDate) + @"\par ");
            sb.Append(EncodeRtf("發文字號：") + EncodeRtf(fullId) + @"\par ");
            sb.Append(EncodeRtf("速別：") + EncodeRtf(speed) + @"\par ");
            sb.Append(EncodeRtf("密等及解密條件或保密期限：") + EncodeRtf(secret) + @"\par ");
            sb.Append(EncodeRtf("附件：") + EncodeRtf(appendix) + @"\par ");
            sb.Append(@"\par\pard\fs32 " + EncodeRtf("會勘事由：") + EncodeRtf(reason) + @"\par ");
            sb.Append(EncodeRtf("會勘時間：") + EncodeRtf(time) + @"\par ");
            sb.Append(EncodeRtf("會勘地點：") + EncodeRtf(location) + @"\par ");
            sb.Append(@"\fi-1280\li1280 " + EncodeRtf("主持人：") + EncodeRtf(chair) + @"\par ");
            sb.Append(@"\fi-2240\li2240 " + EncodeRtf("聯絡人及電話：") + EncodeRtf(contactStr) + @"\par ");

            List<string> attendees = new List<string>();
            XmlNodeList attendeeNodes = GetNodesByLocalName(xmlDoc, "出席者");
            if (attendeeNodes != null && attendeeNodes.Count > 0) foreach (XmlNode n in attendeeNodes) attendees.Add(n.InnerText.Trim());
            sb.Append(@"\par\pard\fi-960\li960\fs24 " + EncodeRtf("出席者：") + EncodeRtf(string.Join("、", attendees.ToArray())) + @"\par ");

            List<string> copies = new List<string>();
            XmlNodeList copyNodes = GetNodesByLocalName(xmlDoc, "副本");
            if (copyNodes != null) foreach (XmlNode n in copyNodes) {
                XmlNode nameNode = GetNodeByLocalName(n, "顯示名稱");
                if (nameNode != null) copies.Add(nameNode.InnerText.Trim());
                else copies.Add(n.InnerText.Trim());
            }
            sb.Append(@"\pard\fi-720\li720\fs24 " + EncodeRtf("副本：") + EncodeRtf(string.Join("、", copies.ToArray())) + @"\par ");

            XmlNodeList pNodes = GetNodesByLocalName(xmlDoc, "條列");
            if (pNodes != null && pNodes.Count > 0) {
                sb.Append(@"\par\pard\fi-960\li960\fs32 " + EncodeRtf("備註：") + @"\par ");
                foreach (XmlNode p in pNodes) {
                    string prefix = (p.Attributes["序號"] != null) ? p.Attributes["序號"].Value : "";
                    sb.Append(@"\fi-640\li960\fs32 " + EncodeRtf(prefix) + EncodeRtf(p.InnerText.Trim()) + @"\par ");
                }
            } else {
                string remark = GetNodeTextByLocalName(xmlDoc, "備註");
                if (!string.IsNullOrEmpty(remark)) sb.Append(@"\par\pard\fi-960\li960\fs32 " + EncodeRtf("備註：") + EncodeRtf(remark) + @"\par ");
            }

            sb.Append(@"}");
            File.WriteAllText(outRtf, sb.ToString(), System.Text.Encoding.ASCII);
            Console.WriteLine("       -> [會勘通知單] RTF 轉換完成。");
        }

        static void SaveSiteInspectionNoticeTxt(string outTxt, XmlDocument xmlDoc, string sdiFile)
        {
            System.Text.StringBuilder sb = new System.Text.StringBuilder();
            string unitName = GetNodeTextByLocalName(xmlDoc, "全銜");
            if (string.IsNullOrEmpty(unitName)) unitName = "高雄市政府地政局大寮地政事務所";
            XmlNode dateNode = GetNodeByLocalName(xmlDoc, "發文日期");
            string rocDate = "";
            if (dateNode != null && dateNode.Attributes["年"] != null) {
                try {
                    int wYear = int.Parse(dateNode.Attributes["年"].Value);
                    rocDate = string.Format("中華民國{0}年{1}月{2}日", (wYear > 1911 ? wYear - 1911 : wYear), dateNode.Attributes["月"].Value, dateNode.Attributes["日"].Value);
                } catch { }
            }
            string rIdYear = GetNodeTextByLocalName(xmlDoc, "年度");
            string rIdWord = GetNodeTextByLocalName(xmlDoc, "字");
            string rIdNum = GetNodeTextByLocalName(xmlDoc, "流水號");
            string fullId = string.Format("{0}{1}第{2}號", rIdYear, rIdWord, rIdNum);

            sb.AppendLine("【會勘通知單】");
            sb.AppendLine("機關：" + unitName);
            sb.AppendLine("發文日期：" + rocDate);
            sb.AppendLine("發文字號：" + fullId);
            sb.AppendLine("會勘事由：" + (GetNodeTextByLocalName(xmlDoc, "會勘事由") ?? GetNodeTextByLocalName(xmlDoc, "主旨")));
            sb.AppendLine("會勘時間：" + GetNodeTextByLocalName(xmlDoc, "會勘時間"));
            sb.AppendLine("會勘地點：" + GetNodeTextByLocalName(xmlDoc, "會勘地點"));
            sb.AppendLine("主持人：" + GetNodeTextByLocalName(xmlDoc, "主持人"));
            string contactMan = GetNodeTextByLocalName(xmlDoc, "聯絡人") ?? GetNodeTextByLocalName(xmlDoc, "聯絡方式", "承辦人");
            string phone = GetNodeTextByLocalName(xmlDoc, "電話") ?? GetNodeTextByLocalName(xmlDoc, "聯絡方式", "電話");
            sb.AppendLine("聯絡人及電話：" + contactMan + " " + phone);
            sb.AppendLine("出席者：" + string.Join("、", GetNodesByLocalName(xmlDoc, "出席者").Cast<XmlNode>().Select(n => n.InnerText.Trim()).ToArray()));
            sb.AppendLine("備註：");

            XmlNodeList pNodes = GetNodesByLocalName(xmlDoc, "條列");
            if (pNodes != null && pNodes.Count > 0) {
                foreach (XmlNode p in pNodes) {
                    string prefix = (p.Attributes["序號"] != null) ? p.Attributes["序號"].Value : "";
                    sb.AppendLine("\t" + prefix + p.InnerText.Trim());
                }
            } else {
                string remark = GetNodeTextByLocalName(xmlDoc, "備註");
                if (!string.IsNullOrEmpty(remark)) sb.AppendLine("\t" + remark);
            }

            File.WriteAllText(outTxt, sb.ToString(), System.Text.Encoding.Unicode);
            Console.WriteLine("       -> [會勘通知單] TXT 轉換完成。");
        }

        static void SaveExplanationRtf(string outRtf, XmlDocument xmlDoc, string sdiFile)
        {
            string rtfHeader = @"{\rtf1\ansi\ansicpg950\deff0\deflang1033\deflangfe1028" +
                                @"{\fonttbl{\f0\fnil\fprq1\fcharset136 \'bc\'d0\'b7\'a2\'c5'e9;}}" +
                                @"{\colortbl ;\red0\green0\blue255;}" +
                                @"\paperw12240\paperh15840\margl1080\margr1080\margt1080\margb1080\gutter0" +
                                @"\viewkind4\uc1\pard\lang1028\f0 ";

            System.Text.StringBuilder sb = new System.Text.StringBuilder();
            sb.Append(rtfHeader);

            // 附件處理 (比照範例必印標頭，並優先嘗試由 XML 提取內容)
            sb.Append(@"\pard\fs24 " + EncodeRtf("附件："));
            XmlNode attachmentNode = GetNodeByLocalName(xmlDoc, "附件");
            if (attachmentNode != null) {
                string appendixDesc = GetNodeTextByLocalName(attachmentNode, "說明");
                if (!string.IsNullOrEmpty(appendixDesc)) sb.Append(EncodeRtf(appendixDesc));
                
                XmlNodeList fileNodes = GetNodesByLocalName(attachmentNode, "檔案");
                if (fileNodes != null && fileNodes.Count > 0) {
                    List<string> fNames = new List<string>();
                    foreach (XmlNode f in fileNodes) {
                        string name = "";
                        if (f.Attributes["原始檔名"] != null) name = f.Attributes["原始檔名"].Value;
                        else if (f.Attributes["名稱"] != null) name = f.Attributes["名稱"].Value;
                        else if (f.Attributes["檔名"] != null) name = f.Attributes["檔名"].Value;
                        else name = f.InnerText.Trim();
                        
                        if (!string.IsNullOrEmpty(name)) fNames.Add(name);
                    }
                    if (fNames.Count > 0) {
                        if (!string.IsNullOrEmpty(appendixDesc)) sb.Append(EncodeRtf("、"));
                        sb.Append(EncodeRtf(string.Join("、", fNames.ToArray())));
                    }
                } else if (string.IsNullOrEmpty(appendixDesc)) {
                    string totalText = attachmentNode.InnerText.Trim();
                    if (!string.IsNullOrEmpty(totalText)) sb.Append(EncodeRtf(totalText));
                }
            }
            sb.Append(@"\par ");

            // 內容段落處理 (說明、擬辦)
            XmlNodeList paragraphs = GetNodesByLocalName(xmlDoc, "段落");
            if (paragraphs != null && paragraphs.Count > 0) {
                foreach (XmlNode p in paragraphs) {
                    string pName = p.Attributes["段名"] != null ? p.Attributes["段名"].Value : "";
                    if (!string.IsNullOrEmpty(pName)) {
                        sb.Append(@"\par\pard\fi-960\li960\fs32 " + EncodeRtf(pName) + @"\u12288?\par ");
                    }
                    
                    XmlNodeList items = GetNodesByLocalName(p, "條列");
                    foreach (XmlNode item in items) {
                        string prefix = (item.Attributes["序號"] != null) ? item.Attributes["序號"].Value : "";
                        sb.Append(@"\fi-640\li960\fs32 " + EncodeRtf(prefix) + EncodeRtf(item.InnerText.Trim()) + @"\par ");
                    }
                }
            } else {
                // 退回模式：若無段落標記，則嘗試全局抓取條列節點
                XmlNodeList pNodes = GetNodesByLocalName(xmlDoc, "條列");
                if (pNodes != null && pNodes.Count > 0) {
                    sb.Append(@"\par\pard\fi-960\li960\fs32 " + EncodeRtf("說明：") + @"\u12288?\par ");
                    foreach (XmlNode p in pNodes) {
                        string prefix = (p.Attributes["序號"] != null) ? p.Attributes["序號"].Value : "";
                        sb.Append(@"\fi-640\li960\fs32 " + EncodeRtf(prefix) + EncodeRtf(p.InnerText.Trim()) + @"\par ");
                    }
                }
            }

            sb.Append(@"\par\par}");
            File.WriteAllText(outRtf, sb.ToString(), System.Text.Encoding.ASCII);
            Console.WriteLine("       -> [說明] RTF 轉換完成。");
        }

        static void SaveExplanationTxt(string outTxt, XmlDocument xmlDoc, string sdiFile)
        {
            System.Text.StringBuilder sb = new System.Text.StringBuilder();
            sb.AppendLine("【說明】");

            // 附件
            string appendixLine = "附件：";
            XmlNode attachmentNode = GetNodeByLocalName(xmlDoc, "附件");
            if (attachmentNode != null) {
                string appendixDesc = GetNodeTextByLocalName(attachmentNode, "說明");
                if (!string.IsNullOrEmpty(appendixDesc)) appendixLine += appendixDesc;
                
                XmlNodeList fileNodes = GetNodesByLocalName(attachmentNode, "檔案");
                if (fileNodes != null && fileNodes.Count > 0) {
                    List<string> fNames = new List<string>();
                    foreach (XmlNode f in fileNodes) {
                        string name = "";
                        if (f.Attributes["原始檔名"] != null) name = f.Attributes["原始檔名"].Value;
                        else if (f.Attributes["名稱"] != null) name = f.Attributes["名稱"].Value;
                        else if (f.Attributes["檔名"] != null) name = f.Attributes["檔名"].Value;
                        else name = f.InnerText.Trim();
                        
                        if (!string.IsNullOrEmpty(name)) fNames.Add(name);
                    }
                    if (fNames.Count > 0) {
                        if (!appendixLine.Equals("附件：")) appendixLine += "、";
                        appendixLine += string.Join("、", fNames.ToArray());
                    }
                } else if (appendixLine.Equals("附件：")) {
                    string totalText = attachmentNode.InnerText.Trim();
                    if (!string.IsNullOrEmpty(totalText)) appendixLine += totalText;
                }
            }
            sb.AppendLine(appendixLine);

            // 內容段落
            XmlNodeList paragraphs = GetNodesByLocalName(xmlDoc, "段落");
            if (paragraphs != null && paragraphs.Count > 0) {
                foreach (XmlNode p in paragraphs) {
                    string pName = p.Attributes["段名"] != null ? p.Attributes["段名"].Value : "";
                    if (!string.IsNullOrEmpty(pName)) sb.AppendLine(pName);
                    
                    XmlNodeList items = GetNodesByLocalName(p, "條列");
                    foreach (XmlNode item in items) {
                        string prefix = (item.Attributes["序號"] != null) ? item.Attributes["序號"].Value : "";
                        sb.AppendLine("\t" + prefix + item.InnerText.Trim());
                    }
                }
            } else {
                sb.AppendLine("說明：");
                XmlNodeList pNodes = GetNodesByLocalName(xmlDoc, "條列");
                if (pNodes != null) {
                    foreach (XmlNode p in pNodes) {
                        string prefix = (p.Attributes["序號"] != null) ? p.Attributes["序號"].Value : "";
                        sb.AppendLine("\t" + prefix + p.InnerText.Trim());
                    }
                }
            }

            File.WriteAllText(outTxt, sb.ToString(), System.Text.Encoding.Unicode);
            Console.WriteLine("       -> [說明] TXT 轉換完成。");
        }

        static void SaveMemorandumRtf(string outRtf, XmlDocument xmlDoc, string sdiFile)
        {
            XmlNode dateNode = GetNodeByLocalName(xmlDoc, "年月日");
            string rocDate = "";
            if (dateNode != null && dateNode.Attributes["年"] != null) {
                try {
                    int wYear = int.Parse(dateNode.Attributes["年"].Value);
                    rocDate = string.Format("中華民國{0}年{1}月{2}日", (wYear > 1911 ? wYear - 1911 : wYear), dateNode.Attributes["月"].Value, dateNode.Attributes["日"].Value);
                } catch { }
            } else {
                dateNode = GetNodeByLocalName(xmlDoc, "發文日期");
                if (dateNode != null && dateNode.Attributes["年"] != null) {
                    try {
                        int wYear = int.Parse(dateNode.Attributes["年"].Value);
                        rocDate = string.Format("中華民國{0}年{1}月{2}日", (wYear > 1911 ? wYear - 1911 : wYear), dateNode.Attributes["月"].Value, dateNode.Attributes["日"].Value);
                    } catch { }
                }
            }

            XmlNode unitNode = GetNodeByLocalName(xmlDoc, "全銜");
            string unitName = (unitNode != null) ? unitNode.InnerText : "高雄市政府地政局大寮地政事務所";
            
            string rIdYear = GetNodeTextByLocalName(xmlDoc, "年度");
            string rIdWord = GetNodeTextByLocalName(xmlDoc, "字");
            string rIdNum = GetNodeTextByLocalName(xmlDoc, "流水號");
            if (string.IsNullOrEmpty(rIdNum)) rIdNum = GetNodeTextByLocalName(xmlDoc, "內部流水號");
            string fullId = (string.IsNullOrEmpty(rIdYear) && string.IsNullOrEmpty(rIdWord) && string.IsNullOrEmpty(rIdNum)) ? "" : string.Format("{0}{1}第{2}號", rIdYear, rIdWord, rIdNum);
            
            string address = GetNodeTextByLocalName(xmlDoc, "地址");
            if (string.IsNullOrEmpty(address)) address = "83157高雄市大寮區仁勇路69號";

            string rtfHeader = @"{\rtf1\ansi\ansicpg950\deff0\deflang1033\deflangfe1028" +
                                @"{\fonttbl{\f0\fnil\fprq1\fcharset136 \'bc\'d0\'b7\'a2\'c5'e9;}}" +
                                @"{\colortbl ;\red0\green0\blue255;}" +
                                @"\paperw12240\paperh15840\margl1080\margr1080\margt1080\margb1080\gutter0" +
                                @"\viewkind4\uc1\pard\lang1028\f0 ";

            System.Text.StringBuilder sb = new System.Text.StringBuilder();
            sb.Append(rtfHeader);
            
            // Header: Unit + "簽"
            sb.Append(@"\pard\qc\fs44\b " + EncodeRtf(unitName) + @"  " + EncodeRtf("簽") + @"\b0\par ");
            
            // Address: Right aligned
            sb.Append(@"\pard\qr\fs20 " + EncodeRtf("地址：") + EncodeRtf(address) + @"\par ");
            
            // Metadata Block: Left aligned
            sb.Append(@"\pard\ql\fs24\par " + EncodeRtf("受文者：") + @"\par ");
            sb.Append(EncodeRtf("發文日期：") + EncodeRtf(rocDate) + @"\par ");
            sb.Append(EncodeRtf("發文字號：") + EncodeRtf(fullId) + @"\par ");
            
            XmlNode speedNode = GetNodeByLocalName(xmlDoc, "速別");
            string speed = (speedNode != null && speedNode.Attributes["代碼"] != null) ? speedNode.Attributes["代碼"].Value : "普通件";
            sb.Append(EncodeRtf("速別：") + EncodeRtf(speed) + @"\par ");
            
            XmlNode secretNode = GetNodeByLocalName(xmlDoc, "密等及解密條件或保密期限");
            string secret = (secretNode != null && secretNode.Attributes["代碼"] != null) ? secretNode.Attributes["代碼"].Value : "";
            sb.Append(EncodeRtf("密等及解密條件或保密期限：") + EncodeRtf(secret) + @"\par ");

            // Attachments
            sb.Append(EncodeRtf("附件："));
            XmlNode attachmentNode = GetNodeByLocalName(xmlDoc, "附件");
            if (attachmentNode != null) {
                string appendixDesc = GetNodeTextByLocalName(attachmentNode, "說明");
                if (!string.IsNullOrEmpty(appendixDesc)) sb.Append(EncodeRtf(appendixDesc));
                
                XmlNodeList fileNodes = GetNodesByLocalName(attachmentNode, "檔案");
                if (fileNodes != null && fileNodes.Count > 0) {
                    List<string> fNames = new List<string>();
                    foreach (XmlNode f in fileNodes) {
                        string name = "";
                        if (f.Attributes["原始檔名"] != null) name = f.Attributes["原始檔名"].Value;
                        else if (f.Attributes["名稱"] != null) name = f.Attributes["名稱"].Value;
                        else if (f.Attributes["檔名"] != null) name = f.Attributes["檔名"].Value;
                        else name = f.InnerText.Trim();
                        if (!string.IsNullOrEmpty(name)) fNames.Add(name);
                    }
                    if (fNames.Count > 0) {
                        if (!string.IsNullOrEmpty(appendixDesc)) sb.Append(EncodeRtf("、"));
                        sb.Append(EncodeRtf(string.Join("、", fNames.ToArray())));
                    }
                }
            }
            sb.Append(@"\par\par ");

            // Body: Subject
            string subject = GetNodeTextByLocalName(xmlDoc, "主旨");
            sb.Append(@"\pard\fi-800\li800\fs32\b " + EncodeRtf("主旨：") + @"\b0 " + EncodeRtf(subject) + @"\par ");
            
            // Body: Paragraphs (说明、拟办)
            XmlNodeList paragraphs = GetNodesByLocalName(xmlDoc, "段落");
            if (paragraphs != null && paragraphs.Count > 0) {
                foreach (XmlNode p in paragraphs) {
                    string pName = p.Attributes["段名"] != null ? p.Attributes["段名"].Value : "";
                    if (!string.IsNullOrEmpty(pName)) {
                        sb.Append(@"\par\pard\fi-800\li800\fs32\b " + EncodeRtf(pName) + @"\b0\par ");
                    }
                    
                    // Paragraph lead-in text
                    XmlNode textNode = GetNodeByLocalName(p, "文字");
                    if (textNode != null && !string.IsNullOrEmpty(textNode.InnerText)) {
                         sb.Append(@"\fi-560\li800\fs32 " + EncodeRtf(textNode.InnerText.Trim()) + @"\par ");
                    }

                    XmlNodeList items = GetNodesByLocalName(p, "條列");
                    foreach (XmlNode item in items) {
                        string prefix = (item.Attributes["序號"] != null) ? item.Attributes["序號"].Value : "";
                        sb.Append(@"\fi-560\li800\fs32 " + EncodeRtf(prefix) + EncodeRtf(item.InnerText.Trim()) + @"\par ");
                    }
                }
            } else {
                // Fallback: If no paragraphs, try to get all items
                XmlNodeList items = GetNodesByLocalName(xmlDoc, "條列");
                if (items != null && items.Count > 0) {
                    sb.Append(@"\par\pard\fi-800\li800\fs32\b " + EncodeRtf("說明：") + @"\b0\par ");
                    foreach (XmlNode item in items) {
                        string prefix = (item.Attributes["序號"] != null) ? item.Attributes["序號"].Value : "";
                        sb.Append(@"\fi-560\li800\fs32 " + EncodeRtf(prefix) + EncodeRtf(item.InnerText.Trim()) + @"\par ");
                    }
                }
            }

            sb.Append(@"\par\par}");
            File.WriteAllText(outRtf, sb.ToString(), System.Text.Encoding.ASCII);
            Console.WriteLine("       -> [簽] RTF 轉換完成。");
        }

        static void SaveMemorandumTxt(string outTxt, XmlDocument xmlDoc, string sdiFile)
        {
            System.Text.StringBuilder sb = new System.Text.StringBuilder();
            sb.AppendLine("【簽】");
            
            XmlNode unitNode = GetNodeByLocalName(xmlDoc, "全銜");
            string unitName = (unitNode != null) ? unitNode.InnerText : "高雄市政府地政局大寮地政事務所";
            sb.AppendLine("機關：" + unitName);

            XmlNode dateNode = GetNodeByLocalName(xmlDoc, "年月日");
            string rocDate = "";
            if (dateNode != null && dateNode.Attributes["年"] != null) {
                try {
                    int wYear = int.Parse(dateNode.Attributes["年"].Value);
                    rocDate = string.Format("中華民國{0}年{1}月{2}日", (wYear > 1911 ? wYear - 1911 : wYear), dateNode.Attributes["月"].Value, dateNode.Attributes["日"].Value);
                } catch { }
            }
            sb.AppendLine("日期：" + rocDate);

            string subject = GetNodeTextByLocalName(xmlDoc, "主旨");
            sb.AppendLine("主旨：" + subject);

            XmlNodeList paragraphs = GetNodesByLocalName(xmlDoc, "段落");
            if (paragraphs != null && paragraphs.Count > 0) {
                foreach (XmlNode p in paragraphs) {
                    string pName = p.Attributes["段名"] != null ? p.Attributes["段名"].Value : "";
                    if (!string.IsNullOrEmpty(pName)) sb.AppendLine(pName);
                    
                    XmlNode textNode = GetNodeByLocalName(p, "文字");
                    if (textNode != null && !string.IsNullOrEmpty(textNode.InnerText)) {
                        sb.AppendLine("\t" + textNode.InnerText.Trim());
                    }

                    XmlNodeList items = GetNodesByLocalName(p, "條列");
                    foreach (XmlNode item in items) {
                        string prefix = (item.Attributes["序號"] != null) ? item.Attributes["序號"].Value : "";
                        sb.AppendLine("\t" + prefix + item.InnerText.Trim());
                    }
                }
            } else {
                sb.AppendLine("說明：");
                XmlNodeList items = GetNodesByLocalName(xmlDoc, "條列");
                if (items != null) {
                    foreach (XmlNode p in items) {
                        string prefix = (p.Attributes["序號"] != null) ? p.Attributes["序號"].Value : "";
                        sb.AppendLine("\t" + prefix + p.InnerText.Trim());
                    }
                }
            }

            File.WriteAllText(outTxt, sb.ToString(), System.Text.Encoding.Unicode);
            Console.WriteLine("       -> [簽] TXT 轉換完成。");
        }

        static void SaveDi(string outDi, XmlDocument xmlDoc, string sdiFile)
        {
            string docType = "未知";
            XmlNode typeNode = GetNodeByLocalName(xmlDoc, "文別");
            if (typeNode != null && typeNode.Attributes["名稱"] != null) docType = typeNode.Attributes["名稱"].Value;

            if (docType.Contains("簽") && !docType.Contains("稿"))
            {
                // 方案 A: 略過簽的 di 轉換
                return;
            }

            string rootTag = "函";
            string dtd = "104_2_utf8.dtd";

            if (docType.Contains("開會通知單"))
            {
                rootTag = "開會通知單";
                dtd = "104_4_utf8.dtd";
            }
            else if (docType.Contains("會勘通知單"))
            {
                rootTag = "會勘通知單";
                dtd = "104_4_utf8.dtd";
            }

            string baseName = Path.GetFileNameWithoutExtension(outDi);
            System.Text.StringBuilder sb = new System.Text.StringBuilder();
            sb.AppendLine("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
            sb.AppendLine(string.Format("<!DOCTYPE {0} SYSTEM \"{1}\" [", rootTag, dtd));
            sb.AppendLine(string.Format("<!ENTITY 表單 SYSTEM \"{0}1_A00.sw\" NDATA DI>", baseName));
            
            // 附件處理 (如果有附件檔名)
            List<string> attachmentNames = new List<string>();
            XmlNodeList attNodes = GetNodesByLocalName(xmlDoc, "檔案");
            if (attNodes != null)
            {
                int attIdx = 1;
                foreach (XmlNode att in attNodes)
                {
                    string attName = att.Attributes["原始檔名"] != null ? att.Attributes["原始檔名"].Value : 
                                   (att.Attributes["名稱"] != null ? att.Attributes["名稱"].Value : "");
                    if (!string.IsNullOrEmpty(attName))
                    {
                        string entName = "ATTCH" + attIdx;
                        sb.AppendLine(string.Format("<!ENTITY {0} SYSTEM \"{1}\" NDATA _X>", entName, attName));
                        attachmentNames.Add(entName);
                        attIdx++;
                    }
                }
            }

            sb.AppendLine("<!NOTATION DI SYSTEM \"\">");
            sb.AppendLine("<!NOTATION _X SYSTEM \"\">");
            sb.AppendLine("]>");

            sb.AppendLine("<" + rootTag + ">");

            // 1. 發文機關
            XmlNode unitNode = GetNodeByLocalName(xmlDoc, "全銜");
            string unitName = (unitNode != null) ? unitNode.InnerText : "高雄市政府地政局大寮地政事務所";
            string unitCode = (unitNode != null && unitNode.Attributes["機關代碼"] != null) ? unitNode.Attributes["機關代碼"].Value : "397163100A";
            sb.AppendLine("\t<發文機關>");
            sb.AppendLine("\t\t<全銜>" + unitName + "</全銜>");
            sb.AppendLine("\t\t<機關代碼>" + unitCode + "</機關代碼>");
            sb.AppendLine("\t</發文機關>");

            if (rootTag == "函")
            {
                sb.AppendLine("\t<函類別 代碼=\"" + rootTag + "\"/>");
            }

            // 2. 基本資訊
            string address = GetNodeTextByLocalName(xmlDoc, "地址");
            if (string.IsNullOrEmpty(address)) address = "83157高雄市大寮區仁勇路69號";
            sb.AppendLine("\t<地址>" + address + "</地址>");

            if (rootTag == "函")
            {
                string contactMan = GetNodeTextByLocalName(xmlDoc, "聯絡方式", "承辦人");
                string phone = GetNodeTextByLocalName(xmlDoc, "聯絡方式", "電話");
                string fax = GetNodeTextByLocalName(xmlDoc, "聯絡方式", "傳真");
                string email = GetNodeTextByLocalName(xmlDoc, "聯絡方式", "電子信箱");
                string dept = GetNodeTextByLocalName(xmlDoc, "聯絡方式", "承辦單位");
                if (string.IsNullOrEmpty(dept)) dept = "測量課";

                if (!string.IsNullOrEmpty(dept)) sb.AppendLine("\t<聯絡方式>承辦單位：" + dept + "</聯絡方式>");
                if (!string.IsNullOrEmpty(contactMan)) sb.AppendLine("\t<聯絡方式>承辦人：" + contactMan + "</聯絡方式>");
                if (!string.IsNullOrEmpty(phone)) sb.AppendLine("\t<聯絡方式>電話：" + phone + "</聯絡方式>");
                if (!string.IsNullOrEmpty(fax)) sb.AppendLine("\t<聯絡方式>傳真：" + fax + "</聯絡方式>");
                if (!string.IsNullOrEmpty(email)) sb.AppendLine("\t<聯絡方式>電子信箱：" + email + "</聯絡方式>");
            }
            else // 通知單格式
            {
                string contactMan = GetNodeTextByLocalName(xmlDoc, "聯絡方式", "承辦人");
                string phone = GetNodeTextByLocalName(xmlDoc, "聯絡方式", "電話");
                sb.AppendLine("\t<聯絡人及電話>");
                sb.AppendLine("\t\t<姓名>" + contactMan + "</姓名>");
                sb.AppendLine("\t\t<電話>" + phone + "</電話>");
                sb.AppendLine("\t</聯絡人及電話>");
            }

            sb.AppendLine("\t<受文者>");
            sb.AppendLine("\t\t<交換表 交換表單=\"表單\">如正副本行文單位</交換表>");
            sb.AppendLine("\t</受文者>");

            sb.AppendLine("\t<發文日期>");
            sb.AppendLine("\t\t<年月日></年月日>"); // 通常由系統填入
            sb.AppendLine("\t</發文日期>");

            XmlNode speedNode = GetNodeByLocalName(xmlDoc, "速別");
            string speed = (speedNode != null && speedNode.Attributes["代碼"] != null) ? speedNode.Attributes["代碼"].Value : "普通件";
            sb.AppendLine("\t<速別 代碼=\"" + speed + "\"/>");

            XmlNode secretNode = GetNodeByLocalName(xmlDoc, "密等及解密條件或保密期限");
            string secret = (secretNode != null && secretNode.Attributes["代碼"] != null) ? secretNode.Attributes["代碼"].Value : "";
            sb.AppendLine("\t<密等及解密條件或保密期限>");
            sb.AppendLine("\t\t<密等>" + secret + "</密等>");
            sb.AppendLine("\t\t<解密條件或保密期限></解密條件或保密期限>");
            sb.AppendLine("\t</密等及解密條件或保密期限>");

            // 3. 附件
            sb.AppendLine("\t<附件>");
            string attText = "";
            XmlNodeList attList = GetNodesByLocalName(xmlDoc, "檔案");
            if (attList != null)
            {
                List<string> names = new List<string>();
                foreach (XmlNode n in attList)
                {
                    string nStr = n.Attributes["原始檔名"] != null ? n.Attributes["原始檔名"].Value : 
                                 (n.Attributes["名稱"] != null ? n.Attributes["名稱"].Value : "");
                    if (!string.IsNullOrEmpty(nStr)) names.Add(nStr);
                }
                attText = string.Join("、", names.ToArray());
            }
            sb.AppendLine("\t\t<文字>" + attText + "</文字>");
            foreach (var ent in attachmentNames)
            {
                sb.AppendLine("\t\t<附件檔名 附件名=\"" + ent + "\"/>");
            }
            sb.AppendLine("\t</附件>");

            // 4. 主旨 / 開會事由
            string subject = GetNodeTextByLocalName(xmlDoc, "主旨");
            if (rootTag == "函")
            {
                sb.AppendLine("\t<主旨>");
                sb.AppendLine("\t\t<文字>" + subject + "</文字>");
                sb.AppendLine("\t</主旨>");
            }
            else
            {
                sb.AppendLine("\t<開會事由>");
                sb.AppendLine("\t\t<文字>" + subject + "</文字>");
                sb.AppendLine("\t</開會事由>");
                
                string location = GetNodeTextByLocalName(xmlDoc, "開會地點");
                sb.AppendLine("\t<開會地點>");
                sb.AppendLine("\t\t<文字>" + location + "</文字>");
                sb.AppendLine("\t</開會地點>");
                
                string host = GetNodeTextByLocalName(xmlDoc, "主持人");
                sb.AppendLine("\t<主持人>");
                sb.AppendLine("\t\t<姓名>" + host + "</姓名>");
                sb.AppendLine("\t</主持人>");
                
                sb.AppendLine("\t<出席者>");
                XmlNodeList attendees = GetNodesByLocalName(xmlDoc, "出席者"); // 或正本
                if (attendees == null || attendees.Count == 0) attendees = GetNodesByLocalName(xmlDoc, "正本");
                foreach (XmlNode att in attendees)
                {
                    XmlNode nameNode = GetNodeByLocalName(att, "顯示名稱");
                    if (nameNode != null) sb.AppendLine("\t\t<全銜>" + nameNode.InnerText + "</全銜>");
                }
                sb.AppendLine("\t</出席者>");
            }

            // 5. 段落 (說明 / 備註)
            string pTag = (rootTag == "函") ? "段落" : "備註";
            if (rootTag != "函") sb.AppendLine("\t<備註>");
            
            XmlNodeList paragraphs = GetNodesByLocalName(xmlDoc, "段落");
            if (paragraphs != null && paragraphs.Count > 0)
            {
                foreach (XmlNode p in paragraphs)
                {
                    string pName = p.Attributes["段名"] != null ? p.Attributes["段名"].Value : "說明：";
                    sb.AppendLine("\t\t<段落 段名=\"" + pName + "\">");
                    sb.AppendLine("\t\t\t<文字></文字>");
                    
                    XmlNodeList items = GetNodesByLocalName(p, "條列");
                    foreach (XmlNode item in items)
                    {
                        string prefix = (item.Attributes["序號"] != null) ? item.Attributes["序號"].Value : "";
                        sb.AppendLine("\t\t\t<條列 序號=\"" + prefix + "\">");
                        sb.AppendLine("\t\t\t\t<文字>" + item.InnerText.Trim() + "</文字>");
                        sb.AppendLine("\t\t\t</條列>");
                    }
                    sb.AppendLine("\t\t</段落>");
                }
            }
            else
            {
                // Fallback: 找全域條列
                sb.AppendLine("\t\t<段落 段名=\"說明：\">");
                sb.AppendLine("\t\t\t<文字></文字>");
                XmlNodeList items = GetNodesByLocalName(xmlDoc, "條列");
                if (items != null)
                {
                    foreach (XmlNode item in items)
                    {
                        string prefix = (item.Attributes["序號"] != null) ? item.Attributes["序號"].Value : "";
                        sb.AppendLine("\t\t\t<條列 序號=\"" + prefix + "\">");
                        sb.AppendLine("\t\t\t\t<文字>" + item.InnerText.Trim() + "</文字>");
                        sb.AppendLine("\t\t\t</條列>");
                    }
                }
                sb.AppendLine("\t\t</段落>");
            }
            if (rootTag != "函") sb.AppendLine("\t</備註>");

            // 6. 正副本
            XmlNodeList originalNodes = GetNodesByLocalName(xmlDoc, "正本");
            if (originalNodes != null)
            {
                foreach (XmlNode n in originalNodes)
                {
                    XmlNode nameNode = GetNodeByLocalName(n, "顯示名稱");
                    if (nameNode != null)
                    {
                        sb.AppendLine("\t<正本>");
                        sb.AppendLine("\t\t<全銜>" + nameNode.InnerText + "</全銜>");
                        sb.AppendLine("\t</正本>");
                    }
                }
            }

            XmlNodeList copyNodes = GetNodesByLocalName(xmlDoc, "副本");
            if (copyNodes != null)
            {
                foreach (XmlNode n in copyNodes)
                {
                    XmlNode nameNode = GetNodeByLocalName(n, "顯示名稱");
                    if (nameNode != null)
                    {
                        sb.AppendLine("\t<副本>");
                        sb.AppendLine("\t\t<全銜>" + nameNode.InnerText + "</全銜>");
                        sb.AppendLine("\t</副本>");
                    }
                }
            }

            sb.AppendLine("</" + rootTag + ">");

            File.WriteAllText(outDi, sb.ToString(), System.Text.Encoding.UTF8);
            Console.WriteLine("       -> .di 轉換完成 -> " + Path.GetFileName(outDi));
        }

    }

    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public class DummyWindow
    {
        public bool isWindow = true;
    }
}
