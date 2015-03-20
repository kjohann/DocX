using System;
using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Linq;

namespace Novacode
{
    /// <summary>
    /// Represents a table of contents in the document
    /// </summary>
    public class TableOfContents : DocXElement
    {
        #region TocBaseValues

        private const string HeaderStyle = "TOCHeading";
        private const int RightTabPos = 9350;
        private const string TocXmlBase = @"
            <w:sdt xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>
              <w:sdtPr>
                <w:docPartObj>
                  <w:docPartGallery w:val='Table of Contents'/>
                  <w:docPartUnique/>
                </w:docPartObj>\
              </w:sdtPr>
              <w:sdtEndPr>
                <w:rPr>
                  <w:rFonts w:asciiTheme='minorHAnsi' w:cstheme='minorBidi' w:eastAsiaTheme='minorHAnsi' w:hAnsiTheme='minorHAnsi'/>
                  <w:color w:val='auto'/>
                  <w:sz w:val='22'/>
                  <w:szCs w:val='22'/>
                  <w:lang w:eastAsia='en-US'/>
                </w:rPr>
              </w:sdtEndPr>
              <w:sdtContent>
                <w:p>
                  <w:pPr>
                    <w:pStyle w:val='{0}'/>
                  </w:pPr>
                  <w:r>
                    <w:t>{1}</w:t>
                  </w:r>
                </w:p>
                <w:p>
                  <w:pPr>
                    <w:pStyle w:val='TOC1'/>
                    <w:tabs>
                      <w:tab w:val='right' w:leader='dot' w:pos='{2}'/>
                    </w:tabs>
                    <w:rPr>
                      <w:noProof/>
                    </w:rPr>
                  </w:pPr>
                  <w:r>
                    <w:fldChar w:fldCharType='begin' w:dirty='true'/>
                  </w:r>
                  <w:r>
                    <w:instrText xml:space='preserve'> {3} </w:instrText>
                  </w:r>
                  <w:r>
                    <w:fldChar w:fldCharType='separate'/>
                  </w:r>
                </w:p>
                <w:p>
                  <w:r>
                    <w:rPr>
                      <w:b/>
                      <w:bCs/>
                      <w:noProof/>
                    </w:rPr>
                    <w:fldChar w:fldCharType='end'/>
                  </w:r>
                </w:p>
              </w:sdtContent>
            </w:sdt>
        ";

        #endregion

        private TableOfContents(DocX document, XElement xml) : base(document, xml)
        {
            UpdateSettings(document);
        }        

        internal static TableOfContents CreateTableOfContents(DocX document, string title, TableOfContentsSwitches switches, string headerStyle = null, int lastIncludeLevel = 3, int? rightTabPos = null)
        {
            var reader = XmlReader.Create(new StringReader(string.Format(TocXmlBase, headerStyle ?? HeaderStyle, title, rightTabPos ?? RightTabPos, BuildSwitchString(switches, lastIncludeLevel))));
            var xml = XElement.Load(reader);
            return new TableOfContents(document, xml);
        }

        private void UpdateSettings(DocX document)
        {
            if (document.settings.Descendants().Any(x => x.Name.Equals(DocX.w + "updateFields"))) return;
            
            var element = new XElement(XName.Get("updateFields", DocX.w.NamespaceName), new XAttribute(DocX.w + "val", true));
            document.settings.Root.Add(element);
        }

        private static string BuildSwitchString(TableOfContentsSwitches switches, int lastIncludeLevel)
        {
            var allSwitches = Enum.GetValues(typeof (TableOfContentsSwitches)).Cast<TableOfContentsSwitches>();
            var switchString = "TOC";
            foreach (var s in allSwitches.Where(s => s != TableOfContentsSwitches.None && switches.HasFlag(s)))
            {
                switchString += " " + s.EnumDescription();
                if (s == TableOfContentsSwitches.O)
                {
                    switchString += string.Format(" '{0}-{1}'", 1, lastIncludeLevel);
                }
            }

            return switchString;
        }

    }
}