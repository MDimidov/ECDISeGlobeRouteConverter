using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace ECDIS_eGloebe___RouteConverter.ExportRoute
{
	public static class ExportRouteData
	{
        public static StringBuilder DocStartXml = new StringBuilder(@"
<?xml version=""1.0""?>
<?mso-application progid=""Excel.Sheet""?>
<Workbook xmlns=""urn:schemas-microsoft-com:office:spreadsheet""
 xmlns:o=""urn:schemas-microsoft-com:office:office""
 xmlns:x=""urn:schemas-microsoft-com:office:excel""
 xmlns:ss=""urn:schemas-microsoft-com:office:spreadsheet""
 xmlns:html=""http://www.w3.org/TR/REC-html40"">
 <DocumentProperties xmlns=""urn:schemas-microsoft-com:office:office"">
  <Title>FORMS ISO SMS ISPS</Title>
  <Subject>FORM</Subject>
  <Author>Mariqn Dimidov</Author>
  <LastAuthor>Mariqn Dimidov</LastAuthor>
  <Revision>7</Revision>
  <Created>2024-05-04T16:41:26Z</Created>
  <Company>GRIMALDI GROUP</Company>
  <Version>16.00</Version>
 </DocumentProperties>
 <OfficeDocumentSettings xmlns=""urn:schemas-microsoft-com:office:office"">
  <AllowPNG/>
 </OfficeDocumentSettings>
 <ExcelWorkbook xmlns=""urn:schemas-microsoft-com:office:excel"">
  <WindowHeight>11055</WindowHeight>
  <WindowWidth>28800</WindowWidth>
  <WindowTopX>32767</WindowTopX>
  <WindowTopY>32767</WindowTopY>
  <ProtectStructure>False</ProtectStructure>
  <ProtectWindows>False</ProtectWindows>
 </ExcelWorkbook>
 <Styles>
  <Style ss:ID=""Default"" ss:Name=""Normal"">
   <Alignment ss:Vertical=""Bottom""/>
   <Borders/>
   <Font ss:FontName=""Times"" x:Family=""Roman""/>
   <Interior/>
   <NumberFormat/>
   <Protection/>
  </Style>
  <Style ss:ID=""m2353907624288"">
   <Alignment ss:Horizontal=""Left"" ss:Vertical=""Center"" ss:WrapText=""1""/>
   <Borders>
    <Border ss:Position=""Bottom"" ss:LineStyle=""Double"" ss:Weight=""3""/>
    <Border ss:Position=""Left"" ss:LineStyle=""Double"" ss:Weight=""3""/>
    <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Top"" ss:LineStyle=""Double"" ss:Weight=""3""/>
   </Borders>
   <Font ss:FontName=""Arial Narrow"" x:Family=""Swiss"" ss:Size=""8"" ss:Bold=""1""/>
  </Style>
  <Style ss:ID=""m2353907624308"">
   <Alignment ss:Horizontal=""Left"" ss:Vertical=""Center"" ss:WrapText=""1""/>
   <Borders>
    <Border ss:Position=""Bottom"" ss:LineStyle=""Double"" ss:Weight=""3""/>
    <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Right"" ss:LineStyle=""Double"" ss:Weight=""3""/>
    <Border ss:Position=""Top"" ss:LineStyle=""Double"" ss:Weight=""3""/>
   </Borders>
   <Font ss:FontName=""Arial Narrow"" x:Family=""Swiss"" ss:Size=""8"" ss:Bold=""1""/>
  </Style>
  <Style ss:ID=""m2353907304704"">
   <Alignment ss:Horizontal=""Center"" ss:Vertical=""Center"" ss:Rotate=""90""
    ss:WrapText=""1""/>
   <Borders>
    <Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Top"" ss:LineStyle=""Double"" ss:Weight=""3""/>
   </Borders>
   <Font ss:FontName=""Arial Narrow"" x:Family=""Swiss"" ss:Size=""8"" ss:Color=""#000000""
    ss:Italic=""1""/>
  </Style>
  <Style ss:ID=""m2353907304724"">
   <Alignment ss:Horizontal=""Center"" ss:Vertical=""Center"" ss:Rotate=""90""
    ss:WrapText=""1""/>
   <Borders>
    <Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Top"" ss:LineStyle=""Double"" ss:Weight=""3""/>
   </Borders>
   <Font ss:FontName=""Arial Narrow"" x:Family=""Swiss"" ss:Size=""9"" ss:Italic=""1""/>
  </Style>
  <Style ss:ID=""m2353907304744"">
   <Alignment ss:Horizontal=""Center"" ss:Vertical=""Center"" ss:Rotate=""90""
    ss:WrapText=""1""/>
   <Borders>
    <Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Right"" ss:LineStyle=""Double"" ss:Weight=""3""/>
    <Border ss:Position=""Top"" ss:LineStyle=""Double"" ss:Weight=""3""/>
   </Borders>
   <Font ss:FontName=""Arial Narrow"" x:Family=""Swiss"" ss:Size=""8"" ss:Italic=""1""/>
  </Style>
  <Style ss:ID=""m2353907304764"">
   <Alignment ss:Horizontal=""Center"" ss:Vertical=""Center"" ss:WrapText=""1""/>
   <Borders>
    <Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
   </Borders>
   <Font ss:FontName=""Arial Narrow"" x:Family=""Swiss"" ss:Size=""8"" ss:Italic=""1""/>
  </Style>
  <Style ss:ID=""m2353907304784"">
   <Alignment ss:Horizontal=""Center"" ss:Vertical=""Center"" ss:Rotate=""90""
    ss:WrapText=""1""/>
   <Borders>
    <Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
   </Borders>
   <Font ss:FontName=""Arial Narrow"" x:Family=""Swiss"" ss:Size=""8"" ss:Italic=""1""/>
  </Style>
  <Style ss:ID=""m2353907304804"">
   <Alignment ss:Horizontal=""Center"" ss:Vertical=""Center"" ss:Rotate=""90""
    ss:WrapText=""1""/>
   <Borders>
    <Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
   </Borders>
   <Font ss:FontName=""Arial Narrow"" x:Family=""Swiss"" ss:Size=""8"" ss:Italic=""1""/>
  </Style>
  <Style ss:ID=""m2353907307616"">
   <Alignment ss:Horizontal=""Center"" ss:Vertical=""Center"" ss:Rotate=""90""
    ss:WrapText=""1""/>
   <Borders>
    <Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Top"" ss:LineStyle=""Double"" ss:Weight=""3""/>
   </Borders>
   <Font ss:FontName=""Arial Narrow"" x:Family=""Swiss"" ss:Size=""8"" ss:Italic=""1""/>
  </Style>
  <Style ss:ID=""m2353907307636"">
   <Alignment ss:Horizontal=""Center"" ss:Vertical=""Center"" ss:Rotate=""90""
    ss:WrapText=""1""/>
   <Borders>
    <Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Top"" ss:LineStyle=""Double"" ss:Weight=""3""/>
   </Borders>
   <Font ss:FontName=""Arial Narrow"" x:Family=""Swiss"" ss:Size=""8"" ss:Italic=""1""/>
  </Style>
  <Style ss:ID=""m2353907307656"">
   <Alignment ss:Horizontal=""Center"" ss:Vertical=""Center"" ss:Rotate=""90""
    ss:WrapText=""1""/>
   <Borders>
    <Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Top"" ss:LineStyle=""Double"" ss:Weight=""3""/>
   </Borders>
   <Font ss:FontName=""Arial Narrow"" x:Family=""Swiss"" ss:Size=""8"" ss:Italic=""1""/>
  </Style>
  <Style ss:ID=""m2353907307676"">
   <Alignment ss:Horizontal=""Center"" ss:Vertical=""Center"" ss:Rotate=""90""
    ss:WrapText=""1""/>
   <Borders>
    <Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Top"" ss:LineStyle=""Double"" ss:Weight=""3""/>
   </Borders>
   <Font ss:FontName=""Arial Narrow"" x:Family=""Swiss"" ss:Size=""8"" ss:Color=""#000000""
    ss:Italic=""1""/>
  </Style>
  <Style ss:ID=""m2353907307696"">
   <Alignment ss:Horizontal=""Center"" ss:Vertical=""Top"" ss:Rotate=""90""
    ss:WrapText=""1""/>
   <Borders>
    <Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Top"" ss:LineStyle=""Double"" ss:Weight=""3""/>
   </Borders>
   <Font ss:FontName=""Arial Narrow"" x:Family=""Swiss"" ss:Size=""8"" ss:Color=""#000000""
    ss:Italic=""1""/>
  </Style>
  <Style ss:ID=""m2353907307716"">
   <Alignment ss:Horizontal=""Center"" ss:Vertical=""Center"" ss:WrapText=""1""/>
   <Borders>
    <Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Top"" ss:LineStyle=""Double"" ss:Weight=""3""/>
   </Borders>
   <Font ss:FontName=""Arial Narrow"" x:Family=""Swiss"" ss:Size=""8"" ss:Italic=""1""/>
  </Style>
  <Style ss:ID=""m2353907307408"">
   <Alignment ss:Horizontal=""Center"" ss:Vertical=""Center"" ss:Rotate=""90""
    ss:WrapText=""1""/>
   <Borders>
    <Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Top"" ss:LineStyle=""Double"" ss:Weight=""3""/>
   </Borders>
   <Font ss:FontName=""Arial Narrow"" x:Family=""Swiss"" ss:Size=""8"" ss:Italic=""1""/>
  </Style>
  <Style ss:ID=""m2353907307428"">
   <Alignment ss:Horizontal=""Center"" ss:Vertical=""Center"" ss:Rotate=""90""
    ss:WrapText=""1""/>
   <Borders>
    <Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Top"" ss:LineStyle=""Double"" ss:Weight=""3""/>
   </Borders>
   <Font ss:FontName=""Arial Narrow"" x:Family=""Swiss"" ss:Size=""8"" ss:Italic=""1""/>
  </Style>
  <Style ss:ID=""m2353907307448"">
   <Alignment ss:Horizontal=""Center"" ss:Vertical=""Center"" ss:Rotate=""90""
    ss:WrapText=""1""/>
   <Borders>
    <Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Top"" ss:LineStyle=""Double"" ss:Weight=""3""/>
   </Borders>
   <Font ss:FontName=""Arial Narrow"" x:Family=""Swiss"" ss:Size=""8"" ss:Italic=""1""/>
  </Style>
  <Style ss:ID=""m2353907307468"">
   <Alignment ss:Horizontal=""Center"" ss:Vertical=""Center"" ss:Rotate=""90""
    ss:WrapText=""1""/>
   <Borders>
    <Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Top"" ss:LineStyle=""Double"" ss:Weight=""3""/>
   </Borders>
   <Font ss:FontName=""Arial Narrow"" x:Family=""Swiss"" ss:Size=""8"" ss:Italic=""1""/>
  </Style>
  <Style ss:ID=""m2353907307488"">
   <Alignment ss:Horizontal=""Center"" ss:Vertical=""Center"" ss:Rotate=""90""
    ss:WrapText=""1""/>
   <Borders>
    <Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Top"" ss:LineStyle=""Double"" ss:Weight=""3""/>
   </Borders>
   <Font ss:FontName=""Arial Narrow"" x:Family=""Swiss"" ss:Size=""8"" ss:Italic=""1""/>
  </Style>
  <Style ss:ID=""m2353907307508"">
   <Alignment ss:Horizontal=""Center"" ss:Vertical=""Center"" ss:Rotate=""90""
    ss:WrapText=""1""/>
   <Borders>
    <Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Top"" ss:LineStyle=""Double"" ss:Weight=""3""/>
   </Borders>
   <Font ss:FontName=""Arial Narrow"" x:Family=""Swiss"" ss:Size=""8"" ss:Italic=""1""/>
  </Style>
  <Style ss:ID=""m2353907573856"">
   <Alignment ss:Horizontal=""Center"" ss:Vertical=""Center"" ss:Rotate=""90""
    ss:WrapText=""1""/>
   <Borders>
    <Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Left"" ss:LineStyle=""Double"" ss:Weight=""3""/>
    <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Top"" ss:LineStyle=""Double"" ss:Weight=""3""/>
   </Borders>
   <Font ss:FontName=""Arial Narrow"" x:Family=""Swiss"" ss:Size=""8"" ss:Italic=""1""/>
  </Style>
  <Style ss:ID=""m2353907573876"">
   <Alignment ss:Horizontal=""Center"" ss:Vertical=""Center"" ss:Rotate=""90""
    ss:WrapText=""1""/>
   <Borders>
    <Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Top"" ss:LineStyle=""Double"" ss:Weight=""3""/>
   </Borders>
   <Font ss:FontName=""Arial Narrow"" x:Family=""Swiss"" ss:Size=""8"" ss:Italic=""1""/>
  </Style>
  <Style ss:ID=""m2353907573896"">
   <Alignment ss:Horizontal=""Center"" ss:Vertical=""Center"" ss:Rotate=""90""
    ss:WrapText=""1""/>
   <Borders>
    <Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Top"" ss:LineStyle=""Double"" ss:Weight=""3""/>
   </Borders>
   <Font ss:FontName=""Arial Narrow"" x:Family=""Swiss"" ss:Size=""8"" ss:Italic=""1""/>
  </Style>
  <Style ss:ID=""m2353907573916"">
   <Alignment ss:Horizontal=""Center"" ss:Vertical=""Center"" ss:Rotate=""90""
    ss:WrapText=""1""/>
   <Borders>
    <Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Top"" ss:LineStyle=""Double"" ss:Weight=""3""/>
   </Borders>
   <Font ss:FontName=""Arial Narrow"" x:Family=""Swiss"" ss:Size=""9""/>
  </Style>
  <Style ss:ID=""m2353907573936"">
   <Alignment ss:Horizontal=""Center"" ss:Vertical=""Top"" ss:Rotate=""90""
    ss:WrapText=""1""/>
   <Borders>
    <Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Top"" ss:LineStyle=""Double"" ss:Weight=""3""/>
   </Borders>
   <Font ss:FontName=""Arial Narrow"" x:Family=""Swiss"" ss:Size=""8"" ss:Italic=""1""/>
  </Style>
  <Style ss:ID=""m2353907573956"">
   <Alignment ss:Horizontal=""Center"" ss:Vertical=""Center"" ss:Rotate=""90""
    ss:WrapText=""1""/>
   <Borders>
    <Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Top"" ss:LineStyle=""Double"" ss:Weight=""3""/>
   </Borders>
   <Font ss:FontName=""Arial Narrow"" x:Family=""Swiss"" ss:Size=""8"" ss:Italic=""1""/>
  </Style>
  <Style ss:ID=""m2353907575728"">
   <Alignment ss:Horizontal=""Left"" ss:Vertical=""Center"" ss:WrapText=""1""/>
   <Borders>
    <Border ss:Position=""Bottom"" ss:LineStyle=""Double"" ss:Weight=""3""/>
    <Border ss:Position=""Left"" ss:LineStyle=""Double"" ss:Weight=""3""/>
    <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
   </Borders>
   <Font ss:FontName=""Arial Narrow"" x:Family=""Swiss"" ss:Size=""7"" ss:Italic=""1""/>
  </Style>
  <Style ss:ID=""m2353907575748"">
   <Alignment ss:Horizontal=""Left"" ss:Vertical=""Center"" ss:WrapText=""1""/>
   <Borders>
    <Border ss:Position=""Bottom"" ss:LineStyle=""Double"" ss:Weight=""3""/>
    <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
   </Borders>
   <Font ss:FontName=""Arial Narrow"" x:Family=""Swiss"" ss:Size=""8"" ss:Italic=""1""/>
  </Style>
  <Style ss:ID=""m2353907575768"">
   <Alignment ss:Horizontal=""Left"" ss:Vertical=""Center"" ss:WrapText=""1""/>
   <Borders>
    <Border ss:Position=""Bottom"" ss:LineStyle=""Double"" ss:Weight=""3""/>
    <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
   </Borders>
   <Font ss:FontName=""Arial Narrow"" x:Family=""Swiss"" ss:Size=""7"" ss:Italic=""1""/>
  </Style>
  <Style ss:ID=""m2353907575788"">
   <Alignment ss:Horizontal=""Left"" ss:Vertical=""Center"" ss:WrapText=""1""/>
   <Borders>
    <Border ss:Position=""Bottom"" ss:LineStyle=""Double"" ss:Weight=""3""/>
    <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
   </Borders>
   <Font ss:FontName=""Arial Narrow"" x:Family=""Swiss"" ss:Size=""8"" ss:Italic=""1""/>
  </Style>
  <Style ss:ID=""m2353907575808"">
   <Alignment ss:Horizontal=""Left"" ss:Vertical=""Center"" ss:WrapText=""1""/>
   <Borders>
    <Border ss:Position=""Bottom"" ss:LineStyle=""Double"" ss:Weight=""3""/>
    <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
   </Borders>
   <Font ss:FontName=""Arial Narrow"" x:Family=""Swiss"" ss:Size=""7"" ss:Italic=""1""/>
  </Style>
  <Style ss:ID=""m2353907575828"">
   <Alignment ss:Horizontal=""Left"" ss:Vertical=""Center"" ss:WrapText=""1""/>
   <Borders>
    <Border ss:Position=""Bottom"" ss:LineStyle=""Double"" ss:Weight=""3""/>
    <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Right"" ss:LineStyle=""Double"" ss:Weight=""3""/>
    <Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
   </Borders>
   <Font ss:FontName=""Arial Narrow"" x:Family=""Swiss"" ss:Size=""8"" ss:Italic=""1""/>
  </Style>
  <Style ss:ID=""m2353907575520"">
   <Alignment ss:Horizontal=""Left"" ss:Vertical=""Center"" ss:WrapText=""1""/>
   <Borders>
    <Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Left"" ss:LineStyle=""Double"" ss:Weight=""3""/>
    <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
   </Borders>
   <Font ss:FontName=""Arial Narrow"" x:Family=""Swiss"" ss:Size=""7"" ss:Bold=""1""/>
  </Style>
  <Style ss:ID=""m2353907575540"">
   <Alignment ss:Horizontal=""Left"" ss:Vertical=""Center"" ss:WrapText=""1""/>
   <Borders>
    <Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
   </Borders>
   <Font ss:FontName=""Arial Narrow"" x:Family=""Swiss"" ss:Size=""8"" ss:Italic=""1""/>
  </Style>
  <Style ss:ID=""m2353907575560"">
   <Alignment ss:Horizontal=""Left"" ss:Vertical=""Center"" ss:WrapText=""1""/>
   <Borders>
    <Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
   </Borders>
   <Font ss:FontName=""Arial Narrow"" x:Family=""Swiss"" ss:Size=""9""/>
  </Style>
  <Style ss:ID=""m2353907575580"">
   <Alignment ss:Horizontal=""Left"" ss:Vertical=""Center"" ss:WrapText=""1""/>
   <Borders>
    <Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
   </Borders>
   <Font ss:FontName=""Arial Narrow"" x:Family=""Swiss"" ss:Size=""8"" ss:Italic=""1""/>
  </Style>
  <Style ss:ID=""m2353907575600"">
   <Alignment ss:Horizontal=""Left"" ss:Vertical=""Center"" ss:WrapText=""1""/>
   <Borders>
    <Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
   </Borders>
   <Font ss:FontName=""Arial Narrow"" x:Family=""Swiss"" ss:Size=""9""/>
  </Style>
  <Style ss:ID=""m2353907575620"">
   <Alignment ss:Horizontal=""Left"" ss:Vertical=""Center"" ss:WrapText=""1""/>
   <Borders>
    <Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Right"" ss:LineStyle=""Double"" ss:Weight=""3""/>
    <Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
   </Borders>
   <Font ss:FontName=""Arial Narrow"" x:Family=""Swiss"" ss:Size=""8"" ss:Italic=""1""/>
  </Style>
  <Style ss:ID=""m2353907574896"">
   <Alignment ss:Horizontal=""Left"" ss:Vertical=""Center"" ss:WrapText=""1""/>
   <Borders>
    <Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Left"" ss:LineStyle=""Double"" ss:Weight=""3""/>
    <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Top"" ss:LineStyle=""Double"" ss:Weight=""3""/>
   </Borders>
   <Font ss:FontName=""Arial Narrow"" x:Family=""Swiss"" ss:Size=""7"" ss:Italic=""1""/>
  </Style>
  <Style ss:ID=""m2353907574916"">
   <Alignment ss:Horizontal=""Left"" ss:Vertical=""Center"" ss:WrapText=""1""/>
   <Borders>
    <Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Top"" ss:LineStyle=""Double"" ss:Weight=""3""/>
   </Borders>
   <Font ss:FontName=""Arial Narrow"" x:Family=""Swiss"" ss:Size=""9""/>
  </Style>
  <Style ss:ID=""m2353907574936"">
   <Alignment ss:Horizontal=""Left"" ss:Vertical=""Center"" ss:WrapText=""1""/>
   <Borders>
    <Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Top"" ss:LineStyle=""Double"" ss:Weight=""3""/>
   </Borders>
   <Font ss:FontName=""Arial Narrow"" x:Family=""Swiss"" ss:Size=""8"" ss:Bold=""1""/>
  </Style>
  <Style ss:ID=""m2353907574956"">
   <Alignment ss:Horizontal=""Left"" ss:Vertical=""Center"" ss:WrapText=""1""/>
   <Borders>
    <Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Top"" ss:LineStyle=""Double"" ss:Weight=""3""/>
   </Borders>
   <Font ss:FontName=""Arial Narrow"" x:Family=""Swiss"" ss:Size=""8"" ss:Italic=""1""/>
  </Style>
  <Style ss:ID=""m2353907574976"">
   <Alignment ss:Horizontal=""Left"" ss:Vertical=""Center"" ss:WrapText=""1""/>
   <Borders>
    <Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Top"" ss:LineStyle=""Double"" ss:Weight=""3""/>
   </Borders>
   <Font ss:FontName=""Arial Narrow"" x:Family=""Swiss"" ss:Italic=""1""/>
  </Style>
  <Style ss:ID=""m2353907574996"">
   <Alignment ss:Horizontal=""Left"" ss:Vertical=""Center"" ss:WrapText=""1""/>
   <Borders>
    <Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Right"" ss:LineStyle=""Double"" ss:Weight=""3""/>
    <Border ss:Position=""Top"" ss:LineStyle=""Double"" ss:Weight=""3""/>
   </Borders>
   <Font ss:FontName=""Arial Narrow"" x:Family=""Swiss"" ss:Size=""8"" ss:Italic=""1""/>
  </Style>
  <Style ss:ID=""m2353907575016"">
   <Alignment ss:Horizontal=""Left"" ss:Vertical=""Center"" ss:WrapText=""1""/>
   <Borders>
    <Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Left"" ss:LineStyle=""Double"" ss:Weight=""3""/>
    <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
   </Borders>
   <Font ss:FontName=""Arial Narrow"" x:Family=""Swiss"" ss:Size=""8"" ss:Bold=""1""/>
  </Style>
  <Style ss:ID=""m2353907575036"">
   <Alignment ss:Horizontal=""Left"" ss:Vertical=""Center"" ss:WrapText=""1""/>
   <Borders>
    <Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
   </Borders>
   <Font ss:FontName=""Arial Narrow"" x:Family=""Swiss"" ss:Size=""8"" ss:Bold=""1""/>
  </Style>
  <Style ss:ID=""m2353907575056"">
   <Alignment ss:Horizontal=""Left"" ss:Vertical=""Center"" ss:WrapText=""1""/>
   <Borders>
    <Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
   </Borders>
   <Font ss:FontName=""Arial Narrow"" x:Family=""Swiss"" ss:Size=""8"" ss:Italic=""1""/>
  </Style>
  <Style ss:ID=""m2353907575076"">
   <Alignment ss:Horizontal=""Left"" ss:Vertical=""Center"" ss:WrapText=""1""/>
   <Borders>
    <Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Right"" ss:LineStyle=""Double"" ss:Weight=""3""/>
    <Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
   </Borders>
   <Font ss:FontName=""Arial Narrow"" x:Family=""Swiss"" ss:Size=""8"" ss:Italic=""1""/>
  </Style>
  <Style ss:ID=""m2353907236560"">
   <Alignment ss:Horizontal=""Center"" ss:Vertical=""Top"" ss:WrapText=""1""/>
   <Borders>
    <Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
   </Borders>
   <Font ss:FontName=""Arial Narrow"" x:Family=""Swiss"" ss:Size=""8"" ss:Bold=""1""/>
  </Style>
  <Style ss:ID=""m2353907236580"">
   <Alignment ss:Horizontal=""Center"" ss:Vertical=""Top"" ss:WrapText=""1""/>
   <Borders>
    <Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
   </Borders>
   <Font ss:FontName=""Arial Narrow"" x:Family=""Swiss"" ss:Size=""8"" ss:Bold=""1""/>
  </Style>
  <Style ss:ID=""m2353907236600"">
   <Alignment ss:Horizontal=""Center"" ss:Vertical=""Top"" ss:WrapText=""1""/>
   <Borders>
    <Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
   </Borders>
   <Font ss:FontName=""Arial Narrow"" x:Family=""Swiss"" ss:Size=""8"" ss:Bold=""1""/>
  </Style>
  <Style ss:ID=""m2353907236620"">
   <Alignment ss:Horizontal=""Center"" ss:Vertical=""Top"" ss:WrapText=""1""/>
   <Borders>
    <Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
   </Borders>
   <Font ss:FontName=""Arial Narrow"" x:Family=""Swiss"" ss:Size=""8"" ss:Bold=""1""/>
  </Style>
  <Style ss:ID=""m2353907236640"">
   <Alignment ss:Horizontal=""Center"" ss:Vertical=""Top"" ss:WrapText=""1""/>
   <Borders>
    <Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
   </Borders>
   <Font ss:FontName=""Arial Narrow"" x:Family=""Swiss"" ss:Size=""8"" ss:Bold=""1""/>
  </Style>
  <Style ss:ID=""m2353907236660"">
   <Alignment ss:Horizontal=""Center"" ss:Vertical=""Top"" ss:WrapText=""1""/>
   <Borders>
    <Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Right"" ss:LineStyle=""Double"" ss:Weight=""3""/>
    <Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
   </Borders>
   <Font ss:FontName=""Arial Narrow"" x:Family=""Swiss"" ss:Size=""8"" ss:Bold=""1""/>
  </Style>
  <Style ss:ID=""m2353907238016"">
   <Alignment ss:Horizontal=""Center"" ss:Vertical=""Top"" ss:WrapText=""1""/>
   <Borders>
    <Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
   </Borders>
   <Font ss:FontName=""Arial Narrow"" x:Family=""Swiss"" ss:Size=""8"" ss:Bold=""1""/>
  </Style>
  <Style ss:ID=""m2353907238036"">
   <Alignment ss:Horizontal=""Center"" ss:Vertical=""Top"" ss:WrapText=""1""/>
   <Borders>
    <Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
   </Borders>
   <Font ss:FontName=""Arial Narrow"" x:Family=""Swiss"" ss:Size=""8"" ss:Bold=""1""/>
  </Style>
  <Style ss:ID=""m2353907238056"">
   <Alignment ss:Horizontal=""Center"" ss:Vertical=""Top"" ss:WrapText=""1""/>
   <Borders>
    <Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
   </Borders>
   <Font ss:FontName=""Arial Narrow"" x:Family=""Swiss"" ss:Size=""8"" ss:Bold=""1""/>
  </Style>
  <Style ss:ID=""m2353907238076"">
   <Alignment ss:Horizontal=""Center"" ss:Vertical=""Top"" ss:WrapText=""1""/>
   <Borders>
    <Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
   </Borders>
   <Font ss:FontName=""Arial Narrow"" x:Family=""Swiss"" ss:Size=""8"" ss:Bold=""1""/>
  </Style>
  <Style ss:ID=""m2353907238096"">
   <Alignment ss:Horizontal=""Center"" ss:Vertical=""Top"" ss:WrapText=""1""/>
   <Borders>
    <Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
   </Borders>
   <Font ss:FontName=""Arial Narrow"" x:Family=""Swiss"" ss:Size=""8"" ss:Bold=""1""/>
  </Style>
  <Style ss:ID=""m2353907238116"">
   <Alignment ss:Horizontal=""Center"" ss:Vertical=""Top"" ss:WrapText=""1""/>
   <Borders>
    <Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
   </Borders>
   <Font ss:FontName=""Arial Narrow"" x:Family=""Swiss"" ss:Size=""8"" ss:Bold=""1""/>
  </Style>
  <Style ss:ID=""m2353907237184"">
   <Alignment ss:Horizontal=""Center"" ss:Vertical=""Top"" ss:WrapText=""1""/>
   <Borders>
    <Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
   </Borders>
   <Font ss:FontName=""Arial Narrow"" x:Family=""Swiss"" ss:Size=""8"" ss:Bold=""1""/>
  </Style>
  <Style ss:ID=""m2353907237204"">
   <Alignment ss:Horizontal=""Center"" ss:Vertical=""Top"" ss:WrapText=""1""/>
   <Borders>
    <Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
   </Borders>
   <Font ss:FontName=""Arial Narrow"" x:Family=""Swiss"" ss:Size=""8"" ss:Bold=""1""/>
  </Style>
  <Style ss:ID=""m2353907237224"">
   <Alignment ss:Horizontal=""Center"" ss:Vertical=""Top"" ss:WrapText=""1""/>
   <Borders>
    <Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
   </Borders>
   <Font ss:FontName=""Arial Narrow"" x:Family=""Swiss"" ss:Size=""8"" ss:Bold=""1""/>
  </Style>
  <Style ss:ID=""m2353907237244"">
   <Alignment ss:Horizontal=""Center"" ss:Vertical=""Top"" ss:WrapText=""1""/>
   <Borders>
    <Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
   </Borders>
   <Font ss:FontName=""Arial Narrow"" x:Family=""Swiss"" ss:Size=""8"" ss:Bold=""1""/>
  </Style>
  <Style ss:ID=""m2353907237264"">
   <Alignment ss:Horizontal=""Center"" ss:Vertical=""Top"" ss:WrapText=""1""/>
   <Borders>
    <Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
   </Borders>
   <Font ss:FontName=""Arial Narrow"" x:Family=""Swiss"" ss:Size=""8"" ss:Bold=""1""/>
  </Style>
  <Style ss:ID=""m2353907237284"">
   <Alignment ss:Horizontal=""Center"" ss:Vertical=""Top"" ss:WrapText=""1""/>
   <Borders>
    <Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
   </Borders>
   <Font ss:FontName=""Arial Narrow"" x:Family=""Swiss"" ss:Size=""8"" ss:Bold=""1""/>
  </Style>
  <Style ss:ID=""m2353907238224"">
   <Alignment ss:Horizontal=""Center"" ss:Vertical=""Center"" ss:WrapText=""1""/>
   <Borders>
    <Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Left"" ss:LineStyle=""Double"" ss:Weight=""3""/>
    <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
   </Borders>
   <Font ss:FontName=""Arial Narrow"" x:Family=""Swiss"" ss:Size=""8"" ss:Bold=""1""/>
  </Style>
  <Style ss:ID=""m2353907238244"">
   <Alignment ss:Horizontal=""Center"" ss:Vertical=""Top"" ss:WrapText=""1""/>
   <Borders>
    <Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
   </Borders>
   <Font ss:FontName=""Arial Narrow"" x:Family=""Swiss"" ss:Size=""8"" ss:Bold=""1""/>
  </Style>
  <Style ss:ID=""m2353907238264"">
   <Alignment ss:Horizontal=""Center"" ss:Vertical=""Top"" ss:WrapText=""1""/>
   <Borders>
    <Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
   </Borders>
   <Font ss:FontName=""Arial Narrow"" x:Family=""Swiss"" ss:Size=""8"" ss:Bold=""1""/>
  </Style>
  <Style ss:ID=""m2353907238284"">
   <Alignment ss:Horizontal=""Center"" ss:Vertical=""Top"" ss:WrapText=""1""/>
   <Borders>
    <Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
   </Borders>
   <Font ss:FontName=""Arial Narrow"" x:Family=""Swiss"" ss:Size=""8"" ss:Bold=""1""/>
  </Style>
  <Style ss:ID=""m2353907238304"">
   <Alignment ss:Horizontal=""Center"" ss:Vertical=""Top"" ss:WrapText=""1""/>
   <Borders>
    <Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
   </Borders>
   <Font ss:FontName=""Arial Narrow"" x:Family=""Swiss"" ss:Size=""8"" ss:Bold=""1""/>
  </Style>
  <Style ss:ID=""m2353907238324"">
   <Alignment ss:Horizontal=""Center"" ss:Vertical=""Top"" ss:WrapText=""1""/>
   <Borders>
    <Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
   </Borders>
   <Font ss:FontName=""Arial Narrow"" x:Family=""Swiss"" ss:Size=""8"" ss:Bold=""1""/>
  </Style>
  <Style ss:ID=""m2353907236768"">
   <Alignment ss:Horizontal=""Center"" ss:Vertical=""Center"" ss:Rotate=""90""
    ss:WrapText=""1""/>
   <Borders>
    <Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Top"" ss:LineStyle=""Double"" ss:Weight=""3""/>
   </Borders>
   <Font ss:FontName=""Arial Narrow"" x:Family=""Swiss"" ss:Size=""8"" ss:Color=""#000000""
    ss:Italic=""1""/>
  </Style>
  <Style ss:ID=""m2353907236788"">
   <Alignment ss:Horizontal=""Center"" ss:Vertical=""Top"" ss:Rotate=""90""
    ss:WrapText=""1""/>
   <Borders>
    <Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Top"" ss:LineStyle=""Double"" ss:Weight=""3""/>
   </Borders>
   <Font ss:FontName=""Arial Narrow"" x:Family=""Swiss"" ss:Size=""8"" ss:Color=""#000000""
    ss:Italic=""1""/>
  </Style>
  <Style ss:ID=""m2353907236808"">
   <Alignment ss:Horizontal=""Center"" ss:Vertical=""Center"" ss:WrapText=""1""/>
   <Borders>
    <Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Top"" ss:LineStyle=""Double"" ss:Weight=""3""/>
   </Borders>
   <Font ss:FontName=""Arial Narrow"" x:Family=""Swiss"" ss:Size=""8"" ss:Italic=""1""/>
  </Style>
  <Style ss:ID=""m2353907236828"">
   <Alignment ss:Horizontal=""Center"" ss:Vertical=""Center"" ss:Rotate=""90""
    ss:WrapText=""1""/>
   <Borders>
    <Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Top"" ss:LineStyle=""Double"" ss:Weight=""3""/>
   </Borders>
   <Font ss:FontName=""Arial Narrow"" x:Family=""Swiss"" ss:Size=""8"" ss:Color=""#000000""
    ss:Italic=""1""/>
  </Style>
  <Style ss:ID=""m2353907236848"">
   <Alignment ss:Horizontal=""Center"" ss:Vertical=""Center"" ss:Rotate=""90""
    ss:WrapText=""1""/>
   <Borders>
    <Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Top"" ss:LineStyle=""Double"" ss:Weight=""3""/>
   </Borders>
   <Font ss:FontName=""Arial Narrow"" x:Family=""Swiss"" ss:Size=""9"" ss:Italic=""1""/>
  </Style>
  <Style ss:ID=""m2353907236868"">
   <Alignment ss:Horizontal=""Center"" ss:Vertical=""Center"" ss:Rotate=""90""
    ss:WrapText=""1""/>
   <Borders>
    <Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Right"" ss:LineStyle=""Double"" ss:Weight=""3""/>
    <Border ss:Position=""Top"" ss:LineStyle=""Double"" ss:Weight=""3""/>
   </Borders>
   <Font ss:FontName=""Arial Narrow"" x:Family=""Swiss"" ss:Size=""8"" ss:Italic=""1""/>
  </Style>
  <Style ss:ID=""m2353907236888"">
   <Alignment ss:Horizontal=""Center"" ss:Vertical=""Center"" ss:WrapText=""1""/>
   <Borders>
    <Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
   </Borders>
   <Font ss:FontName=""Arial Narrow"" x:Family=""Swiss"" ss:Size=""8"" ss:Italic=""1""/>
  </Style>
  <Style ss:ID=""m2353907236908"">
   <Alignment ss:Horizontal=""Center"" ss:Vertical=""Center"" ss:Rotate=""90""
    ss:WrapText=""1""/>
   <Borders>
    <Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
   </Borders>
   <Font ss:FontName=""Arial Narrow"" x:Family=""Swiss"" ss:Size=""8"" ss:Italic=""1""/>
  </Style>
  <Style ss:ID=""m2353907236928"">
   <Alignment ss:Horizontal=""Center"" ss:Vertical=""Center"" ss:Rotate=""90""
    ss:WrapText=""1""/>
   <Borders>
    <Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
   </Borders>
   <Font ss:FontName=""Arial Narrow"" x:Family=""Swiss"" ss:Size=""8"" ss:Italic=""1""/>
  </Style>
  <Style ss:ID=""m2353907234688"">
   <Alignment ss:Horizontal=""Center"" ss:Vertical=""Center"" ss:Rotate=""90""
    ss:WrapText=""1""/>
   <Borders>
    <Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Top"" ss:LineStyle=""Double"" ss:Weight=""3""/>
   </Borders>
   <Font ss:FontName=""Arial Narrow"" x:Family=""Swiss"" ss:Size=""8"" ss:Italic=""1""/>
  </Style>
  <Style ss:ID=""m2353907234708"">
   <Alignment ss:Horizontal=""Center"" ss:Vertical=""Center"" ss:Rotate=""90""
    ss:WrapText=""1""/>
   <Borders>
    <Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Top"" ss:LineStyle=""Double"" ss:Weight=""3""/>
   </Borders>
   <Font ss:FontName=""Arial Narrow"" x:Family=""Swiss"" ss:Size=""8"" ss:Italic=""1""/>
  </Style>
  <Style ss:ID=""m2353907234728"">
   <Alignment ss:Horizontal=""Center"" ss:Vertical=""Center"" ss:Rotate=""90""
    ss:WrapText=""1""/>
   <Borders>
    <Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Top"" ss:LineStyle=""Double"" ss:Weight=""3""/>
   </Borders>
   <Font ss:FontName=""Arial Narrow"" x:Family=""Swiss"" ss:Size=""8"" ss:Italic=""1""/>
  </Style>
  <Style ss:ID=""m2353907234748"">
   <Alignment ss:Horizontal=""Center"" ss:Vertical=""Center"" ss:Rotate=""90""
    ss:WrapText=""1""/>
   <Borders>
    <Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Top"" ss:LineStyle=""Double"" ss:Weight=""3""/>
   </Borders>
   <Font ss:FontName=""Arial Narrow"" x:Family=""Swiss"" ss:Size=""8"" ss:Italic=""1""/>
  </Style>
  <Style ss:ID=""m2353907234768"">
   <Alignment ss:Horizontal=""Center"" ss:Vertical=""Center"" ss:Rotate=""90""
    ss:WrapText=""1""/>
   <Borders>
    <Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Top"" ss:LineStyle=""Double"" ss:Weight=""3""/>
   </Borders>
   <Font ss:FontName=""Arial Narrow"" x:Family=""Swiss"" ss:Size=""8"" ss:Italic=""1""/>
  </Style>
  <Style ss:ID=""m2353907234788"">
   <Alignment ss:Horizontal=""Center"" ss:Vertical=""Center"" ss:Rotate=""90""
    ss:WrapText=""1""/>
   <Borders>
    <Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Top"" ss:LineStyle=""Double"" ss:Weight=""3""/>
   </Borders>
   <Font ss:FontName=""Arial Narrow"" x:Family=""Swiss"" ss:Size=""8"" ss:Italic=""1""/>
  </Style>
  <Style ss:ID=""m2353907234916"">
   <Alignment ss:Horizontal=""Center"" ss:Vertical=""Center"" ss:Rotate=""90""
    ss:WrapText=""1""/>
   <Borders>
    <Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Left"" ss:LineStyle=""Double"" ss:Weight=""3""/>
    <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Top"" ss:LineStyle=""Double"" ss:Weight=""3""/>
   </Borders>
   <Font ss:FontName=""Arial Narrow"" x:Family=""Swiss"" ss:Size=""8"" ss:Italic=""1""/>
  </Style>
  <Style ss:ID=""m2353907234936"">
   <Alignment ss:Horizontal=""Center"" ss:Vertical=""Center"" ss:Rotate=""90""
    ss:WrapText=""1""/>
   <Borders>
    <Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Top"" ss:LineStyle=""Double"" ss:Weight=""3""/>
   </Borders>
   <Font ss:FontName=""Arial Narrow"" x:Family=""Swiss"" ss:Size=""8"" ss:Italic=""1""/>
  </Style>
  <Style ss:ID=""m2353907234956"">
   <Alignment ss:Horizontal=""Center"" ss:Vertical=""Center"" ss:Rotate=""90""
    ss:WrapText=""1""/>
   <Borders>
    <Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Top"" ss:LineStyle=""Double"" ss:Weight=""3""/>
   </Borders>
   <Font ss:FontName=""Arial Narrow"" x:Family=""Swiss"" ss:Size=""8"" ss:Italic=""1""/>
  </Style>
  <Style ss:ID=""m2353907234976"">
   <Alignment ss:Horizontal=""Center"" ss:Vertical=""Center"" ss:Rotate=""90""
    ss:WrapText=""1""/>
   <Borders>
    <Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Top"" ss:LineStyle=""Double"" ss:Weight=""3""/>
   </Borders>
   <Font ss:FontName=""Arial Narrow"" x:Family=""Swiss"" ss:Size=""9""/>
  </Style>
  <Style ss:ID=""m2353907234996"">
   <Alignment ss:Horizontal=""Center"" ss:Vertical=""Top"" ss:Rotate=""90""
    ss:WrapText=""1""/>
   <Borders>
    <Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Top"" ss:LineStyle=""Double"" ss:Weight=""3""/>
   </Borders>
   <Font ss:FontName=""Arial Narrow"" x:Family=""Swiss"" ss:Size=""8"" ss:Italic=""1""/>
  </Style>
  <Style ss:ID=""m2353907235016"">
   <Alignment ss:Horizontal=""Center"" ss:Vertical=""Center"" ss:Rotate=""90""
    ss:WrapText=""1""/>
   <Borders>
    <Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Top"" ss:LineStyle=""Double"" ss:Weight=""3""/>
   </Borders>
   <Font ss:FontName=""Arial Narrow"" x:Family=""Swiss"" ss:Size=""8"" ss:Italic=""1""/>
  </Style>
  <Style ss:ID=""m2353907235036"">
   <Alignment ss:Horizontal=""Center"" ss:Vertical=""Center"" ss:Rotate=""90""
    ss:WrapText=""1""/>
   <Borders>
    <Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Top"" ss:LineStyle=""Double"" ss:Weight=""3""/>
   </Borders>
   <Font ss:FontName=""Arial Narrow"" x:Family=""Swiss"" ss:Size=""8"" ss:Italic=""1""/>
  </Style>
  <Style ss:ID=""m2353907235056"">
   <Alignment ss:Horizontal=""Center"" ss:Vertical=""Center"" ss:Rotate=""90""
    ss:WrapText=""1""/>
   <Borders>
    <Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Top"" ss:LineStyle=""Double"" ss:Weight=""3""/>
   </Borders>
   <Font ss:FontName=""Arial Narrow"" x:Family=""Swiss"" ss:Size=""8"" ss:Italic=""1""/>
  </Style>
  <Style ss:ID=""m2353907235076"">
   <Alignment ss:Horizontal=""Center"" ss:Vertical=""Center"" ss:Rotate=""90""
    ss:WrapText=""1""/>
   <Borders>
    <Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Top"" ss:LineStyle=""Double"" ss:Weight=""3""/>
   </Borders>
   <Font ss:FontName=""Arial Narrow"" x:Family=""Swiss"" ss:Size=""8"" ss:Italic=""1""/>
  </Style>
  <Style ss:ID=""m2353907237888"">
   <Alignment ss:Horizontal=""Left"" ss:Vertical=""Center""/>
   <Borders>
    <Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
   </Borders>
  </Style>
  <Style ss:ID=""s63"">
   <Alignment ss:Horizontal=""Left"" ss:Vertical=""Top"" ss:WrapText=""1""/>
   <Borders/>
   <Font ss:FontName=""Times New Roman"" x:CharSet=""204"" x:Family=""Roman""
    ss:Color=""#808080"" ss:Bold=""1""/>
   <Interior ss:Color=""#FFFFFF"" ss:Pattern=""Gray25"" ss:PatternColor=""#000000""/>
  </Style>
  <Style ss:ID=""s64"">
   <Font ss:FontName=""Times"" x:CharSet=""204"" x:Family=""Roman"" ss:Color=""#808080""/>
  </Style>
  <Style ss:ID=""s66"">
   <Alignment ss:Horizontal=""Center"" ss:Vertical=""Center"" ss:WrapText=""1""/>
   <Font ss:FontName=""Arial"" x:CharSet=""204"" x:Family=""Swiss"" ss:Size=""11""
    ss:Color=""#808080"" ss:Bold=""1""/>
  </Style>
  <Style ss:ID=""s68"">
   <Alignment ss:Horizontal=""Right"" ss:Vertical=""Top"" ss:WrapText=""1""/>
   <Borders/>
   <Font ss:FontName=""Times New Roman"" x:CharSet=""204"" x:Family=""Roman""
    ss:Color=""#808080"" ss:Bold=""1""/>
   <Interior ss:Color=""#FFFFFF"" ss:Pattern=""Gray25"" ss:PatternColor=""#000000""/>
  </Style>
  <Style ss:ID=""s69"">
   <Interior/>
  </Style>
  <Style ss:ID=""s70"">
   <Alignment ss:Vertical=""Top"" ss:WrapText=""1""/>
   <Borders/>
   <Font ss:FontName=""Times New Roman"" x:CharSet=""204"" x:Family=""Roman""
    ss:Color=""#808080"" ss:Bold=""1""/>
   <Interior/>
  </Style>
  <Style ss:ID=""s72"">
   <Alignment ss:Horizontal=""Center"" ss:Vertical=""Top"" ss:WrapText=""1""/>
   <Font ss:FontName=""Arial"" x:Family=""Swiss"" ss:Size=""12"" ss:Bold=""1""/>
  </Style>
  <Style ss:ID=""s73"">
   <Alignment ss:Vertical=""Top"" ss:WrapText=""1""/>
   <Font ss:FontName=""Times New Roman"" x:Family=""Roman"" ss:Size=""9"" ss:Italic=""1""/>
  </Style>
  <Style ss:ID=""s74"">
   <Alignment ss:Vertical=""Bottom""/>
  </Style>
  <Style ss:ID=""s82"">
   <Alignment ss:Vertical=""Top"" ss:WrapText=""1""/>
   <Font ss:FontName=""Arial"" x:Family=""Swiss"" ss:Size=""9"" ss:Italic=""1""/>
  </Style>
  <Style ss:ID=""s84"">
   <Alignment ss:Horizontal=""Left"" ss:Vertical=""Bottom""/>
  </Style>
  <Style ss:ID=""s86"">
   <Alignment ss:Horizontal=""Left"" ss:Vertical=""Bottom""/>
   <Borders>
    <Border ss:Position=""Bottom"" ss:LineStyle=""Double"" ss:Weight=""3""/>
   </Borders>
  </Style>
  <Style ss:ID=""s173"">
   <Alignment ss:Horizontal=""Center"" ss:Vertical=""Center"" ss:Rotate=""90""
    ss:WrapText=""1""/>
   <Borders>
    <Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
   </Borders>
   <Font ss:FontName=""Arial Narrow"" x:Family=""Swiss"" ss:Size=""8"" ss:Color=""#000000""
    ss:Italic=""1""/>
  </Style>
  <Style ss:ID=""s180"">
   <Alignment ss:Horizontal=""Center"" ss:Vertical=""Center"" ss:WrapText=""1""/>
   <Borders>
    <Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
    <Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>
   </Borders>
   <Font ss:FontName=""Arial Narrow"" x:Family=""Swiss"" ss:Size=""8"" ss:Color=""#000000""/>
  </Style>
 </Styles>
 <Worksheet ss:Name=""PassagePlan"">
  <Names>
   <NamedRange ss:Name=""Print_Titles"" ss:RefersTo=""=PassagePlan!R1:R2""/>
  </Names>
  <Table ss:ExpandedColumnCount=""31"" ss:ExpandedRowCount=""220"" x:FullColumns=""1""
   x:FullRows=""1"" ss:DefaultColumnWidth=""17.25"">
   <Column ss:Index=""2"" ss:AutoFitWidth=""0"" ss:Width=""76.5""/>
   <Column ss:AutoFitWidth=""0"" ss:Width=""33.75""/>
   <Column ss:Index=""5"" ss:AutoFitWidth=""0"" ss:Width=""19.5""/>
   <Column ss:AutoFitWidth=""0"" ss:Width=""28.5""/>
   <Column ss:AutoFitWidth=""0"" ss:Width=""15""/>
   <Column ss:Index=""10"" ss:AutoFitWidth=""0"" ss:Width=""42"" ss:Span=""1""/>
   <Column ss:Index=""12"" ss:AutoFitWidth=""0"" ss:Width=""30.75""/>
   <Column ss:AutoFitWidth=""0"" ss:Width=""20.25""/>
   <Column ss:Index=""16"" ss:AutoFitWidth=""0"" ss:Width=""25.5""/>
   <Column ss:AutoFitWidth=""0"" ss:Width=""31.5""/>
   <Column ss:Index=""19"" ss:AutoFitWidth=""0"" ss:Width=""40.5""/>
   <Column ss:AutoFitWidth=""0"" ss:Width=""16.5""/>
   <Column ss:Index=""22"" ss:AutoFitWidth=""0"" ss:Width=""21""/>
   <Column ss:AutoFitWidth=""0"" ss:Width=""20.25""/>
   <Column ss:AutoFitWidth=""0"" ss:Width=""18.75""/>
   <Column ss:Index=""26"" ss:AutoFitWidth=""0"" ss:Width=""15""/>
   <Column ss:AutoFitWidth=""0"" ss:Width=""15.75"" ss:Span=""1""/>
   <Column ss:Index=""29"" ss:AutoFitWidth=""0"" ss:Width=""45.75""/>
   <Column ss:AutoFitWidth=""0"" ss:Width=""57.75""/>
   <Column ss:AutoFitWidth=""0"" ss:Width=""49.5""/>
   <Row ss:AutoFitHeight=""0"" ss:Height=""28.5"">
    <Cell ss:Index=""2"" ss:MergeAcross=""2"" ss:StyleID=""s63""><Data ss:Type=""String"">Rev.  0&#10;Data rev. 28/02/2023</Data><NamedCell
      ss:Name=""Print_Titles""/></Cell>
    <Cell ss:StyleID=""s64""><NamedCell ss:Name=""Print_Titles""/></Cell>
    <Cell ss:StyleID=""s64""><NamedCell ss:Name=""Print_Titles""/></Cell>
    <Cell ss:StyleID=""s64""><NamedCell ss:Name=""Print_Titles""/></Cell>
    <Cell ss:StyleID=""s64""><NamedCell ss:Name=""Print_Titles""/></Cell>
    <Cell ss:StyleID=""s64""><NamedCell ss:Name=""Print_Titles""/></Cell>
    <Cell ss:StyleID=""s64""><NamedCell ss:Name=""Print_Titles""/></Cell>
    <Cell ss:StyleID=""s64""><NamedCell ss:Name=""Print_Titles""/></Cell>
    <Cell ss:MergeAcross=""6"" ss:StyleID=""s66""><Data ss:Type=""String"">Voyage  Planning</Data><NamedCell
      ss:Name=""Print_Titles""/></Cell>
    <Cell ss:StyleID=""s64""><NamedCell ss:Name=""Print_Titles""/></Cell>
    <Cell ss:StyleID=""s64""><NamedCell ss:Name=""Print_Titles""/></Cell>
    <Cell ss:StyleID=""s64""><NamedCell ss:Name=""Print_Titles""/></Cell>
    <Cell ss:StyleID=""s64""><NamedCell ss:Name=""Print_Titles""/></Cell>
    <Cell ss:StyleID=""s64""><NamedCell ss:Name=""Print_Titles""/></Cell>
    <Cell ss:StyleID=""s64""><NamedCell ss:Name=""Print_Titles""/></Cell>
    <Cell ss:Index=""27"" ss:MergeAcross=""3"" ss:StyleID=""s68""><Data ss:Type=""String"">SMS-F-041</Data><NamedCell
      ss:Name=""Print_Titles""/></Cell>
   </Row>
   <Row ss:AutoFitHeight=""0"">
    <Cell ss:Index=""27"" ss:StyleID=""s69""><NamedCell ss:Name=""Print_Titles""/></Cell>
    <Cell ss:StyleID=""s70""><NamedCell ss:Name=""Print_Titles""/></Cell>
    <Cell ss:StyleID=""s70""><NamedCell ss:Name=""Print_Titles""/></Cell>
    <Cell ss:StyleID=""s70""><NamedCell ss:Name=""Print_Titles""/></Cell>
    <Cell ss:StyleID=""s70""><NamedCell ss:Name=""Print_Titles""/></Cell>
   </Row>
   <Row ss:AutoFitHeight=""0"" ss:Height=""15.9375"">
    <Cell ss:MergeAcross=""30"" ss:StyleID=""s72""><Data ss:Type=""String"">SAFETY MANAGEMENT SYSTEM</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight=""0"" ss:Height=""14.25"">
    <Cell ss:StyleID=""s73""/>
    <Cell ss:StyleID=""s74""/>
    <Cell ss:StyleID=""s74""/>
    <Cell ss:StyleID=""s74""/>
    <Cell ss:StyleID=""s74""/>
    <Cell ss:StyleID=""s74""/>
    <Cell ss:StyleID=""s74""/>
    <Cell ss:StyleID=""s74""/>
    <Cell ss:StyleID=""s74""/>
    <Cell ss:MergeAcross=""12"" ss:StyleID=""m2353907237888""><Data ss:Type=""String"">Ship’s Name ##shipName##</Data></Cell>
    <Cell ss:StyleID=""s74""/>
    <Cell ss:StyleID=""s74""/>
    <Cell ss:StyleID=""s74""/>
    <Cell ss:StyleID=""s74""/>
    <Cell ss:StyleID=""s74""/>
    <Cell ss:StyleID=""s74""/>
    <Cell ss:StyleID=""s74""/>
    <Cell ss:StyleID=""s74""/>
    <Cell ss:StyleID=""s74""/>
   </Row>
   <Row ss:AutoFitHeight=""0"" ss:Height=""15"">
    <Cell ss:StyleID=""s82""/>
    <Cell ss:MergeAcross=""29"" ss:StyleID=""s84""><Data ss:Type=""String"">Voy nr. ##voyage## From : ##portFrom## To : ##portTo##  Ets : ##ets## Eta ##eta##   Safety contour/Safety depth : ##safetContDepth##</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight=""0"" ss:Height=""15"">
    <Cell ss:StyleID=""s82""/>
    <Cell ss:MergeAcross=""29"" ss:StyleID=""s86""><Data ss:Type=""String"">Draft  : Fwd  ##draftFwd##m. Centre  ##draftMiddle##m.   Aft  ##draftAft##m.</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight=""0"" ss:Height=""20.25"">
    <Cell ss:MergeDown=""2"" ss:StyleID=""m2353907234916""><Data ss:Type=""String"">Way Points nr.</Data></Cell>
    <Cell ss:MergeDown=""2"" ss:StyleID=""m2353907234936""><Data ss:Type=""String"">Position – Lat.Long.</Data></Cell>
    <Cell ss:MergeDown=""2"" ss:StyleID=""m2353907234956""><Data ss:Type=""String"">Date</Data></Cell>
    <Cell ss:MergeAcross=""1"" ss:MergeDown=""2"" ss:StyleID=""m2353907234976""><ss:Data
      ss:Type=""String"" xmlns=""http://www.w3.org/TR/REC-html40"">&#10;<I><Font
        html:Size=""8"">Course</Font></I></ss:Data></Cell>
    <Cell ss:MergeDown=""2"" ss:StyleID=""m2353907234996""><Data ss:Type=""String"">Estimated Speed (safe speed)</Data></Cell>
    <Cell ss:MergeAcross=""1"" ss:MergeDown=""2"" ss:StyleID=""m2353907235016""><Data
      ss:Type=""String"">Miles between wayponts</Data></Cell>
    <Cell ss:MergeDown=""2"" ss:StyleID=""m2353907235036""><Data ss:Type=""String"">Miles to final destination</Data></Cell>
    <Cell ss:MergeDown=""2"" ss:StyleID=""m2353907235056""><Data ss:Type=""String"">Chart to be used nr.</Data></Cell>
    <Cell ss:MergeDown=""2"" ss:StyleID=""m2353907235076""><Data ss:Type=""String"">Ship’s reporting System (Ares, etc.)</Data></Cell>
    <Cell ss:MergeDown=""2"" ss:StyleID=""m2353907234688""><Data ss:Type=""String"">VTS station and  VHF Channel</Data></Cell>
    <Cell ss:MergeDown=""2"" ss:StyleID=""m2353907234708""><Data ss:Type=""String"">Marpol special areas Yes/No</Data></Cell>
    <Cell ss:MergeAcross=""1"" ss:MergeDown=""2"" ss:StyleID=""m2353907234728""><Data
      ss:Type=""String"">Squat effect</Data></Cell>
    <Cell ss:MergeDown=""2"" ss:StyleID=""m2353907234748""><Data ss:Type=""String"">Minimum underkeel clearance  included squat effect</Data></Cell>
    <Cell ss:MergeDown=""2"" ss:StyleID=""m2353907234768""><Data ss:Type=""String"">Safe distance from obstacles</Data></Cell>
    <Cell ss:MergeDown=""2"" ss:StyleID=""m2353907234788""><Data ss:Type=""String"">Ship positioning system (Gps, etc.)</Data></Cell>
    <Cell ss:MergeDown=""2"" ss:StyleID=""m2353907236768""><Data ss:Type=""String"">Current Direction/Speed</Data></Cell>
    <Cell ss:MergeAcross=""1"" ss:MergeDown=""2"" ss:StyleID=""m2353907236788""><Data
      ss:Type=""String"">Stand by in the engine Yes/No</Data></Cell>
    <Cell ss:MergeAcross=""6"" ss:StyleID=""m2353907236808""><ss:Data ss:Type=""String""
      xmlns=""http://www.w3.org/TR/REC-html40""><I>From pilot station to the key or  v.v</I><Font
       html:Size=""9"">.</Font></ss:Data></Cell>
    <Cell ss:MergeDown=""2"" ss:StyleID=""m2353907236828""><Data ss:Type=""String"">Refuge port of emergency anchorage</Data></Cell>
    <Cell ss:MergeDown=""2"" ss:StyleID=""m2353907236848""><Data ss:Type=""String"">Nautical publications to be used</Data></Cell>
    <Cell ss:MergeDown=""2"" ss:StyleID=""m2353907236868""><Data ss:Type=""String"">Notes</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight=""0"" ss:Height=""19.5"">
    <Cell ss:Index=""22"" ss:MergeAcross=""2"" ss:StyleID=""m2353907236888""><Data
      ss:Type=""String"">Tide time</Data></Cell>
    <Cell ss:MergeAcross=""1"" ss:MergeDown=""1"" ss:StyleID=""m2353907236908""><Data
      ss:Type=""String"">Miles pilot/berth (or  v.v.) </Data></Cell>
    <Cell ss:MergeAcross=""1"" ss:MergeDown=""1"" ss:StyleID=""m2353907236928""><Data
      ss:Type=""String"">Miles to the lock </Data></Cell>
   </Row>
   <Row ss:AutoFitHeight=""0"" ss:Height=""105"">
    <Cell ss:Index=""22"" ss:StyleID=""s173""><ss:Data ss:Type=""String""
      xmlns=""http://www.w3.org/TR/REC-html40""><I><Font html:Color=""#000000"">Low</Font></I><Font
       html:Color=""#000000""> </Font></ss:Data></Cell>
    <Cell ss:StyleID=""s173""><Data ss:Type=""String"">High</Data></Cell>
    <Cell ss:StyleID=""s173""><Data ss:Type=""String"">Slack</Data></Cell>");

		public static int DocStartXmlTotalRows =
			Regex.Matches(TableHeadRows, "Row", RegexOptions.IgnoreCase).Count;

		public const string TableHeadRows = @"
    <Row ss:AutoFitHeight=""0"" ss:Height=""20.25"">
    <Cell ss:MergeDown=""2"" ss:StyleID=""m1536069615444""><Data ss:Type=""String"">Way Points nr.</Data></Cell>
    <Cell ss:MergeDown=""2"" ss:StyleID=""m1536069615464""><Data ss:Type=""String"">Position – Lat.Long.</Data></Cell>
    <Cell ss:MergeDown=""2"" ss:StyleID=""m1536069615484""><Data ss:Type=""String"">Date</Data></Cell>
    <Cell ss:MergeAcross=""1"" ss:MergeDown=""2"" ss:StyleID=""m1536069615504""><ss:Data
      ss:Type=""String"" xmlns=""http://www.w3.org/TR/REC-html40"">&#10;<I><Font
        html:Size=""8"">Course</Font></I></ss:Data></Cell>
    <Cell ss:MergeDown=""2"" ss:StyleID=""m1536069615524""><Data ss:Type=""String"">Estimated Speed (safe speed)</Data></Cell>
    <Cell ss:MergeAcross=""1"" ss:MergeDown=""2"" ss:StyleID=""m1536116215076""><Data
      ss:Type=""String"">Miles between wayponts</Data></Cell>
    <Cell ss:MergeDown=""2"" ss:StyleID=""m1536116200664""><Data ss:Type=""String"">Miles to final destination</Data></Cell>
    <Cell ss:MergeDown=""2"" ss:StyleID=""m1536116198792""><Data ss:Type=""String"">Chart to be used nr.</Data></Cell>
    <Cell ss:MergeDown=""2"" ss:StyleID=""m1536116198812""><Data ss:Type=""String"">Ship’s reporting System (Ares, etc.)</Data></Cell>
    <Cell ss:MergeDown=""2"" ss:StyleID=""m1536116198832""><Data ss:Type=""String"">VTS station and  VHF Channel</Data></Cell>
    <Cell ss:MergeDown=""2"" ss:StyleID=""m1536116198852""><Data ss:Type=""String"">Marpol special areas Yes/No</Data></Cell>
    <Cell ss:MergeAcross=""1"" ss:MergeDown=""2"" ss:StyleID=""m1536116198376""><Data
      ss:Type=""String"">Squat effect</Data></Cell>
    <Cell ss:MergeDown=""2"" ss:StyleID=""m1536116198396""><Data ss:Type=""String"">Minimum underkeel clearance  included squat effect</Data></Cell>
    <Cell ss:MergeDown=""2"" ss:StyleID=""m1536116200872""><Data ss:Type=""String"">Safe distance from obstacles</Data></Cell>
    <Cell ss:MergeDown=""2"" ss:StyleID=""m1536116200892""><Data ss:Type=""String"">Ship positioning system (Gps, etc.)</Data></Cell>
    <Cell ss:MergeDown=""2"" ss:StyleID=""m1536116200912""><Data ss:Type=""String"">Current Direction/Speed</Data></Cell>
    <Cell ss:MergeAcross=""1"" ss:MergeDown=""2"" ss:StyleID=""m1536116200932""><Data
      ss:Type=""String"">Stand by in the engine Yes/No</Data></Cell>
    <Cell ss:MergeAcross=""6"" ss:StyleID=""m1536116205864""><ss:Data ss:Type=""String""
      xmlns=""http://www.w3.org/TR/REC-html40""><I>From pilot station to the key or  v.v</I><Font
       html:Size=""9"">.</Font></ss:Data></Cell>
    <Cell ss:MergeDown=""2"" ss:StyleID=""m1536116199416""><Data ss:Type=""String"">Refuge port of emergency anchorage</Data></Cell>
    <Cell ss:MergeDown=""2"" ss:StyleID=""m1536116199436""><Data ss:Type=""String"">Nautical publications to be used</Data></Cell>
    <Cell ss:MergeDown=""2"" ss:StyleID=""m1536116199456""><Data ss:Type=""String"">Notes</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight=""0"" ss:Height=""19.5"">
    <Cell ss:Index=""22"" ss:MergeAcross=""2"" ss:StyleID=""m1536116199476""><Data
      ss:Type=""String"">Tide time</Data></Cell>
    <Cell ss:MergeAcross=""1"" ss:MergeDown=""1"" ss:StyleID=""m1536116199316""><Data
      ss:Type=""String"">Miles pilot/berth (or  v.v.) </Data></Cell>
    <Cell ss:MergeAcross=""1"" ss:MergeDown=""1"" ss:StyleID=""m1536116199336""><Data
      ss:Type=""String"">Miles to the lock </Data></Cell>
   </Row>
   <Row ss:AutoFitHeight=""0"" ss:Height=""105"">
    <Cell ss:Index=""22"" ss:StyleID=""s172""><ss:Data ss:Type=""String""
      xmlns=""http://www.w3.org/TR/REC-html40""><I><Font html:Color=""#000000"">Low</Font></I><Font
       html:Color=""#000000""> </Font></ss:Data></Cell>
    <Cell ss:StyleID=""s172""><Data ss:Type=""String"">High</Data></Cell>
    <Cell ss:StyleID=""s172""><Data ss:Type=""String"">Slack</Data></Cell>
   </Row>";

		public static int TableHeadTotalRows =
			Regex.Matches(TableHeadRows, "/Row", RegexOptions.IgnoreCase).Count;
	}
}
