#region TODO
//TODO: Add the "DataBar" extended options
//TODO: Add tests for all the rules
//TODO: Add the IconSet options
//TODO: Add all the "extList" options
#endregion

#region §18.3.1.18 conditionalFormatting (Conditional Formatting)
//Childs:
//cfRule          (Conditional Formatting Rule) §18.3.1.10
//extLst          (Future Feature Data Storage Area) §18.2.10

//Attributes:
//pivot
//sqref           ST_Sqref simple type (§18.18.76)
#endregion

#region §18.3.1.10 cfRule (Conditional Formatting Rule)
//Childs:
//colorScale      (Color Scale) §18.3.1.16
//dataBar         (Data Bar) §18.3.1.28
//extLst          (Future Feature Data Storage Area) §18.2.10
//formula         (Formula) §18.3.1.43
//iconSet         (Icon Set) §18.3.1.49

//Attributes:
//-----------
//priority        (Priority) The priority of this conditional formatting rule. This value is used to determine which
//                format should be evaluated and rendered. Lower numeric values are higher priority than
//                higher numeric values, where 1 is the highest priority.
//stopIfTrue      (Stop If True) If this flag is 1, no rules with lower priority shall be applied over this rule, when this rule
//                evaluates to true.
//type            (Type) Type of conditional formatting rule. ST_CfType §18.18.12.
//aboveAverage    Indicates whether the rule is an "above average" rule. 1 indicates 'above average'.
//                This attribute is ignored if type is not equal to aboveAverage.
//equalAverage    (Equal Average)
//                Flag indicating whether the 'aboveAverage' and 'belowAverage' criteria is inclusive of the
//                average itself, or exclusive of that value. 1 indicates to include the average value in the
//                criteria. This attribute is ignored if type is not equal to aboveAverage.
//bottom          (Bottom N) Indicates whether a "top/bottom n" rule is a "bottom n" rule. 1 indicates 'bottom'.
//                This attribute is ignored if type is not equal to top10.
//dxfId           (Differential Formatting Id)
//                This is an index to a dxf element in the Styles Part indicating which cell formatting to
//                apply when the conditional formatting rule criteria is met. ST_DxfId simple type (§18.18.25).
//operator        (Operator) The operator in a "cell value is" conditional formatting rule. This attribute is ignored if
//                type is not equal to cellIs. The possible values ST_ConditionalFormattingOperator simple type (§18.18.15).
//percent         (Top 10 Percent)
//                Indicates whether a "top/bottom n" rule is a "top/bottom n percent" rule. This attribute
//                is ignored if type is not equal to top10.
//rank            (Rank) The value of "n" in a "top/bottom n" conditional formatting rule. This attribute is ignored
//                if type is not equal to top10.
//stdDev          (StdDev) The number of standard deviations to include above or below the average in the
//                conditional formatting rule. This attribute is ignored if type is not equal to aboveAverage.
//                If a value is present for stdDev and the rule type = aboveAverage, then this rule is automatically an
//                "above or below N standard deviations" rule.
//text            (Text) The text value in a "text contains" conditional formatting rule. This attribute is ignored if
//                type is not equal to containsText.
//timePeriod      (Time Period) The applicable time period in a "date occurring…" conditional formatting rule. This
//                attribute is ignored if type is not equal to timePeriod. ST_TimePeriod §18.18.82.
#endregion

#region Conditional Formatting XML examples
// All the examples are assumed to be inside <conditionalFormatting sqref="A1:A10">

#region Example "beginsWith"
//<x:cfRule type="beginsWith" dxfId="6" priority="5" operator="beginsWith" text="a">
//  <x:formula>LEFT(A1,LEN("a"))="a"</x:formula>
//</x:cfRule>

//<x:cfRule type="beginsWith" dxfId="5" priority="14" operator="beginsWith" text="&quot;&lt;&gt;">
//  <x:formula>LEFT(A3,LEN("""&lt;&gt;"))="""&lt;&gt;"</x:formula>
//</x:cfRule>
#endregion

#region Example "between"
//<x:cfRule type="cellIs" dxfId="8" priority="10" operator="between">
//  <x:formula>3</x:formula>
//  <x:formula>7</x:formula>
//</x:cfRule>
#endregion

#region Example "containsText"
//<x:cfRule type="containsText" dxfId="5" priority="4" operator="containsText" text="c">
//  <x:formula>NOT(ISERROR(SEARCH("c",A1)))</x:formula>
//</x:cfRule>
#endregion

#region Example "endsWith"
//<x:cfRule type="endsWith" dxfId="9" priority="11" operator="endsWith" text="c">
//  <x:formula>RIGHT(A1,LEN("c"))="c"</x:formula>
//</x:cfRule>
#endregion

#region Example "equal"
//<x:cfRule type="cellIs" dxfId="7" priority="8" operator="equal">
//  <x:formula>"ab"</x:formula>
//</x:cfRule>
#endregion

#region Example "greaterThan"
//<x:cfRule type="cellIs" dxfId="6" priority="7" operator="greaterThan">
//  <x:formula>4</x:formula>
//</x:cfRule>
#endregion

#region Example "greaterThanOrEqual"
//<x:cfRule type="cellIs" dxfId="3" priority="4" operator="greaterThanOrEqual">
//  <x:formula>4</x:formula>
//</x:cfRule>
#endregion

#region Example "lessThan"
//<x:cfRule type="cellIs" dxfId="5" priority="6" operator="lessThan">
//  <x:formula>4</x:formula>
//</x:cfRule>
#endregion

#region Example "lessThanOrEqual"
//<x:cfRule type="cellIs" dxfId="4" priority="5" operator="lessThanOrEqual">
//  <x:formula>4</x:formula>
//</x:cfRule>
#endregion

#region Example "notBetween"
//<x:cfRule type="cellIs" dxfId="2" priority="3" operator="notBetween">
//  <x:formula>3</x:formula>
//  <x:formula>7</x:formula>
//</x:cfRule>
#endregion

#region Example "notContainsText"
//<x:cfRule type="notContainsText" dxfId="4" priority="3" operator="notContains" text="c">
//  <x:formula>ISERROR(SEARCH("c",A1))</x:formula>
//</x:cfRule>
#endregion

#region Example "notEqual"
//<x:cfRule type="cellIs" dxfId="1" priority="2" operator="notEqual">
//  <x:formula>"ab"</x:formula>
//</x:cfRule>
#endregion

#region Example "containsBlanks"
//<x:cfRule type="containsBlanks" dxfId="20" priority="37">
//  <x:formula>LEN(TRIM(A1))=0</x:formula>
//</x:cfRule>
#endregion

#region Example "containsErrors"
//<x:cfRule type="containsErrors" dxfId="15" priority="19">
//  <x:formula>ISERROR(A1)</x:formula>
//</x:cfRule>
#endregion

#region Example "expression"
//<x:cfRule type="expression" dxfId="0" priority="1">
//  <x:formula>RIGHT(J16,1)="b"</x:formula>
//</x:cfRule>
#endregion

#region Example "duplicateValues"
//<x:cfRule type="duplicateValues" dxfId="14" priority="16" />
#endregion

#region Example "notContainsBlanks"
//<x:cfRule type="notContainsBlanks" dxfId="12" priority="14">
//  <x:formula>LEN(TRIM(A1))&gt;0</x:formula>
//</x:cfRule>
#endregion

#region Example "notContainsErrors"
//<x:cfRule type="notContainsErrors" dxfId="11" priority="36">
//  <x:formula>NOT(ISERROR(A1))</x:formula>
//</x:cfRule>
#endregion

#region Example "uniqueValues"
//<x:cfRule type="uniqueValues" dxfId="13" priority="15" />
#endregion

#region Example "last7Days"
//<x:cfRule type="timePeriod" dxfId="39" priority="10" timePeriod="last7Days">
//  <x:formula>AND(TODAY()-FLOOR(A1,1)&lt;=6,FLOOR(A1,1)&lt;=TODAY())</x:formula>
//</x:cfRule>
#endregion

#region Example "lastMonth"
//<x:cfRule type="timePeriod" dxfId="38" priority="9" timePeriod="lastMonth">
//  <x:formula>AND(MONTH(A1)=MONTH(EDATE(TODAY(),0-1)),YEAR(A1)=YEAR(EDATE(TODAY(),0-1)))</x:formula>
//</x:cfRule>
#endregion

#region Example "lastWeek"
//<x:cfRule type="timePeriod" dxfId="37" priority="8" timePeriod="lastWeek">
//  <x:formula>AND(TODAY()-ROUNDDOWN(A1,0)&gt;=(WEEKDAY(TODAY())),TODAY()-ROUNDDOWN(A1,0)&lt;(WEEKDAY(TODAY())+7))</x:formula>
//</x:cfRule>
#endregion

#region Example "nextMonth"
//<x:cfRule type="timePeriod" dxfId="36" priority="7" timePeriod="nextMonth">
//  <x:formula>AND(MONTH(A1)=MONTH(EDATE(TODAY(),0+1)),YEAR(A1)=YEAR(EDATE(TODAY(),0+1)))</x:formula>
//</x:cfRule>
#endregion

#region Example "nextWeek"
//<x:cfRule type="timePeriod" dxfId="35" priority="6" timePeriod="nextWeek">
//  <x:formula>AND(ROUNDDOWN(A1,0)-TODAY()&gt;(7-WEEKDAY(TODAY())),ROUNDDOWN(A1,0)-TODAY()&lt;(15-WEEKDAY(TODAY())))</x:formula>
//</x:cfRule>
#endregion

#region Example "thisMonth"
//<x:cfRule type="timePeriod" dxfId="34" priority="5" timePeriod="thisMonth">
//  <x:formula>AND(MONTH(A1)=MONTH(TODAY()),YEAR(A1)=YEAR(TODAY()))</x:formula>
//</x:cfRule>
#endregion

#region Example "thisWeek"
//<x:cfRule type="timePeriod" dxfId="33" priority="4" timePeriod="thisWeek">
//  <x:formula>AND(TODAY()-ROUNDDOWN(A1,0)&lt;=WEEKDAY(TODAY())-1,ROUNDDOWN(A1,0)-TODAY()&lt;=7-WEEKDAY(TODAY()))</x:formula>
//</x:cfRule>
#endregion

#region Example "today"
//<x:cfRule type="timePeriod" dxfId="32" priority="3" timePeriod="today">
//  <x:formula>FLOOR(A1,1)=TODAY()</x:formula>
//</x:cfRule>
#endregion

#region Example "tomorrow"
//<x:cfRule type="timePeriod" dxfId="31" priority="2" timePeriod="tomorrow">
//  <x:formula>FLOOR(A1,1)=TODAY()+1</x:formula>
//</x:cfRule>
#endregion

#region Example "yesterday"
//<x:cfRule type="timePeriod" dxfId="1" priority="1" timePeriod="yesterday">
//  <x:formula>FLOOR(A1,1)=TODAY()-1</x:formula>
//</x:cfRule>
#endregion

#region Example "twoColorScale"
//<cfRule type="colorScale" priority="1">
//  <colorScale>
//    <cfvo type="min"/>
//    <cfvo type="max"/>
//    <color rgb="FFF8696B"/>
//    <color rgb="FF63BE7B"/>
//  </colorScale>
//</cfRule>
#endregion

#region Examples "iconSet3" (x all the 3 IconSet options)
//<x:cfRule type="iconSet" priority="30">
//  <x:iconSet>
//    <x:cfvo type="percent" val="0" />
//    <x:cfvo type="percent" val="33" />
//    <x:cfvo type="percent" val="67" />
//  </x:iconSet>
//</x:cfRule>

//<x:cfRule type="iconSet" priority="38">
//  <x:iconSet iconSet="3Arrows">
//    <x:cfvo type="percent" val="0" />
//    <x:cfvo type="percent" val="33" />
//    <x:cfvo type="percent" val="67" />
//  </x:iconSet>
//</x:cfRule>
#endregion

#region Examples "iconSet4" (x all the 4 IconSet options)
//<x:cfRule type="iconSet" priority="34">
//  <x:iconSet iconSet="4ArrowsGray">
//    <x:cfvo type="percent" val="0" />
//    <x:cfvo type="percent" val="25" />
//    <x:cfvo type="percent" val="50" />
//    <x:cfvo type="percent" val="75" />
//  </x:iconSet>
//</x:cfRule>
#endregion

#region Examples "iconSet5" (x all the 5 IconSet options)
//<x:cfRule type="iconSet" priority="32">
//  <x:iconSet iconSet="5ArrowsGray">
//    <x:cfvo type="percent" val="0" />
//    <x:cfvo type="percent" val="20" />
//    <x:cfvo type="percent" val="40" />
//    <x:cfvo type="percent" val="60" />
//    <x:cfvo type="percent" val="80" />
//  </x:iconSet>
//</x:cfRule>
#endregion

#region Examples "iconSet" Extended (not implemented yet)
//<x:extLst>
//  <x:ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}">
//    <x14:conditionalFormattings>
//      <x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
//        <x14:cfRule type="iconSet" priority="35" id="{F5114369-080A-47E6-B7EE-499137A3C896}">
//          <x14:iconSet iconSet="3Triangles">
//            <x14:cfvo type="percent">
//              <xm:f>0</xm:f>
//            </x14:cfvo>
//            <x14:cfvo type="percent">
//              <xm:f>33</xm:f>
//            </x14:cfvo>
//            <x14:cfvo type="percent">
//              <xm:f>67</xm:f>
//            </x14:cfvo>
//          </x14:iconSet>
//        </x14:cfRule>
//        <xm:sqref>C3:C12</xm:sqref>
//      </x14:conditionalFormatting>
//      <x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
//        <x14:cfRule type="iconSet" priority="6" id="{0A327384-BF2F-4BF5-9767-123CD690A536}">
//          <x14:iconSet iconSet="3Stars">
//            <x14:cfvo type="percent">
//              <xm:f>0</xm:f>
//            </x14:cfvo>
//            <x14:cfvo type="percent">
//              <xm:f>33</xm:f>
//            </x14:cfvo>
//            <x14:cfvo type="percent">
//              <xm:f>67</xm:f>
//            </x14:cfvo>
//          </x14:iconSet>
//        </x14:cfRule>
//        <xm:sqref>A16:A25</xm:sqref>
//      </x14:conditionalFormatting>
//      <x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
//        <x14:cfRule type="iconSet" priority="19" id="{0DDCA3E4-3536-44B3-A663-4877587295B8}">
//          <x14:iconSet iconSet="3Triangles">
//            <x14:cfvo type="percent">
//              <xm:f>0</xm:f>
//            </x14:cfvo>
//            <x14:cfvo type="percent">
//              <xm:f>33</xm:f>
//            </x14:cfvo>
//            <x14:cfvo type="percent">
//              <xm:f>67</xm:f>
//            </x14:cfvo>
//          </x14:iconSet>
//        </x14:cfRule>
//        <xm:sqref>C16:C25</xm:sqref>
//      </x14:conditionalFormatting>
//      <x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
//        <x14:cfRule type="iconSet" priority="2" id="{E4EDD7FB-880C-408F-B87C-C8DA446AEB78}">
//          <x14:iconSet iconSet="5Boxes">
//            <x14:cfvo type="percent">
//              <xm:f>0</xm:f>
//            </x14:cfvo>
//            <x14:cfvo type="percent">
//              <xm:f>20</xm:f>
//            </x14:cfvo>
//            <x14:cfvo type="percent">
//              <xm:f>40</xm:f>
//            </x14:cfvo>
//            <x14:cfvo type="percent">
//              <xm:f>60</xm:f>
//            </x14:cfvo>
//            <x14:cfvo type="percent">
//              <xm:f>80</xm:f>
//            </x14:cfvo>
//          </x14:iconSet>
//        </x14:cfRule>
//        <xm:sqref>E16:E25</xm:sqref>
//      </x14:conditionalFormatting>
//      <x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
//        <x14:cfRule type="iconSet" priority="1" id="{4CC82060-CB0A-4A31-AEF2-D1A587AC1674}">
//          <x14:iconSet iconSet="3Stars" showValue="0" custom="1">
//            <x14:cfvo type="percent">
//              <xm:f>0</xm:f>
//            </x14:cfvo>
//            <x14:cfvo type="formula">
//              <xm:f>$F$17</xm:f>
//            </x14:cfvo>
//            <x14:cfvo type="num">
//              <xm:f>4</xm:f>
//            </x14:cfvo>
//            <x14:cfIcon iconSet="3Triangles" iconId="1" />
//            <x14:cfIcon iconSet="4RedToBlack" iconId="3" />
//            <x14:cfIcon iconSet="3Stars" iconId="2" />
//          </x14:iconSet>
//        </x14:cfRule>
//        <xm:sqref>F16:F25</xm:sqref>
//      </x14:conditionalFormatting>
//    </x14:conditionalFormattings>
//  </x:ext>
//</x:extLst>
#endregion


#endregion