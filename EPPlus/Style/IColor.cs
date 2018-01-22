using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Style
{
    //<xsd:complexType name = "CT_Color" >
    //    <xsd:attribute name = "auto" type="xsd:boolean" use="optional"/>
    //    <xsd:attribute name = "indexed" type="xsd:unsignedInt" use="optional"/>
    //    <xsd:attribute name = "rgb" type="ST_UnsignedIntHex" use="optional"/>
    //    <xsd:attribute name = "theme" type="xsd:unsignedInt" use="optional"/>
    //    <xsd:attribute name = "tint" type="xsd:double" use="optional" default="0.0"/>
    //</xsd:complexType>

    interface IColor
    {
        //bool? Auto { get; set; }  //TODO: Add this functionallity
        int Indexed { get; set; }
        string Rgb { get; }
        string Theme { get; }
        decimal Tint { get; set; }
        void SetColor(Color color);
    }
}
