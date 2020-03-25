/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 *
 * EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
 * See https://github.com/JanKallman/EPPlus for details.
 *
 * Copyright (C) 2011  Jan Källman
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.

 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  
 * See the GNU Lesser General Public License for more details.
 *
 * The GNU Lesser General Public License can be viewed at http://www.opensource.org/licenses/lgpl-license.php
 * If you unfamiliar with this license or have questions about it, here is an http://www.gnu.org/licenses/gpl-faq.html
 *
 * All code and executables are provided "as is" with no warranty either express or implied. 
 * The author accepts no liability for any damage or loss of business that this product may cause.
 *
 * Code change notes:
 * 
 * Author							Change						Date
 * ******************************************************************************
 * Mats Alm   		                Added       		        2013-03-01 (Prior file history on https://github.com/swmal/ExcelFormulaParser)
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace OfficeOpenXml.FormulaParsing.ExcelUtilities
{
    public static class ExcelAddressUtil
    {
        public static readonly char[] SheetNameInvalidChars = new char[] { '?', ':', '*', '/', '\\' };

        public static readonly char[] SpecialReservedChars = new char[] { 'C', 'c', 'R', 'r' };

        public static bool OkAtStart(int x)
        {
            if (x >= 1 && x <= 64) return false;
            if (x == 91) return false;
            if (x >= 93 && x <= 94) return false;
            if (x == 96) return false;
            if (x >= 123 && x <= 160) return false;
            if (x >= 162 && x <= 163) return false;
            if (x >= 165 && x <= 166) return false;
            if (x == 169) return false;
            if (x >= 171 && x <= 172) return false;
            if (x == 174) return false;
            if (x == 187) return false;
            if (x >= 697 && x <= 698) return false;
            if (x >= 706 && x <= 710) return false;
            if (x == 712) return false;
            if (x == 716) return false;
            if (x >= 718 && x <= 719) return false;
            if (x >= 722 && x <= 727) return false;
            if (x == 732) return false;
            if (x >= 734 && x <= 735) return false;
            if (x >= 741 && x <= 749) return false;
            if (x >= 751 && x <= 879) return false;
            if (x >= 884 && x <= 885) return false;
            if (x >= 888 && x <= 889) return false;
            if (x >= 894 && x <= 901) return false;
            if (x == 903) return false;
            if (x == 907) return false;
            if (x == 909) return false;
            if (x == 930) return false;
            if (x == 1014) return false;
            if (x >= 1154 && x <= 1161) return false;
            if (x >= 1316 && x <= 1328) return false;
            if (x >= 1367 && x <= 1368) return false;
            if (x >= 1370 && x <= 1376) return false;
            if (x >= 1416 && x <= 1487) return false;
            if (x >= 1515 && x <= 1519) return false;
            if (x >= 1523 && x <= 1568) return false;
            if (x >= 1611 && x <= 1645) return false;
            if (x == 1648) return false;
            if (x == 1748) return false;
            if (x >= 1750 && x <= 1764) return false;
            if (x >= 1767 && x <= 1773) return false;
            if (x >= 1776 && x <= 1785) return false;
            if (x >= 1789 && x <= 1790) return false;
            if (x >= 1792 && x <= 1807) return false;
            if (x == 1809) return false;
            if (x >= 1840 && x <= 1868) return false;
            if (x >= 1958 && x <= 1968) return false;
            if (x >= 1970 && x <= 1993) return false;
            if (x >= 2027 && x <= 2035) return false;
            if (x >= 2038 && x <= 2041) return false;
            if (x >= 2043 && x <= 2307) return false;
            if (x >= 2362 && x <= 2364) return false;
            if (x >= 2366 && x <= 2383) return false;
            if (x >= 2385 && x <= 2391) return false;
            if (x >= 2402 && x <= 2416) return false;
            if (x >= 2419 && x <= 2426) return false;
            if (x >= 2432 && x <= 2436) return false;
            if (x >= 2445 && x <= 2446) return false;
            if (x >= 2449 && x <= 2450) return false;
            if (x == 2473) return false;
            if (x == 2481) return false;
            if (x >= 2483 && x <= 2485) return false;
            if (x >= 2490 && x <= 2492) return false;
            if (x >= 2494 && x <= 2509) return false;
            if (x >= 2511 && x <= 2523) return false;
            if (x == 2526) return false;
            if (x >= 2530 && x <= 2543) return false;
            if (x >= 2546 && x <= 2564) return false;
            if (x >= 2571 && x <= 2574) return false;
            if (x >= 2577 && x <= 2578) return false;
            if (x == 2601) return false;
            if (x == 2609) return false;
            if (x == 2612) return false;
            if (x == 2615) return false;
            if (x >= 2618 && x <= 2648) return false;
            if (x == 2653) return false;
            if (x >= 2655 && x <= 2673) return false;
            if (x >= 2677 && x <= 2692) return false;
            if (x == 2702) return false;
            if (x == 2706) return false;
            if (x == 2729) return false;
            if (x == 2737) return false;
            if (x == 2740) return false;
            if (x >= 2746 && x <= 2748) return false;
            if (x >= 2750 && x <= 2767) return false;
            if (x >= 2769 && x <= 2783) return false;
            if (x >= 2786 && x <= 2820) return false;
            if (x >= 2829 && x <= 2830) return false;
            if (x >= 2833 && x <= 2834) return false;
            if (x == 2857) return false;
            if (x == 2865) return false;
            if (x == 2868) return false;
            if (x >= 2874 && x <= 2876) return false;
            if (x >= 2878 && x <= 2907) return false;
            if (x == 2910) return false;
            if (x >= 2914 && x <= 2928) return false;
            if (x >= 2930 && x <= 2946) return false;
            if (x == 2948) return false;
            if (x >= 2955 && x <= 2957) return false;
            if (x == 2961) return false;
            if (x >= 2966 && x <= 2968) return false;
            if (x == 2971) return false;
            if (x == 2973) return false;
            if (x >= 2976 && x <= 2978) return false;
            if (x >= 2981 && x <= 2983) return false;
            if (x >= 2987 && x <= 2989) return false;
            if (x >= 3002 && x <= 3023) return false;
            if (x >= 3025 && x <= 3076) return false;
            if (x == 3085) return false;
            if (x == 3089) return false;
            if (x == 3113) return false;
            if (x == 3124) return false;
            if (x >= 3130 && x <= 3132) return false;
            if (x >= 3134 && x <= 3159) return false;
            if (x >= 3162 && x <= 3167) return false;
            if (x >= 3170 && x <= 3204) return false;
            if (x == 3213) return false;
            if (x == 3217) return false;
            if (x == 3241) return false;
            if (x == 3252) return false;
            if (x >= 3258 && x <= 3260) return false;
            if (x >= 3262 && x <= 3293) return false;
            if (x == 3295) return false;
            if (x >= 3298 && x <= 3332) return false;
            if (x == 3341) return false;
            if (x == 3345) return false;
            if (x == 3369) return false;
            if (x >= 3386 && x <= 3388) return false;
            if (x >= 3390 && x <= 3423) return false;
            if (x >= 3426 && x <= 3449) return false;
            if (x >= 3456 && x <= 3460) return false;
            if (x >= 3479 && x <= 3481) return false;
            if (x == 3506) return false;
            if (x == 3516) return false;
            if (x >= 3518 && x <= 3519) return false;
            if (x >= 3527 && x <= 3584) return false;
            if (x >= 3643 && x <= 3647) return false;
            if (x >= 3663 && x <= 3712) return false;
            if (x == 3715) return false;
            if (x >= 3717 && x <= 3718) return false;
            if (x == 3721) return false;
            if (x >= 3723 && x <= 3724) return false;
            if (x >= 3726 && x <= 3731) return false;
            if (x == 3736) return false;
            if (x == 3744) return false;
            if (x == 3748) return false;
            if (x == 3750) return false;
            if (x >= 3752 && x <= 3753) return false;
            if (x == 3756) return false;
            if (x == 3761) return false;
            if (x >= 3764 && x <= 3772) return false;
            if (x >= 3774 && x <= 3775) return false;
            if (x == 3781) return false;
            if (x >= 3783 && x <= 3803) return false;
            if (x >= 3806 && x <= 3839) return false;
            if (x >= 3841 && x <= 3903) return false;
            if (x == 3912) return false;
            if (x >= 3949 && x <= 3975) return false;
            if (x >= 3980 && x <= 4095) return false;
            if (x >= 4139 && x <= 4158) return false;
            if (x >= 4160 && x <= 4175) return false;
            if (x >= 4182 && x <= 4185) return false;
            if (x >= 4190 && x <= 4192) return false;
            if (x >= 4194 && x <= 4196) return false;
            if (x >= 4199 && x <= 4205) return false;
            if (x >= 4209 && x <= 4212) return false;
            if (x >= 4226 && x <= 4237) return false;
            if (x >= 4239 && x <= 4255) return false;
            if (x >= 4294 && x <= 4303) return false;
            if (x == 4347) return false;
            if (x >= 4349 && x <= 4351) return false;
            if (x >= 4442 && x <= 4446) return false;
            if (x >= 4515 && x <= 4519) return false;
            if (x >= 4602 && x <= 4607) return false;
            if (x == 4681) return false;
            if (x >= 4686 && x <= 4687) return false;
            if (x == 4695) return false;
            if (x == 4697) return false;
            if (x >= 4702 && x <= 4703) return false;
            if (x == 4745) return false;
            if (x >= 4750 && x <= 4751) return false;
            if (x == 4785) return false;
            if (x >= 4790 && x <= 4791) return false;
            if (x == 4799) return false;
            if (x == 4801) return false;
            if (x >= 4806 && x <= 4807) return false;
            if (x == 4823) return false;
            if (x == 4881) return false;
            if (x >= 4886 && x <= 4887) return false;
            if (x >= 4955 && x <= 4991) return false;
            if (x >= 5008 && x <= 5023) return false;
            if (x >= 5109 && x <= 5120) return false;
            if (x >= 5741 && x <= 5742) return false;
            if (x >= 5751 && x <= 5760) return false;
            if (x >= 5787 && x <= 5791) return false;
            if (x >= 5867 && x <= 5869) return false;
            if (x >= 5873 && x <= 5887) return false;
            if (x == 5901) return false;
            if (x >= 5906 && x <= 5919) return false;
            if (x >= 5938 && x <= 5951) return false;
            if (x >= 5970 && x <= 5983) return false;
            if (x == 5997) return false;
            if (x >= 6001 && x <= 6015) return false;
            if (x >= 6068 && x <= 6102) return false;
            if (x >= 6104 && x <= 6107) return false;
            if (x >= 6109 && x <= 6175) return false;
            if (x >= 6264 && x <= 6271) return false;
            if (x == 6313) return false;
            if (x >= 6315 && x <= 6399) return false;
            if (x >= 6429 && x <= 6479) return false;
            if (x >= 6510 && x <= 6511) return false;
            if (x >= 6517 && x <= 6527) return false;
            if (x >= 6570 && x <= 6592) return false;
            if (x >= 6600 && x <= 6655) return false;
            if (x >= 6679 && x <= 6916) return false;
            if (x >= 6964 && x <= 6980) return false;
            if (x >= 6988 && x <= 7042) return false;
            if (x >= 7073 && x <= 7085) return false;
            if (x >= 7088 && x <= 7167) return false;
            if (x >= 7204 && x <= 7244) return false;
            if (x >= 7248 && x <= 7257) return false;
            if (x >= 7294 && x <= 7423) return false;
            if (x >= 7616 && x <= 7679) return false;
            if (x >= 7958 && x <= 7959) return false;
            if (x >= 7966 && x <= 7967) return false;
            if (x >= 8006 && x <= 8007) return false;
            if (x >= 8014 && x <= 8015) return false;
            if (x == 8024) return false;
            if (x == 8026) return false;
            if (x == 8028) return false;
            if (x == 8030) return false;
            if (x >= 8062 && x <= 8063) return false;
            if (x == 8117) return false;
            if (x == 8125) return false;
            if (x >= 8127 && x <= 8129) return false;
            if (x == 8133) return false;
            if (x >= 8141 && x <= 8143) return false;
            if (x >= 8148 && x <= 8149) return false;
            if (x >= 8156 && x <= 8159) return false;
            if (x >= 8173 && x <= 8177) return false;
            if (x == 8181) return false;
            if (x >= 8189 && x <= 8207) return false;
            if (x >= 8209 && x <= 8210) return false;
            if (x == 8215) return false;
            if (x >= 8217 && x <= 8219) return false;
            if (x >= 8222 && x <= 8223) return false;
            if (x >= 8226 && x <= 8228) return false;
            if (x >= 8232 && x <= 8239) return false;
            if (x == 8241) return false;
            if (x == 8244) return false;
            if (x >= 8246 && x <= 8250) return false;
            if (x >= 8252 && x <= 8304) return false;
            if (x >= 8306 && x <= 8307) return false;
            if (x >= 8309 && x <= 8318) return false;
            if (x == 8320) return false;
            if (x >= 8325 && x <= 8335) return false;
            if (x >= 8341 && x <= 8449) return false;
            if (x == 8452) return false;
            if (x == 8454) return false;
            if (x == 8456) return false;
            if (x == 8468) return false;
            if (x >= 8471 && x <= 8472) return false;
            if (x >= 8478 && x <= 8480) return false;
            if (x == 8483) return false;
            if (x == 8485) return false;
            if (x == 8487) return false;
            if (x == 8489) return false;
            if (x == 8494) return false;
            if (x >= 8506 && x <= 8507) return false;
            if (x >= 8512 && x <= 8516) return false;
            if (x >= 8522 && x <= 8525) return false;
            if (x >= 8527 && x <= 8530) return false;
            if (x >= 8533 && x <= 8538) return false;
            if (x == 8543) return false;
            if (x >= 8585 && x <= 8591) return false;
            if (x >= 8602 && x <= 8657) return false;
            if (x == 8659) return false;
            if (x >= 8661 && x <= 8703) return false;
            if (x == 8705) return false;
            if (x >= 8708 && x <= 8710) return false;
            if (x >= 8713 && x <= 8714) return false;
            if (x >= 8716 && x <= 8718) return false;
            if (x == 8720) return false;
            if (x >= 8722 && x <= 8724) return false;
            if (x >= 8726 && x <= 8729) return false;
            if (x >= 8731 && x <= 8732) return false;
            if (x >= 8737 && x <= 8738) return false;
            if (x == 8740) return false;
            if (x == 8742) return false;
            if (x == 8749) return false;
            if (x >= 8751 && x <= 8755) return false;
            if (x >= 8760 && x <= 8763) return false;
            if (x >= 8766 && x <= 8775) return false;
            if (x >= 8777 && x <= 8779) return false;
            if (x >= 8781 && x <= 8785) return false;
            if (x >= 8787 && x <= 8799) return false;
            if (x >= 8802 && x <= 8803) return false;
            if (x >= 8808 && x <= 8809) return false;
            if (x >= 8812 && x <= 8813) return false;
            if (x >= 8816 && x <= 8833) return false;
            if (x >= 8836 && x <= 8837) return false;
            if (x >= 8840 && x <= 8852) return false;
            if (x >= 8854 && x <= 8856) return false;
            if (x >= 8858 && x <= 8868) return false;
            if (x >= 8870 && x <= 8894) return false;
            if (x >= 8896 && x <= 8977) return false;
            if (x >= 8979 && x <= 9311) return false;
            if (x >= 9398 && x <= 9423) return false;
            if (x >= 9450 && x <= 9471) return false;
            if (x >= 9548 && x <= 9551) return false;
            if (x >= 9589 && x <= 9600) return false;
            if (x >= 9616 && x <= 9617) return false;
            if (x >= 9622 && x <= 9631) return false;
            if (x == 9634) return false;
            if (x >= 9642 && x <= 9649) return false;
            if (x >= 9652 && x <= 9653) return false;
            if (x >= 9656 && x <= 9659) return false;
            if (x >= 9662 && x <= 9663) return false;
            if (x >= 9666 && x <= 9669) return false;
            if (x >= 9673 && x <= 9674) return false;
            if (x >= 9676 && x <= 9677) return false;
            if (x >= 9682 && x <= 9697) return false;
            if (x >= 9702 && x <= 9710) return false;
            if (x >= 9712 && x <= 9732) return false;
            if (x >= 9735 && x <= 9736) return false;
            if (x >= 9738 && x <= 9741) return false;
            if (x >= 9744 && x <= 9755) return false;
            if (x == 9757) return false;
            if (x >= 9759 && x <= 9791) return false;
            if (x == 9793) return false;
            if (x >= 9795 && x <= 9823) return false;
            if (x == 9826) return false;
            if (x == 9830) return false;
            if (x == 9835) return false;
            if (x == 9838) return false;
            if (x >= 9840 && x <= 11263) return false;
            if (x == 11311) return false;
            if (x == 11359) return false;
            if (x == 11376) return false;
            if (x >= 11390 && x <= 11391) return false;
            if (x >= 11493 && x <= 11519) return false;
            if (x >= 11558 && x <= 11567) return false;
            if (x >= 11622 && x <= 11630) return false;
            if (x >= 11632 && x <= 11647) return false;
            if (x >= 11671 && x <= 11679) return false;
            if (x == 11687) return false;
            if (x == 11695) return false;
            if (x == 11703) return false;
            if (x == 11711) return false;
            if (x == 11719) return false;
            if (x == 11727) return false;
            if (x == 11735) return false;
            if (x >= 11743 && x <= 12287) return false;
            if (x == 12292) return false;
            if (x >= 12312 && x <= 12316) return false;
            if (x == 12320) return false;
            if (x >= 12330 && x <= 12336) return false;
            if (x >= 12342 && x <= 12343) return false;
            if (x >= 12349 && x <= 12352) return false;
            if (x >= 12439 && x <= 12442) return false;
            if (x == 12448) return false;
            if (x >= 12544 && x <= 12548) return false;
            if (x >= 12590 && x <= 12592) return false;
            if (x >= 12687 && x <= 12703) return false;
            if (x >= 12728 && x <= 12783) return false;
            if (x >= 12829 && x <= 12831) return false;
            if (x >= 12842 && x <= 12848) return false;
            if (x >= 12851 && x <= 12856) return false;
            if (x >= 12858 && x <= 12895) return false;
            if (x >= 12924 && x <= 12926) return false;
            if (x >= 12928 && x <= 12962) return false;
            if (x >= 12969 && x <= 13058) return false;
            if (x >= 13060 && x <= 13068) return false;
            if (x >= 13070 && x <= 13075) return false;
            if (x >= 13077 && x <= 13079) return false;
            if (x >= 13081 && x <= 13089) return false;
            if (x >= 13092 && x <= 13093) return false;
            if (x >= 13096 && x <= 13098) return false;
            if (x >= 13100 && x <= 13109) return false;
            if (x >= 13111 && x <= 13114) return false;
            if (x >= 13116 && x <= 13128) return false;
            if (x >= 13131 && x <= 13132) return false;
            if (x >= 13134 && x <= 13136) return false;
            if (x >= 13138 && x <= 13142) return false;
            if (x >= 13144 && x <= 13178) return false;
            if (x == 13183) return false;
            if (x >= 13189 && x <= 13191) return false;
            if (x >= 13259 && x <= 13260) return false;
            if (x == 13268) return false;
            if (x == 13271) return false;
            if (x >= 13273 && x <= 13274) return false;
            if (x >= 13278 && x <= 13311) return false;
            if (x >= 19894 && x <= 19967) return false;
            if (x >= 40900 && x <= 40959) return false;
            if (x >= 42125 && x <= 42239) return false;
            if (x >= 42509 && x <= 42511) return false;
            if (x >= 42528 && x <= 42537) return false;
            if (x >= 42540 && x <= 42559) return false;
            if (x >= 42592 && x <= 42593) return false;
            if (x >= 42607 && x <= 42623) return false;
            if (x >= 42648 && x <= 42785) return false;
            if (x >= 42888 && x <= 42890) return false;
            if (x >= 42893 && x <= 43002) return false;
            if (x == 43010) return false;
            if (x == 43014) return false;
            if (x == 43019) return false;
            if (x >= 43043 && x <= 43071) return false;
            if (x >= 43124 && x <= 43137) return false;
            if (x >= 43188 && x <= 43273) return false;
            if (x >= 43302 && x <= 43311) return false;
            if (x >= 43335 && x <= 43519) return false;
            if (x >= 43561 && x <= 43583) return false;
            if (x == 43587) return false;
            if (x >= 43596 && x <= 44031) return false;
            if (x >= 55204 && x <= 57343) return false;
            if (x >= 63561 && x <= 63743) return false;
            if (x >= 64046 && x <= 64047) return false;
            if (x >= 64107 && x <= 64111) return false;
            if (x >= 64218 && x <= 64255) return false;
            if (x >= 64263 && x <= 64274) return false;
            if (x >= 64280 && x <= 64284) return false;
            if (x == 64286) return false;
            if (x == 64297) return false;
            if (x == 64311) return false;
            if (x == 64317) return false;
            if (x == 64319) return false;
            if (x == 64322) return false;
            if (x == 64325) return false;
            if (x >= 64434 && x <= 64466) return false;
            if (x >= 64830 && x <= 64847) return false;
            if (x >= 64912 && x <= 64913) return false;
            if (x >= 64968 && x <= 65007) return false;
            if (x >= 65020 && x <= 65071) return false;
            if (x == 65074) return false;
            if (x >= 65093 && x <= 65096) return false;
            if (x == 65107) return false;
            if (x == 65112) return false;
            if (x == 65127) return false;
            if (x >= 65132 && x <= 65135) return false;
            if (x == 65141) return false;
            if (x >= 65277 && x <= 65280) return false;
            if (x >= 65375 && x <= 65376) return false;
            if (x >= 65471 && x <= 65473) return false;
            if (x >= 65480 && x <= 65481) return false;
            if (x >= 65488 && x <= 65489) return false;
            if (x >= 65496 && x <= 65497) return false;
            if (x >= 65501 && x <= 65503) return false;
            if (x >= 65511 && x <= 65535) return false;
            return true;

        }

        public static bool OkAfterStart(int x)
        {
            if (x >= 1 && x <= 45) return false;
            if (x == 47) return false;
            if (x >= 58 && x <= 62) return false;
            if (x == 64) return false;
            if (x == 91) return false;
            if (x >= 93 && x <= 94) return false;
            if (x == 96) return false;
            if (x >= 123 && x <= 160) return false;
            if (x >= 162 && x <= 163) return false;
            if (x >= 165 && x <= 166) return false;
            if (x == 169) return false;
            if (x >= 171 && x <= 172) return false;
            if (x == 174) return false;
            if (x == 187) return false;
            if (x >= 888 && x <= 889) return false;
            if (x >= 894 && x <= 899) return false;
            if (x == 903) return false;
            if (x == 907) return false;
            if (x == 909) return false;
            if (x == 930) return false;
            if (x >= 1316 && x <= 1328) return false;
            if (x >= 1367 && x <= 1368) return false;
            if (x >= 1370 && x <= 1376) return false;
            if (x >= 1416 && x <= 1424) return false;
            if (x == 1470) return false;
            if (x == 1472) return false;
            if (x == 1475) return false;
            if (x == 1478) return false;
            if (x >= 1480 && x <= 1487) return false;
            if (x >= 1515 && x <= 1519) return false;
            if (x >= 1523 && x <= 1535) return false;
            if (x >= 1540 && x <= 1541) return false;
            if (x >= 1545 && x <= 1546) return false;
            if (x >= 1548 && x <= 1549) return false;
            if (x >= 1563 && x <= 1566) return false;
            if (x == 1568) return false;
            if (x == 1631) return false;
            if (x >= 1642 && x <= 1645) return false;
            if (x == 1748) return false;
            if (x >= 1792 && x <= 1806) return false;
            if (x >= 1867 && x <= 1868) return false;
            if (x >= 1970 && x <= 1983) return false;
            if (x >= 2039 && x <= 2041) return false;
            if (x >= 2043 && x <= 2304) return false;
            if (x >= 2362 && x <= 2363) return false;
            if (x >= 2382 && x <= 2383) return false;
            if (x >= 2389 && x <= 2391) return false;
            if (x >= 2404 && x <= 2405) return false;
            if (x == 2416) return false;
            if (x >= 2419 && x <= 2426) return false;
            if (x == 2432) return false;
            if (x == 2436) return false;
            if (x >= 2445 && x <= 2446) return false;
            if (x >= 2449 && x <= 2450) return false;
            if (x == 2473) return false;
            if (x == 2481) return false;
            if (x >= 2483 && x <= 2485) return false;
            if (x >= 2490 && x <= 2491) return false;
            if (x >= 2501 && x <= 2502) return false;
            if (x >= 2505 && x <= 2506) return false;
            if (x >= 2511 && x <= 2518) return false;
            if (x >= 2520 && x <= 2523) return false;
            if (x == 2526) return false;
            if (x >= 2532 && x <= 2533) return false;
            if (x >= 2555 && x <= 2560) return false;
            if (x == 2564) return false;
            if (x >= 2571 && x <= 2574) return false;
            if (x >= 2577 && x <= 2578) return false;
            if (x == 2601) return false;
            if (x == 2609) return false;
            if (x == 2612) return false;
            if (x == 2615) return false;
            if (x >= 2618 && x <= 2619) return false;
            if (x == 2621) return false;
            if (x >= 2627 && x <= 2630) return false;
            if (x >= 2633 && x <= 2634) return false;
            if (x >= 2638 && x <= 2640) return false;
            if (x >= 2642 && x <= 2648) return false;
            if (x == 2653) return false;
            if (x >= 2655 && x <= 2661) return false;
            if (x >= 2678 && x <= 2688) return false;
            if (x == 2692) return false;
            if (x == 2702) return false;
            if (x == 2706) return false;
            if (x == 2729) return false;
            if (x == 2737) return false;
            if (x == 2740) return false;
            if (x >= 2746 && x <= 2747) return false;
            if (x == 2758) return false;
            if (x == 2762) return false;
            if (x >= 2766 && x <= 2767) return false;
            if (x >= 2769 && x <= 2783) return false;
            if (x >= 2788 && x <= 2789) return false;
            if (x == 2800) return false;
            if (x >= 2802 && x <= 2816) return false;
            if (x == 2820) return false;
            if (x >= 2829 && x <= 2830) return false;
            if (x >= 2833 && x <= 2834) return false;
            if (x == 2857) return false;
            if (x == 2865) return false;
            if (x == 2868) return false;
            if (x >= 2874 && x <= 2875) return false;
            if (x >= 2885 && x <= 2886) return false;
            if (x >= 2889 && x <= 2890) return false;
            if (x >= 2894 && x <= 2901) return false;
            if (x >= 2904 && x <= 2907) return false;
            if (x == 2910) return false;
            if (x >= 2916 && x <= 2917) return false;
            if (x >= 2930 && x <= 2945) return false;
            if (x == 2948) return false;
            if (x >= 2955 && x <= 2957) return false;
            if (x == 2961) return false;
            if (x >= 2966 && x <= 2968) return false;
            if (x == 2971) return false;
            if (x == 2973) return false;
            if (x >= 2976 && x <= 2978) return false;
            if (x >= 2981 && x <= 2983) return false;
            if (x >= 2987 && x <= 2989) return false;
            if (x >= 3002 && x <= 3005) return false;
            if (x >= 3011 && x <= 3013) return false;
            if (x == 3017) return false;
            if (x >= 3022 && x <= 3023) return false;
            if (x >= 3025 && x <= 3030) return false;
            if (x >= 3032 && x <= 3045) return false;
            if (x >= 3067 && x <= 3072) return false;
            if (x == 3076) return false;
            if (x == 3085) return false;
            if (x == 3089) return false;
            if (x == 3113) return false;
            if (x == 3124) return false;
            if (x >= 3130 && x <= 3132) return false;
            if (x == 3141) return false;
            if (x == 3145) return false;
            if (x >= 3150 && x <= 3156) return false;
            if (x == 3159) return false;
            if (x >= 3162 && x <= 3167) return false;
            if (x >= 3172 && x <= 3173) return false;
            if (x >= 3184 && x <= 3191) return false;
            if (x >= 3200 && x <= 3201) return false;
            if (x == 3204) return false;
            if (x == 3213) return false;
            if (x == 3217) return false;
            if (x == 3241) return false;
            if (x == 3252) return false;
            if (x >= 3258 && x <= 3259) return false;
            if (x == 3269) return false;
            if (x == 3273) return false;
            if (x >= 3278 && x <= 3284) return false;
            if (x >= 3287 && x <= 3293) return false;
            if (x == 3295) return false;
            if (x >= 3300 && x <= 3301) return false;
            if (x == 3312) return false;
            if (x >= 3315 && x <= 3329) return false;
            if (x == 3332) return false;
            if (x == 3341) return false;
            if (x == 3345) return false;
            if (x == 3369) return false;
            if (x >= 3386 && x <= 3388) return false;
            if (x == 3397) return false;
            if (x == 3401) return false;
            if (x >= 3406 && x <= 3414) return false;
            if (x >= 3416 && x <= 3423) return false;
            if (x >= 3428 && x <= 3429) return false;
            if (x >= 3446 && x <= 3448) return false;
            if (x >= 3456 && x <= 3457) return false;
            if (x == 3460) return false;
            if (x >= 3479 && x <= 3481) return false;
            if (x == 3506) return false;
            if (x == 3516) return false;
            if (x >= 3518 && x <= 3519) return false;
            if (x >= 3527 && x <= 3529) return false;
            if (x >= 3531 && x <= 3534) return false;
            if (x == 3541) return false;
            if (x == 3543) return false;
            if (x >= 3552 && x <= 3569) return false;
            if (x >= 3572 && x <= 3584) return false;
            if (x >= 3643 && x <= 3646) return false;
            if (x == 3663) return false;
            if (x >= 3674 && x <= 3712) return false;
            if (x == 3715) return false;
            if (x >= 3717 && x <= 3718) return false;
            if (x == 3721) return false;
            if (x >= 3723 && x <= 3724) return false;
            if (x >= 3726 && x <= 3731) return false;
            if (x == 3736) return false;
            if (x == 3744) return false;
            if (x == 3748) return false;
            if (x == 3750) return false;
            if (x >= 3752 && x <= 3753) return false;
            if (x == 3756) return false;
            if (x == 3770) return false;
            if (x >= 3774 && x <= 3775) return false;
            if (x == 3781) return false;
            if (x == 3783) return false;
            if (x >= 3790 && x <= 3791) return false;
            if (x >= 3802 && x <= 3803) return false;
            if (x >= 3806 && x <= 3839) return false;
            if (x >= 3844 && x <= 3858) return false;
            if (x >= 3898 && x <= 3901) return false;
            if (x == 3912) return false;
            if (x >= 3949 && x <= 3952) return false;
            if (x == 3973) return false;
            if (x >= 3980 && x <= 3983) return false;
            if (x == 3992) return false;
            if (x == 4029) return false;
            if (x == 4045) return false;
            if (x >= 4048 && x <= 4095) return false;
            if (x >= 4170 && x <= 4175) return false;
            if (x >= 4250 && x <= 4253) return false;
            if (x >= 4294 && x <= 4303) return false;
            if (x == 4347) return false;
            if (x >= 4349 && x <= 4351) return false;
            if (x >= 4442 && x <= 4446) return false;
            if (x >= 4515 && x <= 4519) return false;
            if (x >= 4602 && x <= 4607) return false;
            if (x == 4681) return false;
            if (x >= 4686 && x <= 4687) return false;
            if (x == 4695) return false;
            if (x == 4697) return false;
            if (x >= 4702 && x <= 4703) return false;
            if (x == 4745) return false;
            if (x >= 4750 && x <= 4751) return false;
            if (x == 4785) return false;
            if (x >= 4790 && x <= 4791) return false;
            if (x == 4799) return false;
            if (x == 4801) return false;
            if (x >= 4806 && x <= 4807) return false;
            if (x == 4823) return false;
            if (x == 4881) return false;
            if (x >= 4886 && x <= 4887) return false;
            if (x >= 4955 && x <= 4958) return false;
            if (x >= 4961 && x <= 4968) return false;
            if (x >= 4989 && x <= 4991) return false;
            if (x >= 5018 && x <= 5023) return false;
            if (x >= 5109 && x <= 5120) return false;
            if (x >= 5741 && x <= 5742) return false;
            if (x >= 5751 && x <= 5759) return false;
            if (x >= 5787 && x <= 5791) return false;
            if (x >= 5867 && x <= 5869) return false;
            if (x >= 5873 && x <= 5887) return false;
            if (x == 5901) return false;
            if (x >= 5909 && x <= 5919) return false;
            if (x >= 5941 && x <= 5951) return false;
            if (x >= 5972 && x <= 5983) return false;
            if (x == 5997) return false;
            if (x == 6001) return false;
            if (x >= 6004 && x <= 6015) return false;
            if (x >= 6100 && x <= 6102) return false;
            if (x >= 6104 && x <= 6106) return false;
            if (x >= 6110 && x <= 6111) return false;
            if (x >= 6122 && x <= 6127) return false;
            if (x >= 6138 && x <= 6154) return false;
            if (x == 6159) return false;
            if (x >= 6170 && x <= 6175) return false;
            if (x >= 6264 && x <= 6271) return false;
            if (x >= 6315 && x <= 6399) return false;
            if (x >= 6429 && x <= 6431) return false;
            if (x >= 6444 && x <= 6447) return false;
            if (x >= 6460 && x <= 6463) return false;
            if (x >= 6465 && x <= 6469) return false;
            if (x >= 6510 && x <= 6511) return false;
            if (x >= 6517 && x <= 6527) return false;
            if (x >= 6570 && x <= 6575) return false;
            if (x >= 6602 && x <= 6607) return false;
            if (x >= 6618 && x <= 6623) return false;
            if (x >= 6684 && x <= 6911) return false;
            if (x >= 6988 && x <= 6991) return false;
            if (x >= 7002 && x <= 7008) return false;
            if (x >= 7037 && x <= 7039) return false;
            if (x >= 7083 && x <= 7085) return false;
            if (x >= 7098 && x <= 7167) return false;
            if (x >= 7224 && x <= 7231) return false;
            if (x >= 7242 && x <= 7244) return false;
            if (x >= 7294 && x <= 7423) return false;
            if (x >= 7655 && x <= 7677) return false;
            if (x >= 7958 && x <= 7959) return false;
            if (x >= 7966 && x <= 7967) return false;
            if (x >= 8006 && x <= 8007) return false;
            if (x >= 8014 && x <= 8015) return false;
            if (x == 8024) return false;
            if (x == 8026) return false;
            if (x == 8028) return false;
            if (x == 8030) return false;
            if (x >= 8062 && x <= 8063) return false;
            if (x == 8117) return false;
            if (x == 8133) return false;
            if (x >= 8148 && x <= 8149) return false;
            if (x == 8156) return false;
            if (x >= 8176 && x <= 8177) return false;
            if (x == 8181) return false;
            if (x == 8191) return false;
            if (x >= 8209 && x <= 8210) return false;
            if (x == 8215) return false;
            if (x >= 8217 && x <= 8219) return false;
            if (x >= 8222 && x <= 8223) return false;
            if (x >= 8226 && x <= 8228) return false;
            if (x == 8241) return false;
            if (x == 8244) return false;
            if (x >= 8246 && x <= 8250) return false;
            if (x >= 8252 && x <= 8259) return false;
            if (x >= 8261 && x <= 8273) return false;
            if (x >= 8275 && x <= 8286) return false;
            if (x >= 8293 && x <= 8297) return false;
            if (x >= 8306 && x <= 8307) return false;
            if (x >= 8317 && x <= 8318) return false;
            if (x >= 8333 && x <= 8335) return false;
            if (x >= 8341 && x <= 8351) return false;
            if (x >= 8374 && x <= 8399) return false;
            if (x >= 8433 && x <= 8447) return false;
            if (x >= 8528 && x <= 8530) return false;
            if (x >= 8585 && x <= 8591) return false;
            if (x >= 9001 && x <= 9002) return false;
            if (x >= 9192 && x <= 9215) return false;
            if (x >= 9255 && x <= 9279) return false;
            if (x >= 9291 && x <= 9311) return false;
            if (x >= 9886 && x <= 9887) return false;
            if (x >= 9917 && x <= 9919) return false;
            if (x >= 9924 && x <= 9984) return false;
            if (x == 9989) return false;
            if (x >= 9994 && x <= 9995) return false;
            if (x == 10024) return false;
            if (x == 10060) return false;
            if (x == 10062) return false;
            if (x >= 10067 && x <= 10069) return false;
            if (x == 10071) return false;
            if (x >= 10079 && x <= 10080) return false;
            if (x >= 10088 && x <= 10101) return false;
            if (x >= 10133 && x <= 10135) return false;
            if (x == 10160) return false;
            if (x == 10175) return false;
            if (x >= 10181 && x <= 10182) return false;
            if (x == 10187) return false;
            if (x >= 10189 && x <= 10191) return false;
            if (x >= 10214 && x <= 10223) return false;
            if (x >= 10627 && x <= 10648) return false;
            if (x >= 10712 && x <= 10715) return false;
            if (x >= 10748 && x <= 10749) return false;
            if (x >= 11085 && x <= 11087) return false;
            if (x >= 11093 && x <= 11263) return false;
            if (x == 11311) return false;
            if (x == 11359) return false;
            if (x == 11376) return false;
            if (x >= 11390 && x <= 11391) return false;
            if (x >= 11499 && x <= 11516) return false;
            if (x >= 11518 && x <= 11519) return false;
            if (x >= 11558 && x <= 11567) return false;
            if (x >= 11622 && x <= 11630) return false;
            if (x >= 11632 && x <= 11647) return false;
            if (x >= 11671 && x <= 11679) return false;
            if (x == 11687) return false;
            if (x == 11695) return false;
            if (x == 11703) return false;
            if (x == 11711) return false;
            if (x == 11719) return false;
            if (x == 11727) return false;
            if (x == 11735) return false;
            if (x == 11743) return false;
            if (x >= 11776 && x <= 11822) return false;
            if (x >= 11824 && x <= 11903) return false;
            if (x == 11930) return false;
            if (x >= 12020 && x <= 12031) return false;
            if (x >= 12246 && x <= 12271) return false;
            if (x >= 12284 && x <= 12287) return false;
            if (x >= 12312 && x <= 12316) return false;
            if (x == 12336) return false;
            if (x == 12349) return false;
            if (x == 12352) return false;
            if (x >= 12439 && x <= 12440) return false;
            if (x == 12448) return false;
            if (x >= 12544 && x <= 12548) return false;
            if (x >= 12590 && x <= 12592) return false;
            if (x == 12687) return false;
            if (x >= 12728 && x <= 12735) return false;
            if (x >= 12772 && x <= 12783) return false;
            if (x == 12831) return false;
            if (x >= 12868 && x <= 12879) return false;
            if (x == 13055) return false;
            if (x >= 19894 && x <= 19903) return false;
            if (x >= 40900 && x <= 40959) return false;
            if (x >= 42125 && x <= 42127) return false;
            if (x >= 42183 && x <= 42239) return false;
            if (x >= 42509 && x <= 42511) return false;
            if (x >= 42540 && x <= 42559) return false;
            if (x >= 42592 && x <= 42593) return false;
            if (x >= 42611 && x <= 42619) return false;
            if (x == 42622) return false;
            if (x >= 42648 && x <= 42751) return false;
            if (x >= 42893 && x <= 43002) return false;
            if (x >= 43052 && x <= 43071) return false;
            if (x >= 43124 && x <= 43135) return false;
            if (x >= 43205 && x <= 43215) return false;
            if (x >= 43226 && x <= 43263) return false;
            if (x == 43311) return false;
            if (x >= 43348 && x <= 43519) return false;
            if (x >= 43575 && x <= 43583) return false;
            if (x >= 43598 && x <= 43599) return false;
            if (x >= 43610 && x <= 44031) return false;
            if (x >= 55204 && x <= 55295) return false;
            if (x >= 64046 && x <= 64047) return false;
            if (x >= 64107 && x <= 64111) return false;
            if (x >= 64218 && x <= 64255) return false;
            if (x >= 64263 && x <= 64274) return false;
            if (x >= 64280 && x <= 64284) return false;
            if (x == 64311) return false;
            if (x == 64317) return false;
            if (x == 64319) return false;
            if (x == 64322) return false;
            if (x == 64325) return false;
            if (x >= 64434 && x <= 64466) return false;
            if (x >= 64830 && x <= 64847) return false;
            if (x >= 64912 && x <= 64913) return false;
            if (x >= 64968 && x <= 65007) return false;
            if (x >= 65022 && x <= 65023) return false;
            if (x >= 65040 && x <= 65055) return false;
            if (x >= 65063 && x <= 65071) return false;
            if (x == 65074) return false;
            if (x >= 65093 && x <= 65096) return false;
            if (x == 65107) return false;
            if (x == 65112) return false;
            if (x == 65127) return false;
            if (x >= 65132 && x <= 65135) return false;
            if (x == 65141) return false;
            if (x >= 65277 && x <= 65278) return false;
            if (x == 65280) return false;
            if (x >= 65375 && x <= 65376) return false;
            if (x >= 65471 && x <= 65473) return false;
            if (x >= 65480 && x <= 65481) return false;
            if (x >= 65488 && x <= 65489) return false;
            if (x >= 65496 && x <= 65497) return false;
            if (x >= 65501 && x <= 65503) return false;
            if (x == 65511) return false;
            if (x >= 65519 && x <= 65528) return false;
            if (x >= 65534 && x <= 65535) return false;
            return true;

        }

        public static bool IsValidAddress(string token)
        {
            int ix;
            if (token[0] == '\'')
            {
                ix = token.LastIndexOf('\'');
                if (ix > 0 && ix < token.Length - 1 && token[ix + 1] == '!')
                {
                    if (token.IndexOfAny(SheetNameInvalidChars, 1, ix - 1) > 0)
                    {
                        return false;
                    }
                    token = token.Substring(ix + 2);
                }
                else
                {
                    return false;
                }
            }
            else if ((ix = token.IndexOf('!')) > 1)
            {
                if (token.IndexOfAny(SheetNameInvalidChars, 0, token.IndexOf('!')) > 0)
                {
                    return false;
                }
                token = token.Substring(token.IndexOf('!') + 1);
            }
            return OfficeOpenXml.ExcelAddress.IsValidAddress(token);
        }
        
        public static bool IsValidName(string name)
        {


            if (name.Length == 0)
            {
                //Empty or Null
                return true;
            }

            if (name.Length > 255)
            {
                //Invalid Length
                return false;
            }

            //Special case '\' followed by 1 character (ok) more than 1 not ok
            if (name[0] == (char)92)
            {
                if (name.Length > 2)
                {
                    //Invalid more than 1 character following '\'
                    return false;
                }
                else
                {
                    if (name.Length > 1)
                    {
                        if (OkAfterStart(Convert.ToInt32(name[1])) == false)
                        {
                            //Invalid Character after start
                            return false;
                        }
                    }
                    else
                    {
                        //Switch length cannot just be '\'
                        return false;
                    }

                }

            }

            //special check 'C','c','R','r' are reserved words 
            if (name.Length == 1 && name.IndexOfAny(SpecialReservedChars) > -1)
            {
                //Invalid Special Reserved Character used
                return false;
            }

            if (OkAtStart(Convert.ToInt32(name[0])) == false)
            {

                //Invalid Start Character
                return false;
            }

            for (int letter = 1; letter < name.Length; letter++)
            {
                if (OkAfterStart(Convert.ToInt32(name[letter])) == false)
                {
                    //Invalid Character after start
                    return false;
                }
            }

            //Logic is reversed here
            if (ExcelCellBase.IsValidAddress(name)==true)
            {
                return false;
            }

            //TODO:Add check for functionnames.
            return true;
        }
        public static string GetValidName(string name)
        {

            if (string.IsNullOrEmpty(name))
            {
                return name;
            }

            string changedName;
            StringBuilder changedNameTmp = new StringBuilder();


            //truncate invalid length            
            if (name.Length > 255)
            {
                string substring;
                substring = name.Substring(0, 255);
                changedNameTmp = changedNameTmp.Insert(0, substring);
            }
            else
            {
                changedNameTmp = changedNameTmp.Insert(0, name);
            }

            changedName = changedNameTmp.ToString();
            changedNameTmp.Length = 0;
            changedNameTmp.Capacity = 0;

            //replace invalid characters at start
            if (OkAtStart(Convert.ToInt32(changedName[0])) == false)
            {
                changedNameTmp.Append('_');
                for (int letter = 1; letter < changedName.Length; letter++)
                {
                    changedNameTmp.Append(changedName[letter]);
                }

                changedName = changedNameTmp.ToString();
                changedNameTmp.Length = 0;
                changedNameTmp.Capacity = 0;
            }

            //replace invalid characters after start
            for (int letter = 0; letter < changedName.Length; letter++)
            {
                if (letter > 0 && OkAfterStart(Convert.ToInt32(changedName[letter])) == false)
                {
                    changedNameTmp.Append('_');
                }
                else
                {
                    changedNameTmp.Append(changedName[letter]);
                }
            }

            changedName = changedNameTmp.ToString();
            changedNameTmp.Length = 0;
            changedNameTmp.Capacity = 0;

            //special check 'C','c','R','r' are reserved words 
            if (name.Length == 1 && name.IndexOfAny(SpecialReservedChars) > -1)
            {
                changedName = "_";
            }

            //Special case '\' followed by 1 character (ok) more than 1 not ok
            if (name[0] == (char)92)
            {
                if (name.Length > 2)
                {
                    //truncate the characters overlimit
                    string substring;
                    substring = name.Substring(0, 2);
                    changedNameTmp = changedNameTmp.Insert(0, substring);
                }
                else
                {
                    if (name.Length > 1)
                    {
                        if (OkAfterStart(Convert.ToInt32(name[1])) == false)
                        {
                            //Replace the invalid character with a valid character
                            changedNameTmp.Append((char)92);
                            changedNameTmp.Append('_');
                        }
                    }
                    else
                    {
                        changedNameTmp.Append('_');
                    }

                }

                changedName = changedNameTmp.ToString();
                changedNameTmp.Length = 0;
                changedNameTmp.Capacity = 0;

            }


            if (ExcelCellBase.IsValidAddress(changedName)==false)
            {

                //replace invalid characters after start
                for (int letter = 0; letter < changedName.Length; letter++)
                {
                    if (letter == 0)
                    {
                        changedNameTmp.Append('_');
                    }
                    else
                    {
                        changedNameTmp.Append(changedName[letter]);
                    }
                }

                changedName = changedNameTmp.ToString();
                
            }


            return changedName;
        }


    }
}
