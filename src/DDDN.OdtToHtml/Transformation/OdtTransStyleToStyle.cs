/*
DDDN.OdtToHtml.Transformation.OdtTransStyleToStyle
Copyright(C) 2017-2019 Lukasz Jaskiewicz (lukasz@jaskiewicz.de)
- This program is free software; you can redistribute it and/or modify it under the terms of the
GNU General Public License as published by the Free Software Foundation; version 2 of the License.
- This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied
warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.
- You should have received a copy of the GNU General Public License along with this program; if not, write
to the Free Software Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301, USA.
*/

using System.Collections.Generic;
using System.Linq;

namespace DDDN.OdtToHtml.Transformation
{
    public class OdtTransStyleToStyle
    {
        public enum RelativeToPage
        {
            None,
            Width,
            Height
        }

        public string LocalName;
        public List<string> PropNames;
        public RelativeToPage AsPercentageTo;
        public List<(OdtStyleValues odtValues, OdtStyleValues cssValues)> Values;

        private OdtTransStyleToStyle()
        {
        }

        public OdtTransStyleToStyle(string localName, string[] propNames, RelativeToPage asPercentageTo = RelativeToPage.None)
        {
            LocalName = localName;
            PropNames = propNames.ToList();
            AsPercentageTo = asPercentageTo; 

            Values = new List<(OdtStyleValues odtValues, OdtStyleValues cssValues)>();
        }

        public OdtTransStyleToStyle Add(OdtStyleValues odt, OdtStyleValues css)
        {
            Values.Add((odt, css));
            return this;
        }
    }

    public class OdtStyleValues
    {
        public OdtStyleValues()
        {
        }

        public OdtStyleValues Add(string prop, string value)
        {
            Values.Add((prop, value));
            return this;
        }

        readonly List<(string name, string value)> Values = new List<(string name, string value)>();
    }
}

