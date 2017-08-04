/*
* DDDN.Office.DOCX.WRunInfo
* 
* Copyright(C) 2017 Lukasz Jaskiewicz
* Author: Lukasz Jaskiewicz (lukasz@jaskiewicz.de, devdone@outlook.com)
*
* This program is free software; you can redistribute it and/or modify it under the terms of the
* GNU General Public License as published by the Free Software Foundation; version 2 of the License.
*
* This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied
* warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.See the GNU General Public License for more details.
*
* You should have received a copy of the GNU General Public License along with this program; if not, write
* to the Free Software Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301, USA.
*/

namespace DDDN.Office.DOCX
{
    /// <summary>
    /// w:r
    /// </summary>
    public class WRunInfo
    {
        /// <summary>
        /// w:t
        /// </summary>
        public string Text { get; set; }
        /// <summary>
        /// w:color
        /// </summary>
        public string FontColor { get; set; } = null;
        /// <summary>
        /// w:sz
        /// </summary>
        public string FontSize { get; set; } = null;
        /// <summary>
        /// Checks if a run has some property values.
        /// </summary>
        /// <returns></returns>
        public bool HasProperties()
        {
            if (FontColor != null
                || FontSize != null)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

    }
}
