using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;

namespace Highlighter
{

    public partial class Ribbon
    {
        public static Color color;
        public static bool toggled = false;
        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {
            this.Toggle.Click += Toggle_Click;
            this.Color.Click += Color_Click;
        }

        private void Color_Click(object sender, RibbonControlEventArgs e)
        {
            if (this.colorDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                color = this.colorDialog1.Color;
                HighlighterWork.highlightCell();
            }
        }

        private void Toggle_Click(object sender, RibbonControlEventArgs e)
        {
            if (toggled)
            {
                toggled = false;
                HighlighterWork.clearLastCell();
            }
            else
            {
                toggled = true;
                HighlighterWork.highlightCell();
            }
        }
    }
}
