using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace _6_лаба
{
    [Serializable]
    public class CustomFont
    {
        public String fontFamily;

        public float size;

        public FontStyle fontStyle;

        public Color color;

        public int id;

        public CustomFont()
        {
        }

        public CustomFont(String fontFamily, float size, FontStyle fontStyle)
        {
            this.fontFamily = fontFamily;
            this.size = size;
            this.fontStyle = fontStyle;
        }

        public CustomFont(int id,String fontFamily, float size, FontStyle fontStyle, Color color)
        {
            this.fontFamily = fontFamily;
            this.size = size;
            this.fontStyle = fontStyle;
            this.color = color;
            this.id = id;
        }

        public CustomFont( String fontFamily, float size, FontStyle fontStyle, Color color)
        {
            this.fontFamily = fontFamily;
            this.size = size;
            this.fontStyle = fontStyle;
            this.color = color;
        }


        public bool EqualsCustom(CustomFont font)
        {
            return
                   this.fontFamily == font.fontFamily &&
                   this.size == font.size &&
                   this.fontStyle == font.fontStyle &&
                   this.color.Equals(font.color);
        }
    }
}
