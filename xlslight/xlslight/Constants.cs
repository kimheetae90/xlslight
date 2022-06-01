using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace xlslight
{
    public struct Offset
    {
        public int x;
        public int y;

        public Offset(int x, int y)
        {
            this.x = x;
            this.y = y;
        }

        public static Offset Parse(string offsetString)
        {
            Offset result = new Offset();
            char[] delimiterChars = { ',' };
            if (offsetString != null && offsetString.Length != 0)
            {
                string[] offsetSplitedString = offsetString.Split(delimiterChars);
                if (offsetSplitedString.Length > 0)
                {
                    int.TryParse(offsetSplitedString[0], out result.x);
                }

                if (offsetSplitedString.Length > 1)
                {
                    int.TryParse(offsetSplitedString[1], out result.y);
                }
            }

            return result;
        }

        public override string ToString()
        {
            string offsetString = string.Empty;

            if (y > 0)
            {
                offsetString += x.ToString();
                offsetString += ",";
                offsetString += y.ToString();
            }
            else
            {
                if (x > 1)
                {
                    offsetString += x.ToString();
                }
            }

            return offsetString;
        }

        public static Offset operator +(Offset offset1, Offset offset2)
        {
            return new Offset(offset1.x + offset2.x, offset1.y + offset2.y);
        }

        public static Offset operator -(Offset offset1, Offset offset2)
        {
            return new Offset(offset1.x - offset2.x, offset1.y - offset2.y);
        }
    }
}
