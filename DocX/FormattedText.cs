using System;

namespace Novacode
{
    public class FormattedText: IComparable
    {
        public int index;
        public string text;
        public Formatting formatting;
        public string containingHyperlinkId;
        public string footnoteId;

        public int CompareTo(object obj)
        {
            FormattedText other = (FormattedText)obj;
            FormattedText tf = this;

            if (other.formatting == null || tf.formatting == null)
                return -1;

            return tf.formatting.CompareTo(other.formatting);   
        }
    }
}
