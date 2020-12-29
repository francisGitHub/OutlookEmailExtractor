namespace OutlookEmailExtractor.Model
{
    public class IndexFormat
    {
        private const string Constant = "C";
        private const char PaddingChar = '0';

        public int MajorNumberPadding { get; set; }
        public int MinorNumberPadding { get; set; }

        public IndexFormat(
            int majorNumberPadding, 
            int minorNumberPadding)
        {
            MajorNumberPadding = majorNumberPadding;
            MinorNumberPadding = minorNumberPadding;
        }

        public string GetIndexNumber(int majorNumber, int minorNumber)
        {
            var majorNumberString = majorNumber.ToString();
            var minorNumberString = minorNumber.ToString();

            return $"{Constant}{majorNumberString.PadLeft(MajorNumberPadding, PaddingChar)}.{minorNumberString.PadLeft(MinorNumberPadding, PaddingChar)}";
        }

    }
}
