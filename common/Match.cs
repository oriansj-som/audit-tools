namespace common
{
    public class Match
    {
        public static bool CaseInsensitive(string? a, string? b)
        {
            if(null == a && null == b) return true;
            if(null == a) return false;
            if(null == b) return false;
            return a.Equals(b, StringComparison.CurrentCultureIgnoreCase);
        }

        public static bool CaseSensitive(string? a, string? b)
        {
            if (null == a && null == b) return true;
            if (null == a) return false;
            if (null == b) return false;
            return a.Equals(b, StringComparison.CurrentCulture);
        }
    }
}
