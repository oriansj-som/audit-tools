namespace common
{
    public static class Require
    {
        public static void True(bool t, string? message)
        {
            if (!t)
            {
                try
                {
                    if (null == message) throw new Exception("unknown error has occurred");
                    else throw new Exception(message);
                }
                catch (Exception ex)
                { 
                    Console.WriteLine(ex.ToString());
                    Environment.Exit(7);
                }
            }
        }

        public static void False(bool t, string? message)
        {
            if (t)
            {
                try
                {
                    if (null == message) throw new Exception("unknown error has occurred");
                    else throw new Exception(message);
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.ToString());
                    Environment.Exit(8);
                }
            }
        }

        public static void True(bool t, string? message, bool die)
        {
            if (!t)
            {
                try
                {
                    if (null == message) throw new Exception("unknown error has occurred");
                    else throw new Exception(message);
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.ToString());
                    if (die) Environment.Exit(7);
                    else throw new Exception(ex.Message);
                }
            }
        }

        public static void False(bool t, string? message, bool die)
        {
            if (t)
            {
                try
                {
                    if (null == message) throw new Exception("unknown error has occurred");
                    else throw new Exception(message);
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.ToString());
                    if(die) Environment.Exit(8);
                    else throw new Exception(ex.Message);
                }
            }
        }
    }
}
