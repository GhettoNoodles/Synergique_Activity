namespace Synergique_Activity_Formatter.Core
{
    public class MonthRerouter
    {
        public int RerouteMonth(string inputMonth)
        {
            var currentMonth = 0;
            switch (inputMonth)
            {
                case "March":
                    currentMonth = 1;
                    break;
                case "April":
                    currentMonth = 2;
                    break;
                case "May":
                    currentMonth = 3;
                    break;
                case "June":
                    currentMonth = 4;
                    break;
                case "July":
                    currentMonth = 5;
                    break;
                case "August":
                    currentMonth = 6;
                    break;
                case "September":
                    currentMonth = 7;
                    break;
                case "October":
                    currentMonth = 8;
                    break;
                case "November":
                    currentMonth = 9;
                    break;
                case "December":
                    currentMonth = 10;
                    break;
                case "January":
                    currentMonth = 11;
                    break;
                case "February":
                    currentMonth = 12;
                    break;
            }

            return currentMonth;
        }
    }
}