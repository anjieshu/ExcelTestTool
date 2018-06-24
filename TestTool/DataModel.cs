using System;
namespace TestTool
{
    public class DataModel
    {
        public String key;
        public String timestamp;
        public String status;

        public DataModel()
        {
        }

        //public bool isFinishedLater(DataModel dataModel) {
        //    DateTime dateTime1 = DateTime.ParseExact(timestamp, "dd/MM/yyyy HH:mm:ss",
        //                               System.Globalization.CultureInfo.InvariantCulture);

        //    DateTime dateTime2 = DateTime.ParseExact(dataModel.timestamp, "dd/MM/yyyy HH:mm:ss",
        //                               System.Globalization.CultureInfo.InvariantCulture);

        //    if (DateTime.Compare(dateTime1, dateTime2) >= 0) {
        //        return true;
        //    } else {
        //        return false;
        //    }
        //}

        public bool isFinishedLater(DataModel dataModel) {
            double time1 = double.Parse(timestamp);
            double time2 = double.Parse(dataModel.timestamp);

            if (time1 > time2) {
                return true;
            } else {
                return false;
            }
        }

        public String toString() {
            return "{key: " + key + ", status: " + status + ", timestamp: " + timestamp + "}";
        }
    }
}
