namespace ImportFromExcel
{
    public class ImprotedReservation
    {
        private string clientBrand;
        private string voucher;
        private string passengerName;
        private int totalPassengers;
        private int adults;
        private int childrens;
        private int infants;
        private string pickupTime;
        private string pickupFrom;
        private string hotelArea;
        private string dropTo;
        private string flightNumber;
        private string departureTime;
        private string arriveTime;

        public ImprotedReservation(string client, string voucher, string passengerName, int passengers, int adults, int childrens, int infants,
            string pickupTime, string pickupFrom, string hotelArea, string dropTo, string flightNumber,
            string departureTime, string arriveTime)
        {
            this.ClientBrand = client;
            this.Voucher = voucher;
            this.PassengerName = passengerName;
            this.TotalPassengers = totalPassengers;
            this.Adults = adults;
            this.Childrens = childrens;
            this.Infants = infants;
            this.PickupTime = pickupTime;
            this.PickupFrom = pickupFrom;
            this.HotelArea = hotelArea;
            this.DropTo = dropTo;
            this.FlightNumber = flightNumber;
            this.DepartureTime = departureTime;
            this.ArriveTime = arriveTime;
        }

        public string ArriveTime
        {
            get { return arriveTime; }
            private set { arriveTime = value; }
        }

        public string DepartureTime
        {
            get { return departureTime; }
            private set { departureTime = value; }
        }

        public string FlightNumber
        {
            get { return flightNumber; }
            private set { flightNumber = value; }
        }

        public string DropTo
        {
            get { return dropTo; }
            private set { dropTo = value; }
        }

        public string HotelArea
        {
            get { return hotelArea; }
            private set { hotelArea = value; }
        }

        public string PickupFrom
        {
            get { return pickupFrom; }
            private set { pickupFrom = value; }
        }

        public string PickupTime
        {
            get { return pickupTime; }
            private set { pickupTime = value; }
        }

        public int Infants
        {
            get { return infants; }
            private set { infants = value; }
        }

        public int Childrens
        {
            get { return childrens; }
            private set { childrens = value; }
        }

        public int Adults
        {
            get { return adults; }
            private set { adults = value; }
        }

        public int TotalPassengers
        {
            get { return totalPassengers; }
            private set { totalPassengers = value; }
        }

        public string PassengerName
        {
            get { return passengerName; }
            private set { passengerName = value; }
        }

        public string Voucher
        {
            get { return voucher; }
            private set { voucher = value; }
        }

        public string ClientBrand
        {
            get { return clientBrand; }
            private set { clientBrand = value; }
        }
    }
}