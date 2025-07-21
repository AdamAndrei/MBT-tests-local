Window("Flight Reservation Dialog").Activate
FillScheduleInformation Parameter("FlightDate"), Parameter("FlyFrom"), Parameter("FlyTo"), Parameter("FlightNo")
FillOrderInformation Parameter("PassengerName"), Parameter("Tickets"), Parameter("Class")
