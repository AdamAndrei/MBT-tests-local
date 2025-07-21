Window("Flight Reservation Dialog").Activate
Window("Flight Reservation Dialog").WinButton("Toolbar.Open Order Button").Click
 Parameter("OrderNumberFound") = FillSearchOrderCriteria(Parameter("PassengerName"), Parameter("FlightDate"), Parameter("OrderNumber_input"))
