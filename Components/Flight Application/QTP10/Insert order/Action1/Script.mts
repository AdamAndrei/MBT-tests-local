Window("Flight Reservation Dialog").Activate
Window("Flight Reservation Dialog").WinButton("Order Information.Insert Order Button").Click
VerifyNewOrder
Parameter("OrderNumber") = Window("Flight Reservation Dialog").WinEdit("Order No").GetValue
