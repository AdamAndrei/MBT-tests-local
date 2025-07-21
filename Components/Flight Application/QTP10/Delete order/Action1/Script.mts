Window("Flight Reservation Dialog").Activate
Window("Flight Reservation Dialog").WinButton("Order Information.Delete Order Button").Click
Window("Flight Reservation Dialog").Dialog("Delete Flight Reservations").Activate
Window("Flight Reservation Dialog").Dialog("Delete Flight Reservations").WinButton("OK").Click
AddToTestResults MicPass, "Order was deleted without problem", "BC run finished"
