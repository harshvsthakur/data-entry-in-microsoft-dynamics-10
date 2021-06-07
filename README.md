# DATA ENTRY IN MICROSFT DYNAMICS 10
These scripts have multiple uses - Purchase Order Entries, Goods Received Vouchers Entry etc.. All via a Citrix client

## 1 Citrix login and GP launch
1. Logs into citrix server via web browser
2. downloads client file
3. Open the client file from downloads folder
4. Opens GP through the Client
5. Logs into Microsoft Dynamics GP

## 2 GP PO Entry dist.
1. Brings the Purchase order entry window in GP forward
2. Copies data from excel sheet and makes entries in the Microsft Dynamics GP.

## 3 GP price format
For some reason the ERP only took entries wihtout including decimal points.
i.e. 45.00 had to be entered as 4500

To overcome this, this code snippet removes the decimal of quantities and prices to be entered in the ERP and restores the decimal at the end of PO entry.

## 4 GRV
1. Automatically generates Goods Recieved Vouchers (GRV) against delivery orders/ delivery notes based on PO number.
2. Opens the scanned PO for verification or print outs.
