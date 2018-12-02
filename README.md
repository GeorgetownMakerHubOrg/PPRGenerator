# PPRGenerator
Generate a Purchase REquest form from our spreadhsset

How it works:

You can run the exe file(pick the one for your architecture), with arguments for the start row and end row on the google sheet you want to turn into a new PPR.

The script will get those rows from the google Maker Hub order form, then make a new PPR.

You can then edit the PPR with any other information you need to add

Expectations for the Order Form:
-- 

* Currently only accesses the Maker Hub order form, but this could change.
* Columns it accesses:
  - Column B : Item Number, or Order Number. OPTIONAL
  - Column C : Item Description : REQUIRED
  - Column D : Cost : REQUIRED
  - Column E : Quantity : REQUIRED
  - Column F : Total : REQUIRED
  - Column I : Vendor Name : Format [Vendor Name],[Vendor Email],[Vendor Phone],[Vendor Contact Name] : All fields are optional, but need to be the SAME for EVERY ROW you are pulling

