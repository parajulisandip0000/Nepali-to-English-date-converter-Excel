# Nepali-to-English-date-converter-Excel
Open the VBA Editor: <br>
You can open VBA editor by goint to excel's Options, Customize Ribbon, and check on Developer in Main Tabs <br>
Then navigate to Developer Tab and select Visual Basic <br>
Insert a new module: Insert > Module <br>
Add the VBA code in the box<br>
Save, you can ctrl+s <br>

Then go to excel <br>
Assuming your BS date is in column D starting from cell D2, use the formula in the adjacent column (e.g., column E):<br>
enter =BSToAD(D2) in the E column<br>


<br><br>
Patientid	 Age	Sex	    Nepali Date	AD Date
*********    ** Y	****	4/22/2077	=BSToAD(D2)
*********	 ** Y	****	6/13/2076	=BSToAD(D3)
*********	 ** Y	****	6/28/2077	=BSToAD(D4)
*********	 ** Y	****	5/25/2077	=BSToAD(D5)