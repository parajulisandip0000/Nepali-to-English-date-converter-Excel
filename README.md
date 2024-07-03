# Nepali-to-English-date-converter-Excel
Open the VBA Editor: <br>
You can open VBA editor by goint to excel's Options, Customize Ribbon, and check on Developer in Main Tabs <br>
Then navigate to Developer Tab and select Visual Basic <br>
Insert a new module: Insert > Module <br>
Add the VBA code in the box<br>
Save, you can ctrl+s <br>

Then go to excel <br>
Assuming your BS date is in column D starting from cell D2, use the formula in the adjacent column (e.g., column E):<br>
enter <b>=BSToAD(D2)</b> in the E column<br>

<b>Example</b>
<table><thead><tr><th>Patientid</th><th>Age</th><th>Sex</th><th>Nepali Date</th><th>AD Date</th></tr></thead><tbody><tr><td>******</td><td>** Y</td><td>****</td><td>4/22/2077</td><td>=BSToAD(D2)</td></tr><tr><td>******</td><td>** Y</td><td>****</td><td>6/13/2076</td><td>=BSToAD(D3)</td></tr><tr><td>******</td><td>** Y</td><td>****</td><td>6/28/2077</td><td>=BSToAD(D4)</td></tr><tr><td>******</td><td>** Y</td><td>****</td><td>5/25/2077</td><td>=BSToAD(D5)</td></tr></tbody></table>