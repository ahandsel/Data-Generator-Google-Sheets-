// Data Generator {Google Sheets}
// Scripts Creates a "Custom menu" w/ "Phone Numbers" button

function onOpen() {
	// Adds the Custom menu to the Active Spreadsheet
	SpreadsheetApp.getUi()
		.createMenu('Data Generator')
            .addItem('Phone Numbers', 'Phone_Numbers')
			.addItem('Full Names', 'Full_Names')
			.addItem('4 Digit Numbers', 'Four_Digits')
			.addSeparator()
			.addToUi();
}
// = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
// Four_Digits
// Steps of the code:
// 1. Determine the Selected Cells
// 2. Generate random Four_Digits numbers => store in tempArray
// 3. Insert the Four_Digits numbers from the tempArray into the Selected Cells

function Four_Digits() {
    var Index_Array = location(Index_Array); //Index_Array = Row 1, Column 1, Row 2, Column 2
    var range = location(range); //getRange(row, column, numRows, numColumns)
    var selectedAddress = location(selectedAddress); //selectedAddress
    var temp_cell_count = cell_count(selectedAddress); // Count of phone_num to generate
    var i;
    var Four_Digits_Array = [];
    for (i = 0; i < temp_cell_count; i++) {
        Four_Digits_Array[i] = gen_Four_Digits(); 
    }
    insert_to_cells(selectedAddress, Four_Digits_Array);
}
// = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
// Generate Random Four Digits Number {e.g. 1234}
function gen_Four_Digits() {
    var i;
    var temp_Num = [];
    for (i = 0; i < 4; i++) {
        if (i === 0)    { temp_Num[i] = Math.floor(Math.random() * 9) + 1;  }
        else            { temp_Num[i] = Math.floor(Math.random() * 10);     }
    }
    var temp_Str = temp_Num.join('');
    return temp_Str;
}
// = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
// = = = = = = = = = = = = = = = = = = = = = = = = = = = =
// Full Name
// Steps of the code:
// 1. Determine the Selected Cells
// 2. Generate random numbers
// 3. Select a name from the list
// 4. Insert the Full Name into the Selected Cells

function Full_Names() {
    var Index_Array = location(Index_Array); //Index_Array = Row 1, Column 1, Row 2, Column 2
    var range = location(range); //getRange(row, column, numRows, numColumns)
    var selectedAddress = location(selectedAddress); //selectedAddress
    var temp_cell_count = cell_count(selectedAddress); // Count of Full Names to generate
    var i;
    var FullNames_Array = [];
    for (i = 0; i < temp_cell_count; i++) {
        FullNames_Array[i] = gen_FullName(); 
    }
    insert_to_cells(selectedAddress, FullNames_Array);
}
// = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
// Generate Random Full Name {e.g. Bob Smith}
function gen_FullName() {
    var FullName_List = ["Warren Peinton","Fletcher Blampey","Kristoffer Folshom","Haily Trusler","Yard Shotboulte","Stefanie Hailston","Jilleen Endecott","Orson Neilson","Mitchel Grabiec","Alli Sexty","Teresa Gallacher","Phylys Colicot","Reid Bertwistle","Filberto Chataignier","Nola Diack","Sheeree Dunkerton","Nixie Wisby","Ettore Marcinkus","Andonis Olyet","Robyn Cozby","Romona Edmonson","Deny Duncanson","Em Bloxsum","Nerte Locke","Mariel Thrift","Bambi Huffy","Janetta Geoghegan","Natale Piggin","Elka Clucas","Serene Giblin","Wilone Solloway","Nikolaus Toynbee","Wrennie Murfett","Adam Geaves","Hobie Reasce","Hale Dear","Remy Mallindine","Kenny Dell'Abbate","Isa Duffyn","Hansiain Greener","Ailis Arnault","Rolf Gooble","Michele Wagen","Drona Vanyutin","Reina Gabby","Barrett Meert","Tisha Blake","Goldarina Dendle","Jorgan Samarth","Albert Morsey","Kendra Felder","Fairleigh Greader","Bria Ornells","Donica Lis","Melita Mordue","Martin Wanderschek","Karlee Bowman","Antons Totterdell","Georg Wilmot","Lulita Biaggiotti","Rock Fanstone","Leelah Walklot","Sarette Marre","Gerhardine Selman","Evvie Tillett","Olwen Likly","Andreana Cavolini","Randal Morcombe","Arabele Corstan","Dulcea Mackrill","Dani Winkless","Ashien Cuxon","Bekki Conyard","Bertie Ewens","Ermentrude Rey","Viviene Fowley","Aldwin Peasegod","Cathrine Crimin","Clerkclaude Swadlin","Cammie Wooffinden","Curtice Giacomuzzi","Cody Cuthill","Kaleena Napper","Belle Ketchen","Jolyn Stanway","Neille Dinan","Jacki Henmarsh","Rhianna Roseman","Isabelita Hengoed","Kalie Gatesman","Ulberto Manginot","Rona Trevon","Vinnie Pensom","Anabella Skep","Jarrad Keward","Farlie Selland","Earl MacNeilly","Livvyy De Filippi","Olympe Tirrey"];
    var i = Math.floor(Math.random() * 101); // Math.floor(Math.random() * (max - min + 1)) + min;
    
    //Browser.msgBox("i = " + i + " Name = " + FullName_List[i]); //Testing Purpose

    var temp_FullName = FullName_List[i];
    return temp_FullName;
}
// = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
// = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
// Phone_Numbers
// Steps of the code:
// 1. Determine the Selected Cells
// 2. Generate random phone numbers => store in tempArray
// 3. Insert the phone numbers from the tempArray into the Selected Cells

function Phone_Numbers() {
    var Index_Array = location(Index_Array); //Index_Array = Row 1, Column 1, Row 2, Column 2
    var range = location(range); //getRange(row, column, numRows, numColumns)
    var selectedAddress = location(selectedAddress); //selectedAddress
    var temp_cell_count = cell_count(selectedAddress); // Count of phone_num to generate
    var i;
    var Phone_Num_Array = [];
    for (i = 0; i < temp_cell_count; i++) {
        Phone_Num_Array[i] = gen_Phone_Num(); 
    }
    insert_to_cells(selectedAddress, Phone_Num_Array);
}
// = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
// Generate Random Phone Number {e.g. 111-111-1111}
function gen_Phone_Num() {
    var i;
    var temp_Num = [];
    for (i = 0; i < 12; i++) {
        if(i === 3 || i === 7){ temp_Num[i] = "-"  }
        else {                  temp_Num[i] = Math.floor(Math.random()*10); }
    }
    var temp_Str = temp_Num.join('');
    return temp_Str;
}
// = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
// = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
function location(str){
	var sheet = SpreadsheetApp.getActiveSheet();
	var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
	
	// Determine the selected cells & returns their location in A1 notation {e.g. B8:D10}
	var selectedAddress = spreadsheet.getActiveSheet().getSelection().getActiveRange().getA1Notation();
	
    // Index_Array = Array with the Selected Address for getRange input format
	// Index_Array = Row 1, Column 1, Row 2, Column 2
    var Index_Array = A1_to_Index(selectedAddress); 

    //getRange(row, column, numRows, numColumns)
    var range = sheet.getRange(Index_Array[0], Index_Array[1], Index_Array[2], Index_Array[3]);
    if (str === "range")                 { return range;     } 
    else if (str === "Index_Array")      { return Index_Array; }
    else if (str === "selectedAddress")  { return selectedAddress; }
    else {   return selectedAddress;    }
}
// Converts A1 Notation to Array with Row 1, Col 1, Row 2, Col 2 Index
// for getRange(row, column, numRows, numColumns)
function A1_to_Index(A1Notation) {
	// Converting A1 Notation to Column & Row Index
	Address_Array = A1Notation.split(":");			//e.g. B8:D10 => B8, D10 as Array
	col_1_Index = Letter_to_Num(Address_Array[0].charAt(0));		//e.g. B => 2
	col_2_Num = Letter_to_Num(Address_Array[1].charAt(0))- col_1_Index + 1;		//e.g. D => 4

	row_1_Index = Address_Array[0].substring(1);	//e.g. 8
	row_2_Num = Address_Array[1].substring(1) - row_1_Index + 1;	//e.g. 10-8 = 2
	
	// Fixing for A:A issues; Column to Column w/out Cell #
	if (row_1_Index === '') { row_1_Index = 1 }
	if (row_2_Num < 1) { row_2_Num = 1 }

	// Index_Array = Row 1, Column 1, Row 2, Column 2
	var Index_Array = [row_1_Index, col_1_Index, row_2_Num, col_2_Num];
	
	return Index_Array;
}
function Letter_to_Num(col_Letter) {
	// Convert the A1 Notation's Column Letter to a Number
	// A = 1, BB = 27, etc.
	var alphabet = ["A","B","C","D","E","F","G","H","I","J","K","L","M", "N","O","P","Q","R","S","T","U","V","W","X","Y","Z","AA","BB","CC","DD","EE","FF","GG","HH","II","JJ","KK","LL","MM", "NN","OO","PP","QQ","RR","SS","TT","UU","VV","WW","XX","YY","ZZ"];
	var col_Index = alphabet.indexOf(col_Letter)+1;	
	return col_Index;
}
// Determines the number of cells in the column selected
// returns the cell count
function cell_count(selectedAddress) {
	// Converting A1 Notation to Column & Row Index
	Address_Array = selectedAddress.split(":"); //e.g. B8:D10 => B8, D10 as Array
	row_1_Index = Address_Array[0].substring(1);	//e.g. 8
	row_2_Num = Address_Array[1].substring(1) - row_1_Index + 1;	//e.g. 10-8 = 2
	
	// Fixing for A:A issues; Column to Column w/out Cell #
	if (row_1_Index === '') { row_1_Index = 1; }
	if (row_2_Num < 1) { row_2_Num = 1; }
	
	return row_2_Num;
}
function insert_to_cells(temp_location, temp_array) {
    //setValues is used to set an array of values into the corresponding cells in the range. 
    //*Note: this method expects a multi-dimensional array
    //          w/ outer array = rows & inner array = columns
    //          e.g. var employees=[["Adam"],["Barb"],["Chris"]];
    var sheet = SpreadsheetApp.getActiveSheet(); //ssActive
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var range = sheet.getRange(temp_location);
    var column = [];
    for (var i=0; i<temp_array.length; i++){ column.push([temp_array[i]]); }
    range.setValues(column);
}
// = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
// = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
