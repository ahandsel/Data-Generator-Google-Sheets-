// Data Generator {Google Sheets}
// Scripts Creates a "Custom menu" w/ "Phone Numbers" button

function onOpen() {
    var ui = SpreadsheetApp.getUi();
    // Or DocumentApp or FormApp.
    ui.createMenu('Data Generator')
        .addItem('Phone Numbers', 'Phone_Numbers')
        .addItem('Full Names', 'Full_Names')
        .addItem('City, State, Zip', 'Address_CSZ')
        .addSeparator()
        .addSubMenu(
            ui.createMenu('Digits')
                .addItem('0 to 10', 'Z_to_T_Digits')
                .addItem('2 Digits', 'Two_Digits')
                .addItem('4 Digits', 'Four_Digits')
                .addItem('5 Digits', 'Five_Digits')
        )
        .addSeparator()
        .addSubMenu(
            ui.createMenu('Dates')
                .addItem('Past 3 Months to Today', 'Past_Months')
                .addItem('Today to Future 3 Months', 'Future_Months')
        )
        .addToUi();
}
// = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
// = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
// Dates(x) {Generates random dates either in between Past 3 Months to Today or Today to Future 3 Months}
// Steps of the code:
// 1. Determine the Selected Cells
// 2. Generate random Dates between the Past 3 Months to Today  => store in tempArray
// 3. Insert the Dates  from the tempArray into the Selected Cells
function Past_Months() { Dates(0)   }
function Future_Months() { Dates(1)   }

function Dates(x) {
    var Index_Array = location(Index_Array); //Index_Array = Row 1, Column 1, Row 2, Column 2
    var range = location(range); //getRange(row, column, numRows, numColumns)
    var selectedAddress = location(selectedAddress); //selectedAddress
    var temp_cell_count = cell_count(selectedAddress); // Count of data to generate
    var i;
    var Dates_Array = [];

    // Determine the Following Dates:
    // Today:
    var Date_Today = new Date();
    //Browser.msgBox("Date_Today = "+Date_Today);

    //3 Months in the Past
    var temp_Today = new Date();
    temp_Today.setMonth(temp_Today.getMonth() - 3);
    var Date_Past = temp_Today;
    //Browser.msgBox("Date_Past = "+Date_Past);
    
    //3 Months in the Future
    var temp_Today = new Date();
    temp_Today.setMonth(temp_Today.getMonth() + 3);
    var Date_Future = temp_Today;
    //Browser.msgBox("Date_Future = "+Date_Future);

    if (x == 0) {
        for (i = 0; i < temp_cell_count; i++){Dates_Array[i] = gen_Dates(Date_Past, Date_Today);}
    } else {
        for (i = 0; i < temp_cell_count; i++){Dates_Array[i] = gen_Dates(Date_Today, Date_Future);}        
    }
    insert_to_cells(selectedAddress, Dates_Array);
}
// = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
// gen_Dates(x) {Generates random dates either in between Past 3 Months to Today or Today to Future 3 Months}
// if x = 0 => Past // if x = 1 => Future
function gen_Dates(start, end) {
    var temp_Date = new Date(start.getTime() + Math.random() * (end.getTime() - start.getTime()));
    var temp_Date_String = getFormattedDate(temp_Date);
    return temp_Date_String;
}
function getFormattedDate(date) {
    var year = date.getFullYear();
  
    var month = (1 + date.getMonth()).toString();
    month = month.length > 1 ? month : '0' + month;
  
    var day = date.getDate().toString();
    day = day.length > 1 ? day : '0' + day;
    
    return month + '/' + day + '/' + year;
  }
// = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =// = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
// = = = = = = = = = = = = = = = = = = = = = = = = = = = =
// Address_CSZ {City, State, Zip}
// Steps of the code:
// 1. Determine the Selected Cells
// 2. Generate random numbers
// 3. Select a name from the list
// 4. Insert the Full Name into the Selected Cells

function Address_CSZ() {
    var Index_Array = location(Index_Array); //Index_Array = Row 1, Column 1, Row 2, Column 2
    var range = location(range); //getRange(row, column, numRows, numColumns)
    var selectedAddress = location(selectedAddress); //selectedAddress
    var temp_cell_count = cell_count(selectedAddress); // Count of data to generate
    var i;
    var CSZ_Array = [];
    for (i = 0; i < temp_cell_count; i++) {
        CSZ_Array[i] = gen_Address_CSZ(); 
    }
    insert_to_cells(selectedAddress, CSZ_Array);
}
// = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
// Generate Random Address_CSZ {City, State, Zip} {e.g. San Jose, CA, 11111}
function gen_Address_CSZ() {
    var CSZ_List = ["Apache Junction, Arizona, 62891","Charleston, South Carolina, 26691","Shawnee Mission, Kansas, 94245","Fort Lauderdale, Florida, 11035","Portland, Oregon, 33157","Biloxi, Mississippi, 42595","Sandy, Utah, 51557","Greenville, South Carolina, 29955","Fresno, California, 57611","Boise, Idaho, 16218","Flint, Michigan, 67313","Washington, District of Columbia, 80573","Dallas, Texas, 47686","Kansas City, Kansas, 36544","Arlington, Virginia, 69404","Baltimore, Maryland, 22648","Houston, Texas, 53755","San Antonio, Texas, 44082","Cambridge, Massachusetts, 62329","Sioux City, Iowa, 73951","Brooklyn, New York, 44764","Houston, Texas, 22035","Omaha, Nebraska, 96386","Anchorage, Alaska, 17731","Springfield, Illinois, 55221","Miami, Florida, 94416","Montpelier, Vermont, 56111","New York City, New York, 61815","Evansville, Indiana, 10994","El Paso, Texas, 55618","Springfield, Illinois, 71094","Detroit, Michigan, 78482","Pueblo, Colorado, 37616","Houston, Texas, 53809","Milwaukee, Wisconsin, 41960","Wilkes Barre, Pennsylvania, 45146","Pensacola, Florida, 38545","Camden, New Jersey, 30553","West Palm Beach, Florida, 98008","Philadelphia, Pennsylvania, 89600","Phoenix, Arizona, 41488","Grand Rapids, Michigan, 31227","Amarillo, Texas, 14403","Pasadena, Texas, 32172","San Francisco, California, 32091","Omaha, Nebraska, 25785","Lexington, Kentucky, 79723","Richmond, Virginia, 57648","Portland, Oregon, 83349","Jackson, Mississippi, 83773","Arlington, Texas, 11716","Milwaukee, Wisconsin, 73963","Kansas City, Kansas, 40700","Syracuse, New York, 85385","Detroit, Michigan, 18647","Des Moines, Iowa, 22479","Pittsburgh, Pennsylvania, 51488","Miami, Florida, 63405","Harrisburg, Pennsylvania, 41936","Erie, Pennsylvania, 93170","San Diego, California, 25288","Lake Worth, Florida, 51662","Austin, Texas, 63863","New Haven, Connecticut, 30633","Corpus Christi, Texas, 35180","Columbus, Georgia, 22422","Anaheim, California, 91949","Peoria, Illinois, 68337","Louisville, Kentucky, 50209","Birmingham, Alabama, 23314","Springfield, Missouri, 33107","Bakersfield, California, 10904","Denver, Colorado, 47146","Brooklyn, New York, 57444","Indianapolis, Indiana, 28224","Houston, Texas, 55163","Greenville, South Carolina, 18641","Lansing, Michigan, 50290","Dallas, Texas, 83655","Jacksonville, Florida, 57120","Hollywood, Florida, 42857","Maple Plain, Minnesota, 27679","Bakersfield, California, 57560","Fort Myers, Florida, 96388","New Orleans, Louisiana, 51932","Lansing, Michigan, 42189","Miami, Florida, 88129","Charlottesville, Virginia, 30600","Anchorage, Alaska, 87093","Bakersfield, California, 41444","Brooklyn, New York, 28259","Lexington, Kentucky, 30631","Los Angeles, California, 45288","Minneapolis, Minnesota, 91147","Grand Rapids, Michigan, 30200","High Point, North Carolina, 48831","Greensboro, North Carolina, 60048","Topeka, Kansas, 14086","Louisville, Kentucky, 51230","South Bend, Indiana, 40746"];
    var i = Math.floor(Math.random() * 101); // Math.floor(Math.random() * (max - min + 1)) + min;
    //Browser.msgBox("i = " + i + " Name = " + FullName_List[i]); //Testing Purpose

    var temp_CSZ = CSZ_List[i];
    return temp_CSZ;
}
// = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
// = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
// Z_to_T_Digits
// Steps of the code:
// 1. Determine the Selected Cells
// 2. Generate random Number numbers => store in tempArray
// 3. Insert the number numbers from the tempArray into the Selected Cells

function Z_to_T_Digits() {
    var Index_Array = location(Index_Array); //Index_Array = Row 1, Column 1, Row 2, Column 2
    var range = location(range); //getRange(row, column, numRows, numColumns)
    var selectedAddress = location(selectedAddress); //selectedAddress
    var temp_cell_count = cell_count(selectedAddress); // Count of Digits to generate
    var i;
    var Digits_Array = [];
    for (i = 0; i < temp_cell_count; i++) {
        Digits_Array[i] = Math.floor(Math.random()*11); 
    }
    insert_to_cells(selectedAddress, Digits_Array);
}
// = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
// = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
// = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
// Digits
// Steps of the code:
// 1. Determine the Selected Cells
// 2. Generate random Four_Digits numbers => store in tempArray
// 3. Insert the Four_Digits numbers from the tempArray into the Selected Cells
function Two_Digits()   { Digits(2) }
function Four_Digits()  { Digits(4) }
function Five_Digits()  { Digits(5) }

function Digits(x) {
    var Index_Array = location(Index_Array); //Index_Array = Row 1, Column 1, Row 2, Column 2
    var range = location(range); //getRange(row, column, numRows, numColumns)
    var selectedAddress = location(selectedAddress); //selectedAddress
    var temp_cell_count = cell_count(selectedAddress); // Count of Digits to generate
    var i;
    var Digits_Array = [];
    for (i = 0; i < temp_cell_count; i++) {
        Digits_Array[i] = gen_Digits(x); 
    }
    insert_to_cells(selectedAddress, Digits_Array);
}
// = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
// Generate Random Four Digits Number {e.g. 1234}
function gen_Digits(x) {
    var i;
    var temp_Num = [];
    for (i = 0; i < x; i++) {
        if (i === 0)    { temp_Num[i] = Math.floor(Math.random() * 9) + 1;  }
        else            { temp_Num[i] = Math.floor(Math.random() * 10);     }
    }
    var temp_Digits = temp_Num.join('');
    return temp_Digits;
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
