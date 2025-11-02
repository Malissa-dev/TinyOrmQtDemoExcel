#include <QApplication>
#include <QDir>
#include <QFile>
#include <QDebug>
#include "MainWindow.h"
#include <orm/schema.hpp>
#include <orm/db.hpp>
#include "Models/User.h"
#include "Models/Product.h"
#include "Models/Order.h"
#include <OpenXLSX.hpp>

using Orm::Schema;
using Orm::DB;
using namespace OpenXLSX;

static auto db = DB::create();

void setupDatabase()
{
    DB::addConnection({
        {"driver", "QSQLITE"},
        {"database", ":memory:"},
    }, "default");

    DB::setDefaultConnection("default");

    Schema::create("users", [](auto &table) {
        table.id();
        table.string("name");
        table.string("email");
        table.timestamps();
    });

    Schema::create("products", [](auto &table) {
        table.id();
        table.string("name");
        table.decimal("price", 8, 2);
        table.timestamps();
    });

    Schema::create("orders", [](auto &table) {
        table.id();
        table.foreignId("user_id").constrained(); 
        table.decimal("total", 10, 2);
        table.timestamps();
    });

    Schema::create("order_product", [](auto &table) {
        table.foreignId("order_id").constrained().cascadeOnDelete();
        table.foreignId("product_id").constrained().cascadeOnDelete();
        
        table.unsignedTinyInteger("quantity").defaultValue(1);
        
        table.decimal("price", 8, 2); 

        table.primary({"order_id", "product_id"});
    });

	auto users = Models::User::all({}); 
    
	//Create excel file
    XLDocument doc{};
	doc.create("../../test/exportDataManual.xlsx", XLForceOverwrite);

    // Seed Users
    auto user1 = Models::User::create({{"name", "Jane Doe"}, {"email", "jane.doe@example.com"}});
    auto user2 = Models::User::create({{"name", "John Smith"}, {"email", "john.smith@example.com"}});
    auto user3 = Models::User::create({{"name", "Alice Brown"}, {"email", "alice.brown@example.com"}});

	// Fill Excel Users
	auto wks = doc.workbook().worksheet(1);
	wks.setName("Users");
    wks.cell("A1") = "id";
    wks.cell("B1") = "name";
	wks.cell("C1") = "email";
 	wks.cell("A2") = user1.id();	
    wks.cell("B2") = user1.getAttribute("name").toString().toStdString();	
    wks.cell("C2") = user1.getAttribute("email").toString().toStdString();
	wks.cell("A3") = user2.id();	
    wks.cell("B3") = user2.getAttribute("name").toString().toStdString();
    wks.cell("C3") = user2.getAttribute("email").toString().toStdString();
	wks.cell("A4") = user3.id();
	wks.cell("B4") = user3.getAttribute("name").toString().toStdString();
    wks.cell("C4") = user3.getAttribute("email").toString().toStdString();
	doc.save();	

	// Add borders to user
	XLCellFormats & cellFormats = doc.styles().cellFormats();
 	XLCellRange cellRange = wks.range("A1:C4");
	XLBorders & borders = doc.styles().borders();
	XLStyleIndex borderFormat = borders.create();	
    XLColor black("ff000000");
    borders[borderFormat].setBottom(XLLineStyle::XLLineStyleThin, black);
    borders[borderFormat].setTop(XLLineStyle::XLLineStyleThin, black);
    borders[borderFormat].setLeft(XLLineStyle::XLLineStyleThin, black);
    borders[borderFormat].setRight(XLLineStyle::XLLineStyleThin, black);
	XLStyleIndex oldFormat = cellFormats.create( cellFormats[ wks.cell("A1").cellFormat() ] );
	XLStyleIndex newFormat = cellFormats.create(cellFormats[oldFormat]);
	cellFormats[newFormat].setBorderIndex(borderFormat);
    cellRange.setFormat(newFormat);
	doc.save();

	// Title row is bold
	XLFonts & fonts = doc.styles().fonts();
	XLStyleIndex fontBold = fonts.create();
	fonts[ fontBold ].setBold();
 	oldFormat = cellFormats.create( cellFormats[ wks.cell("A1").cellFormat() ] );
	XLStyleIndex boldDefault = cellFormats.create(cellFormats[oldFormat]);
	cellFormats[ boldDefault ].setFontIndex( fontBold );
	XLRow title = wks.row(1);
	wks.setRowFormat(title.rowNumber(), boldDefault );
	doc.save();	
	
	// resize
	wks.column("A").setWidth(3.0); 
	wks.column("C").setWidth(25.0); 

    // Seed Products
	std::vector<Models::Product> orders;
   	orders.emplace_back(Models::Product::create({{"name", "Laptop"}, {"price", 1200.00}}));
    orders.emplace_back(Models::Product::create({{"name", "Wireless Mouse"}, {"price", 25.00}}));
    orders.emplace_back(Models::Product::create({{"name", "Mechanical Keyboard"}, {"price", 85.50}}));
    orders.emplace_back(Models::Product::create({{"name", "24\" Monitor"}, {"price", 200.00}}));
    orders.emplace_back(Models::Product::create({{"name", "USB-C Cable"}, {"price", 10.00}}));

	// Fill Excel products
	doc.workbook().addWorksheet("orders");
	auto wks2 = doc.workbook().worksheet("orders");
	wks2.cell("A1") = "id";
    wks2.cell("B1") = "name";
	wks2.cell("C1") = "price";
	uint32_t rowToFill = 2; // start from row 2 (A2, B2, C2).
	for(auto element : orders){
        wks2.cell(XLCellReference(rowToFill, 1)).value() = element.id();    
        wks2.cell(XLCellReference(rowToFill, 2)).value() = element.getAttribute("name").toString().toStdString();   
        wks2.cell(XLCellReference(rowToFill, 3)).value() = element.getAttribute("price").toString().toStdString();  
		rowToFill++;
	}
	doc.save();

    // Seed Orders
    auto order1 = Models::Order::create({{"user_id", user1.id()}, {"total", 1250.00}});
    auto order2 = Models::Order::create({{"user_id", user2.id()}, {"total", 315.50}});

    // Seed pivot table: order_product
    order1.products()->attach(orders.at(0).id(), {{"quantity", 1}, {"price", orders.at(0).price()}});
    order1.products()->attach(orders.at(1).id(), {{"quantity", 2}, {"price", orders.at(1).price()}});

    order2.products()->attach(orders.at(2).id(), {{"quantity", 1}, {"price", orders.at(2).price()}});
    order2.products()->attach(orders.at(3).id(), {{"quantity", 1}, {"price", orders.at(3).price()}});
    order2.products()->attach(orders.at(4).id(), {{"quantity", 3}, {"price", orders.at(4).price()}});

	doc.close();
}

void excelTestCustomIndividual(){
    auto users = Models::User::all({}); 
    
    XLDocument doc{};
	doc.create("../../test/testIndividual.xlsx", XLForceOverwrite);
    auto wks = doc.workbook().worksheet("Sheet1");

    // Fill some
    wks.cell("A1") = "Jane Doe";
    wks.cell("B1") = 42;
    wks.cell("A2") = "Rando 1";
    wks.cell("B1") = 36;
    wks.cell("A3") = "Rando 3";
    wks.cell("B1") = 64;
 
    wks.cell("C1") = "Bold, Red Text";   
    wks.cell("C2") = "highlighted" ;
    wks.cell("C3") = "Has borders" ;

    //------- Change the font size, color, style (bold)
    // Create a new font from the existing font of that cell (A1)
    XLFonts & fonts = doc.styles().fonts();
    XLCellFormats & cellFormats = doc.styles().cellFormats();
	XLStyleIndex cellFormatIndexC1 = wks.cell("C1").cellFormat();
	XLStyleIndex newCellFormatIndex1 = cellFormats.create(cellFormats[cellFormatIndexC1]);
	XLStyleIndex newFontIndex = fonts.create(fonts[cellFormats[cellFormatIndexC1].fontIndex()]);

    //  modify the Font the way we want
    fonts[newFontIndex].setFontName("Arial");
	fonts[newFontIndex].setFontSize(16);
	fonts[newFontIndex].setBold(true);
	XLColor red("00ff0000");
	fonts[newFontIndex].setFontColor(red);

	cellFormats[newCellFormatIndex1].setFontIndex(newFontIndex);
	cellFormats[newCellFormatIndex1].setApplyFont(true);

	wks.cell("C1").setCellFormat(newCellFormatIndex1);

    // Fill style of C2
	XLFills & fills = doc.styles().fills();
    XLStyleIndex cellFormatIndexC2 = wks.cell("C2").cellFormat();                // get index of cell format
	XLStyleIndex newCellFormatIndex = cellFormats.create(cellFormats[cellFormatIndexC2]);
    XLStyleIndex newFillStyle = fills.create(fills[cellFormats[cellFormatIndexC2].fillIndex()]); //copy from existing
	
	XLColor yellow("00ffff00");
	fills[newFillStyle].setPatternType(XLPatternSolid);
	fills[newFillStyle].setColor(yellow);

	cellFormats[newCellFormatIndex].setFillIndex(newFillStyle);
	cellFormats[newCellFormatIndex].setApplyFill(true);

	wks.cell("C2").setCellFormat(newCellFormatIndex);

    // Border C3
    XLBorders & borders = doc.styles().borders();
	XLStyleIndex borderFormat = borders.create();
    XLColor black("ff000000");
    borders[borderFormat].setBottom(XLLineStyle::XLLineStyleThin, black);
    borders[borderFormat].setTop(XLLineStyle::XLLineStyleThin, black);
    borders[borderFormat].setLeft(XLLineStyle::XLLineStyleThin, black);
    borders[borderFormat].setRight(XLLineStyle::XLLineStyleThin, black);
	XLStyleIndex oldFormat = wks.cell("C3").cellFormat();
	XLStyleIndex newFormat = cellFormats.create(cellFormats[oldFormat]);

	cellFormats[newFormat].setBorderIndex(borderFormat);
	wks.cell("C3").setCellFormat(newFormat);

 	doc.saveAs("../../test/testIndividual.xlsx", XLForceOverwrite);
	doc.close();
}

void excelTestCustomRange(){

 	XLDocument doc{};
	doc.create("../../test/testRange.xlsx", XLForceOverwrite);
	XLWorksheet wks = doc.workbook().worksheet(1);

	XLFonts & fonts = doc.styles().fonts();
	XLFills & fills = doc.styles().fills();
	XLBorders & borders = doc.styles().borders();
	XLCellFormats & cellStyleFormats = doc.styles().cellStyleFormats();
	XLCellFormats & cellFormats = doc.styles().cellFormats();
	XLCellStyles & cellStyles = doc.styles().cellStyles();

	//------------------------------------- Font stuff
	// Create format : bold + underlined
	XLStyleIndex fontBold = fonts.create();
	fonts[ fontBold ].setBold();
	fonts[ fontBold ].setUnderline();
	XLStyleIndex boldDefault = cellFormats.create();
	cellFormats[ boldDefault ].setFontIndex( fontBold );

	// Create format : italic
	XLStyleIndex fontItalic = fonts.create();
	fonts[ fontItalic ].setItalic();
	XLStyleIndex italicDefault = cellFormats.create();
	cellFormats[ italicDefault ].setFontIndex( fontItalic );

	// Create format : bigger, red, strikethrough
	XLStyleIndex fontBigStriked = fonts.create();
	XLColor red   ( "ffff0000" );
	fonts[ fontBigStriked ].setFontColor(red);
	fonts[ fontBigStriked ].setStrikethrough();
	fonts[ fontBigStriked ].setFontSize(16);
	XLStyleIndex fontBigStrikedDefault = cellFormats.create();
	cellFormats[ fontBigStrikedDefault ].setFontIndex( fontBigStriked );
	
	// Set the new formats rows and colums at once
	// Using 2 diff methods for rows
	XLRow row = wks.row(2);
	wks.setRowFormat( row.rowNumber(), boldDefault );
	wks.setRowFormat( XLCellReference::rowAsNumber("1"), boldDefault ); 
	wks.setColumnFormat( XLCellReference::columnAsNumber("A"), italicDefault ); 

	// Set the new format on specific ranges
	// Get range method 1
	XLCellRange cellRange = wks.range("A3", "D10");  
	cellRange = "1"; 
	cellRange = wks.range("A3", "D7");  
	cellRange.setFormat( boldDefault );
	// Get range method 2
	cellRange = wks.range("A7:X25"); // range method 2
	cellRange = "2"; 
	cellRange.setFormat( fontItalic ); 
	// Get range method 3
	cellRange = wks.range("A5:Z5"); 
	cellRange = "3"; 
	cellRange.setFormat( fontItalic ); 
	// use a font on one cell
	wks.cell("A4") = "strike";
	wks.cell("A4").setCellFormat(fontBigStriked);

	//------------------------------------- resize stuff
	// Put something too big in the cells
	cellRange = wks.range("H1:H100");
	cellRange = "blablabla text is too big";
	// resize
	wks.column("H").setWidth(25.0); 
	// Note MS : no way to autofit (or didn't find). Can still do it by iterating and finding biggest then adjusting to it.

	//------------------------------------- Highlight stuff. Based on Demo10.
	XLCellRange myCellRange = wks.range("B20:P25");           // create a new range for formatting
	myCellRange = "TEST COLORS";                              // write some values to the cells so we can see format changes
	XLStyleIndex baseFormat = wks.cell("B20").cellFormat();   // determine the style used in B20
	XLStyleIndex newCellStyle = cellFormats.create( cellFormats[ baseFormat ] ); // create a new style based on the style in B20
	XLStyleIndex newFillStyle = fills.create(fills[ cellFormats[ baseFormat ].fillIndex() ]); // create a new fill style based on the used fill
	
	XLColor yellow( "aaffff00" ); // a bit transparent

	fills[ newFillStyle ].setPatternType    ( XLPatternNone );
	fills[ newFillStyle ].setPatternType    ( XLPatternSolid );
	fills[ newFillStyle ].setColor          ( yellow );
	cellFormats[ newCellStyle ].setFillIndex( newFillStyle );

	myCellRange.setFormat( newCellStyle ); // assign the new format to the full range of cells

	//------------------------------------- Gradient Stuff. Based on Demo10.
	XLColor green ( "ff00ff00" );
	XLColor blue  ( "ff0000ff" );

	myCellRange = wks.range("B30:G35");           // create a new range for formatting
	myCellRange = "TEST GRADIENT";                            // write some values to the cells so we can see format changes
	baseFormat = wks.cell("B30").cellFormat();   // determine the style used in B20
	newCellStyle = cellFormats.create( cellFormats[ baseFormat ] ); // create a new style based on the style in B20
	XLStyleIndex testGradientIndex = fills.create(fills[ cellFormats[ baseFormat ].fillIndex() ]); // create a new fill style based on the used fill

	fills[ testGradientIndex ].setBackgroundColor( blue );    // setBackgroundColor only makes sense with gradient fills
	fills[ testGradientIndex ].setFillType( XLGradientFill, XLForceFillType ); 
	fills[ testGradientIndex ].setGradientType( XLGradientLinear );            
	
	// configure the gradient stops
	XLGradientStops stops = fills[ testGradientIndex ].stops();

	// first XLGradientStop
	XLStyleIndex stopIndex = stops.create();
	XLDataBarColor stopColor = stops[ stopIndex ].color();
	stops[ stopIndex ].setPosition(0.1);
	stopColor.setRgb( red );

	// second XLGradientStop
	stopIndex = stops.create(stops[stopIndex]); // create another stop using previous stop as template
	stopColor.set( yellow );
	stops[ stopIndex ].setPosition(0.5);

	cellFormats[ newCellStyle ].setFillIndex( testGradientIndex );
	myCellRange.setFormat( newCellStyle ); //Apply gradient

	// Modify same style for a different range. Get it from B30
	XLStyleIndex newCellStyle2 = cellFormats.create( cellFormats[ wks.cell("B30").cellFormat() ] );
	XLStyleIndex testGradientIndexMod = fills.create(fills[ cellFormats[ wks.cell("B30").cellFormat() ].fillIndex() ]); 
	XLColor orange ( "ffffa500" );
	XLColor purpleish ( "ff800080" );
	myCellRange = wks.range("B45:D50");
	myCellRange = "Blablabla";
	// Channge fill type to allow us to change the background color then change back to gradient
	fills[ testGradientIndexMod ].setFillType( XLPatternFill, XLForceFillType ); 
	fills[ testGradientIndexMod ].setBackgroundColor(red);  
	fills[ testGradientIndexMod ].setFillType( XLGradientFill, XLForceFillType ); 
	fills[ testGradientIndexMod ].setGradientType( XLGradientLinear );  
	// Change font color
	XLStyleIndex newFontStyle = fonts.create(fonts[ cellFormats[ wks.cell("B30").cellFormat() ].fontIndex() ]);  
	fonts[ newFontStyle ].setFontColor(purpleish);
	// configure stops (again, otherwise doesn;t work because we removed then added back)
	XLGradientStops stops2 = fills[ testGradientIndexMod ].stops();	
 	XLStyleIndex stopIndex2 = stops2.create();
	XLDataBarColor stopColor2 = stops2[ stopIndex2 ].color();
	stops2[ stopIndex2 ].setPosition(0.1);
	stopColor2.setRgb( orange );
	stopIndex = stops2.create(stops2[stopIndex2]);
	stopColor2.set( yellow );
	stops2[ stopIndex2 ].setPosition(0.9);
	// Apply
	cellFormats[ newCellStyle2 ].setFontIndex( newFontStyle );   
	cellFormats[ newCellStyle2 ].setFillIndex( testGradientIndexMod );
	myCellRange.setFormat( newCellStyle2 ); 

	//------------------------------------- Border stuff. In new sheeet.
 	doc.workbook().addWorksheet("borderTest");
	XLWorksheet wks2 = doc.workbook().worksheet(2);
	cellRange = wks2.range("A1:D25");
	cellRange = "default";

	XLStyleIndex borderFormat = borders.create();	
    XLColor black("ff000000");
    borders[borderFormat].setBottom(XLLineStyle::XLLineStyleThin, black);
    borders[borderFormat].setTop(XLLineStyle::XLLineStyleThin, black);
    borders[borderFormat].setLeft(XLLineStyle::XLLineStyleThin, black);
    borders[borderFormat].setRight(XLLineStyle::XLLineStyleThin, black);
	doc.save();
	XLStyleIndex oldFormat = cellFormats.create( cellFormats[ wks2.cell("A1").cellFormat() ] );
	XLStyleIndex newFormat = cellFormats.create(cellFormats[oldFormat]);
	cellFormats[newFormat].setBorderIndex(borderFormat);
    cellRange.setFormat(newFormat);

	doc.save();
	doc.close();
}

int main(int argc, char *argv[])
{
    QApplication app(argc, argv);
    setupDatabase();

    QFile file(":/styles/darkstyle.qss");
    if(file.open(QFile::ReadOnly)) {
        QString style = QLatin1String(file.readAll());
        app.setStyleSheet(style);
        file.close();
    } else {
        qWarning() << "Could not open stylesheet file";
    }

    // Excel formatting tests
   excelTestCustomIndividual();
   excelTestCustomRange();

    MainWindow window;
    window.show();
    return app.exec();
}
