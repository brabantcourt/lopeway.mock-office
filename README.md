lopeway.mock-office
===================

An attempt to mock the [JavaScript API for Office](//msdn.microsoft.com/en-us/library/fp142185.aspx) in order to facilitate debugging using a regular browser and its tools. The implementation is currently limited to part of the implementation of the [TableBinding](http://msdn.microsoft.com/en-us/library/fp160977.aspx) object although much of the infrastructure to support this object will support the creation of the additional objects.

# Installation

If you don't already, your project will need to reference [jQuery](//jquery.com) and [KnockoutJS](//knockoutjs.com) as these are required by the spreadsheet 'emulator'.

    <script src="//ajax.aspnetcdn.com/ajax/jQuery/jquery-1.8.1.min.js"></script>
    <script src="//ajax.aspnetcdn.com/ajax/knockout/knockout-2.1.0.js"></script>

Replace the reference to Microsoft's JavaScript API for Office:

    <!--<script src="//.../office.js"></script>-->

With a reference to the mock version:

	<script src ="//.../lopeway.office.mock.1.2.js"></script>

To override the mock library's default settings you must precede the lopeway.office.mock file reference to a configuration JSON of the form:

	lopeway = {
        excel: {
            config: {
                debug: true,				// 'true' to display tables/sheets as they're rendered; 'false' to hide the display
                id: "lopeway-workbook",		// ID of the DIV to which rendered content will be appended
                sheets: "sheets",			// ID of the DIV to which rendered sheets will be appended
                tables: "tables",			// ID of the DIV to which rendered tables will be appended
                workbook: {					// Definition of the Excel spreadsheet
                    name: "Book1",			// Unused
                    sheets: [],				// Format { name: 'sheetname' }
                    tables: []				// Format { name: 'tablename', cols = [c1, c2...], rows = [[r1c1,r1c2...],[r2c1,r2c2...] }
                }
            }
        }
    }


