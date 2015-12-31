#!/usr/bin/env perl
### Test that the pod files run
use Test::More;
use Test::Pod 1.48;
my	$up		= '../';
for my $next ( <*> ){
	if( ($next eq 't') and -d $next ){
		### <where> - found the t directory - must be using prove ...
		$up	= '';
		last;
	}
}
pod_file_ok( $up . 	'README.pod',
						"The README file has good POD" );
pod_file_ok( $up . 	'lib/Spreadsheet/XLSX/Reader/LibXML/Cell.pm',
						"The Spreadsheet::XLSX::Reader::LibXML::Cell file has good POD" );
pod_file_ok( $up . 	'lib/Spreadsheet/XLSX/Reader/LibXML/CellToColumnRow.pm',
						"The Spreadsheet::XLSX::Reader::LibXML::CellToColumnRow file has good POD" );
pod_file_ok( $up . 	'lib/Spreadsheet/XLSX/Reader/LibXML/Chartsheet.pm',
						"The Spreadsheet::XLSX::Reader::LibXML::Chartsheet file has good POD" );
pod_file_ok( $up . 	'lib/Spreadsheet/XLSX/Reader/LibXML/Error.pm',
						"The Spreadsheet::XLSX::Reader::LibXML::Error file has good POD" );
pod_file_ok( $up . 	'lib/Spreadsheet/XLSX/Reader/LibXML/FmtDefault.pm',
						"The Spreadsheet::XLSX::Reader::LibXML::FmtDefault file has good POD" );
pod_file_ok( $up . 	'lib/Spreadsheet/XLSX/Reader/LibXML/FormatInterface.pm',
						"The Spreadsheet::XLSX::Reader::LibXML::FormatInterface file has good POD" );
pod_file_ok( $up . 	'lib/Spreadsheet/XLSX/Reader/LibXML/ParseExcelFormatStrings.pm',
						"The Spreadsheet::XLSX::Reader::LibXML::ParseExcelFormatStrings file has good POD" );
pod_file_ok( $up . 	'lib/Spreadsheet/XLSX/Reader/LibXML/Row.pm',
						"The Spreadsheet::XLSX::Reader::LibXML::Row file has good POD" );
pod_file_ok( $up . 	'lib/Spreadsheet/XLSX/Reader/LibXML/SharedStrings.pm',
						"The Spreadsheet::XLSX::Reader::LibXML::SharedStrings file has good POD" );
pod_file_ok( $up . 	'lib/Spreadsheet/XLSX/Reader/LibXML/Styles.pm',
						"The Spreadsheet::XLSX::Reader::LibXML::Styles file has good POD" );
pod_file_ok( $up . 	'lib/Spreadsheet/XLSX/Reader/LibXML/Types.pm',
						"The Spreadsheet::XLSX::Reader::LibXML::Types file has good POD" );
pod_file_ok( $up . 	'lib/Spreadsheet/XLSX/Reader/LibXML/WorkbookFileInterface.pm',
						"The Spreadsheet::XLSX::Reader::LibXML::WorkbookFileInterface file has good POD" );
pod_file_ok( $up . 	'lib/Spreadsheet/XLSX/Reader/LibXML/WorkbookMetaInterface.pm',
						"The Spreadsheet::XLSX::Reader::LibXML::WorkbookMetaInterface file has good POD" );
pod_file_ok( $up . 	'lib/Spreadsheet/XLSX/Reader/LibXML/WorkbookPropsInterface.pm',
						"The Spreadsheet::XLSX::Reader::LibXML::WorkbookPropsInterface file has good POD" );
pod_file_ok( $up . 	'lib/Spreadsheet/XLSX/Reader/LibXML/WorkbookRelsInterface.pm',
						"The Spreadsheet::XLSX::Reader::LibXML::WorkbookRelsInterface file has good POD" );
pod_file_ok( $up . 	'lib/Spreadsheet/XLSX/Reader/LibXML/Worksheet.pm',
						"The Spreadsheet::XLSX::Reader::LibXML::Worksheet file has good POD" );
pod_file_ok( $up . 	'lib/Spreadsheet/XLSX/Reader/LibXML/WorksheetToRow.pm',
						"The Spreadsheet::XLSX::Reader::LibXML::WorksheetToRow file has good POD" );
pod_file_ok( $up . 	'lib/Spreadsheet/XLSX/Reader/LibXML/XMLReader.pm',
						"The Spreadsheet::XLSX::Reader::LibXML::XMLReader file has good POD" );
pod_file_ok( $up . 	'lib/Spreadsheet/XLSX/Reader/LibXML/XMLToPerlData.pm',
						"The Spreadsheet::XLSX::Reader::LibXML::XMLToPerlData file has good POD" );
pod_file_ok( $up . 	'lib/Spreadsheet/XLSX/Reader/LibXML/XMLReader/ExtractFile.pm',
						"The Spreadsheet::XLSX::Reader::LibXML::XMLReader::ExtractFile file has good POD" );
pod_file_ok( $up . 	'lib/Spreadsheet/XLSX/Reader/LibXML/XMLReader/NamedStyles.pm',
						"The Spreadsheet::XLSX::Reader::LibXML::XMLReader::NamedStyles file has good POD" );
pod_file_ok( $up . 	'lib/Spreadsheet/XLSX/Reader/LibXML/XMLReader/PositionStyles.pm',
						"The Spreadsheet::XLSX::Reader::LibXML::XMLReader::PositionStyles file has good POD" );
pod_file_ok( $up . 	'lib/Spreadsheet/XLSX/Reader/LibXML/ZipReader.pm',
						"The Spreadsheet::XLSX::Reader::LibXML::ZipReader file has good POD" );
pod_file_ok( $up . 	'lib/Spreadsheet/XLSX/Reader/LibXML/ZipReader/ExtractFile.pm',
						"The Spreadsheet::XLSX::Reader::LibXML::ZipReader::ExtractFile file has good POD" );
pod_file_ok( $up . 	'lib/Spreadsheet/XLSX/Reader/LibXML.pm',
						"The Spreadsheet::XLSX::Reader::LibXML file has good POD" );
done_testing();