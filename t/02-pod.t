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
pod_file_ok( $up . 	'lib/Spreadsheet/XLSX/Reader/LibXML/Types.pm',
						"The Spreadsheet::XLSX::Reader::LibXML::Types file has good POD" );
pod_file_ok( $up . 	'lib/Spreadsheet/XLSX/Reader/LibXML/Error.pm',
						"The Spreadsheet::XLSX::Reader::LibXML::Error file has good POD" );
pod_file_ok( $up . 	'lib/Spreadsheet/XLSX/Reader/LibXML/XMLReader.pm',
						"The Spreadsheet::XLSX::Reader::LibXML::XMLReader file has good POD" );
pod_file_ok( $up . 	'lib/Spreadsheet/XLSX/Reader/LibXML/CellToColumnRow.pm',
						"The Spreadsheet::XLSX::Reader::LibXML::CellToColumnRow file has good POD" );
pod_file_ok( $up . 	'lib/Spreadsheet/XLSX/Reader/LibXML/XMLReader/XMLToPerlData.pm',
						"The Spreadsheet::XLSX::Reader::LibXML::XMLReader::XMLToPerlData file has good POD" );
pod_file_ok( $up . 	'lib/Spreadsheet/XLSX/Reader/LibXML/XMLReader/Worksheet.pm',
						"The Spreadsheet::XLSX::Reader::LibXML::XMLReader::Worksheet file has good POD" );
pod_file_ok( $up . 	'lib/Spreadsheet/XLSX/Reader/LibXML/XMLReader/Chartsheet.pm',
						"The Spreadsheet::XLSX::Reader::LibXML::XMLReader::Chartsheet file has good POD" );
pod_file_ok( $up . 	'lib/Spreadsheet/XLSX/Reader/LibXML/FmtDefault.pm',
						"The Spreadsheet::XLSX::Reader::LibXML::FmtDefault file has good POD" );
pod_file_ok( $up . 	'lib/Spreadsheet/XLSX/Reader/LibXML/ParseExcelFormatStrings.pm',
						"The Spreadsheet::XLSX::Reader::LibXML::ParseExcelFormatStrings file has good POD" );
pod_file_ok( $up . 	'lib/Spreadsheet/XLSX/Reader/LibXML/XMLReader/SharedStrings.pm',
						"The Spreadsheet::XLSX::Reader::LibXML::XMLReader::SharedStrings file has good POD" );
pod_file_ok( $up . 	'lib/Spreadsheet/XLSX/Reader/LibXML/XMLReader/CalcChain.pm',
						"The Spreadsheet::XLSX::Reader::LibXML::XMLReader::CalcChain file has good POD" );
pod_file_ok( $up . 	'lib/Spreadsheet/XLSX/Reader/LibXML/XMLReader/Styles.pm',
						"The Spreadsheet::XLSX::Reader::LibXML::XMLReader::Styles file has good POD" );
pod_file_ok( $up . 	'lib/Spreadsheet/XLSX/Reader/LibXML/GetCell.pm',
						"The Spreadsheet::XLSX::Reader::LibXML::GetCell file has good POD" );
pod_file_ok( $up . 	'lib/Spreadsheet/XLSX/Reader/LibXML/Cell.pm',
						"The Spreadsheet::XLSX::Reader::LibXML::Cell file has good POD" );
pod_file_ok( $up . 	'lib/Spreadsheet/XLSX/Reader/LibXML/Worksheet.pod',
						"The Spreadsheet::XLSX::Reader::LibXML::Worksheet file has good POD" );
pod_file_ok( $up . 	'lib/Spreadsheet/XLSX/Reader/LibXML/Chartsheet.pod',
						"The Spreadsheet::XLSX::Reader::LibXML::Chartsheet file has good POD" );
pod_file_ok( $up . 	'lib/Spreadsheet/XLSX/Reader/LibXML/Styles.pod',
						"The Spreadsheet::XLSX::Reader::LibXML::Styles file has good POD" );
pod_file_ok( $up . 	'lib/Spreadsheet/XLSX/Reader/LibXML/SharedStrings.pod',
						"The Spreadsheet::XLSX::Reader::LibXML::SharedStrings file has good POD" );
pod_file_ok( $up . 	'lib/Spreadsheet/XLSX/Reader/LibXML/CalcChain.pod',
						"The Spreadsheet::XLSX::Reader::LibXML::CalcChain file has good POD" );
pod_file_ok( $up . 	'lib/Spreadsheet/XLSX/Reader/LibXML.pm',
						"The Spreadsheet::XLSX::Reader::LibXML file has good POD" );
done_testing();