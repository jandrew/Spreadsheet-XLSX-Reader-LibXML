#!/usr/bin/env perl
my	$dir 	= './';
my	$tests	= 'Spreadsheet/XLSX/Reader/';
my	$up		= '../';
for my $next ( <*> ){
	if( ($next eq 't') and -d $next ){
		$dir	= './t/';
		$up		= '';
		last;
	}
}

use	TAP::Formatter::Console;
my $formatter = TAP::Formatter::Console->new({
					jobs => 1,
					#~ verbosity => 1,
				});
my	$args ={
		lib =>[
			$up . 'lib',
			$up,
			#~ $up . '../Log-Shiras/lib',
		],
		test_args =>{
			load_test					=>[],
			pod_test					=>[],
			error_test					=>[],
			cell_to_column_row_test		=>[],
			default_format_test			=>[],
			excel_format_string_test	=>[],
			format_interface_test		=>[],
			stacked_flag_test			=>[],
			xmlreader_extract_file_test	=>[ $dir . 'test_files/' ],
			no_pivot_bug				=>[ $dir . 'test_files/' ],
			temp_dir_bug				=>[ $dir . 'test_files/' ],
			open_by_worksheet_bug		=>[ $dir . 'test_files/' ],
			has_chart_bug				=>[ $dir . 'test_files/' ],
			workbook_test				=>[ $dir . 'test_files/' ],
			types_test					=>[ $dir . 'test_files/' ],
			empty_sharedstrings_bug		=>[ $dir . 'test_files/' ],
			shared_strings_bug			=>[ $dir . 'test_files/' ],
			percent_file_bug			=>[ $dir . 'test_files/' ],
			hidden_formatting_bug		=>[ $dir . 'test_files/' ],
			losing_is_hidden_bug		=>[ $dir . 'test_files/' ],
			merged_areas_test			=>[ $dir . 'test_files/' ],
			read_xlsm_feature			=>[ $dir . 'test_files/' ],
			text_in_worksheet_bug		=>[ $dir . 'test_files/' ],
			xml_to_perl_test			=>[ $dir . 'test_files/' ],
			open_xml_files				=>[ $dir . 'test_files/' ],
			quote_in_styles_bug			=>[ $dir . 'test_files/' ],
			open_MySQL_files			=>[ $dir . 'test_files/' ],
			generic_reader_test			=>[ $dir . 'test_files/xl/' ],
			worksheet_test				=>[ $dir . 'test_files/xl/' ],
			cell_test					=>[ $dir . 'test_files/xl/' ],
			styles_sheet_test			=>[ $dir . 'test_files/xl/' ],
			worksheet_to_row_test		=>[ $dir . 'test_files/xl/' ],
			shared_strings_interface_test	=>[ $dir . 'test_files/xl/' ],
		},
		formatter => $formatter,
	};
my	@tests =(
		[  $dir . '01-load.t', 'load_test' ],
		[  $dir . '02-pod.t', 'pod_test' ],
		[  $dir . $tests . 'LibXML/01-types.t', 'types_test' ],
		[  $dir . $tests . 'LibXML/02-error.t', 'error_test' ],
		[  $dir . $tests . 'LibXML/04-xml_reader.t', 'generic_reader_test' ],
		[  $dir . $tests . 'LibXML/XMLReader/01-extract_file.t', 'xmlreader_extract_file_test' ],
		[  $dir . $tests . 'LibXML/05-cell_to_column_row.t', 'cell_to_column_row_test' ],
		[  $dir . $tests . 'LibXML/11-xml_to_perl_data.t', 'xml_to_perl_test' ],
		[  $dir . $tests . 'LibXML/07-fmt_default.t', 'default_format_test' ],
		[  $dir . $tests . 'LibXML/08-parse_excel_fmt_string.t', 'excel_format_string_test' ],
		[  $dir . $tests . 'LibXML/12-format_interface.t', 'format_interface_test' ],
		[  $dir . $tests . 'LibXML/13-sharedstrings_interface.t', 'shared_strings_interface_test' ],
		[  $dir . $tests . 'LibXML/15-styles_interface.t', 'styles_sheet_test' ],
		[  $dir . $tests . 'LibXML/16-worksheet_to_row.t', 'worksheet_to_row_test' ],
		[  $dir . $tests . 'LibXML/10-worksheet.t', 'worksheet_test' ],
		[  $dir . $tests . 'LibXML/09-cell.t', 'cell_test' ],
		[  $dir . $tests . '01-libxml.t', 'workbook_test' ],
		[  $dir . $tests . 'LibXML/52-merge_function_alignment.t', 'merged_areas_test' ],
		[  $dir . $tests . 'LibXML/20-empty_sharedstrings_bug.t', 'empty_sharedstrings_bug' ],
		[  $dir . $tests . '05-chart_bug.t', 'has_chart_bug' ],
		[  $dir . $tests . '04-no_pivot_bug.t', 'no_pivot_bug' ],
		[  $dir . $tests . '03-temp_dir_bug.t', 'temp_dir_bug' ],
		[  $dir . $tests . '02-open_by_worksheet_bug.t', 'open_by_worksheet_bug' ],
		[  $dir . $tests . '49-shared_strings_bug.t', 'shared_strings_bug' ],
		[  $dir . $tests . '51-percent_file_bug.t', 'percent_file_bug' ],
		[  $dir . $tests . '60-hidden_formatting_bug.t', 'hidden_formatting_bug' ],
		[  $dir . $tests . '06-stacked_flags.t', 'stacked_flag_test' ],
		[  $dir . $tests . '07-losing_is_hidden_bug.t', 'losing_is_hidden_bug' ],
		[  $dir . $tests . '68-read_xlsm_feature.t', 'read_xlsm_feature' ],
		[  $dir . $tests . 'LibXML/81-text_in_worksheet_bug.t', 'text_in_worksheet_bug' ],
		[  $dir . $tests . 'LibXML/91-quote_in_style_line_bug.t', 'quote_in_styles_bug' ],
		[  $dir . $tests . '08-open_spreadsheet_ml_files.t', 'open_xml_files' ],
		[  $dir . $tests . '09-open_MySQL_data.t', 'open_MySQL_files' ],
	);
use	TAP::Harness;
use	TAP::Parser::Aggregator;
my	$harness	= TAP::Harness->new( $args );
my	$aggregator	= TAP::Parser::Aggregator->new;
	$aggregator->start();
	$harness->aggregate_tests( $aggregator, @tests );
	$aggregator->stop();
use Test::More;
explain $formatter->summary($aggregator);
pass( "Test Harness Testing complete" );
done_testing();