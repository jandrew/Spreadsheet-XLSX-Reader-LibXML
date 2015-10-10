#########1 Test File for Spreadsheet::XLSX::Reader::LibXML::Worksheet   7#########8#########9
#!/usr/bin/env perl
my ( $lib, $test_file );
BEGIN{
	$ENV{PERL_TYPE_TINY_XS} = 0;
	my	$start_deeper = 1;
	$lib		= 'lib';
	$test_file	= 't/test_files/xl/';
	for my $next ( <*> ){
		if( ($next eq 't') and -d $next ){
			$start_deeper = 0;
			last;
		}
	}
	if( $start_deeper ){
		$lib		= '../../../../../' . $lib;
		$test_file	= '../../../../test_files/xl/';
	}
}
$| = 1;

use	Test::Most tests => 799;
use	Test::Moose;
use IO::File;
use XML::LibXML::Reader;
use	DateTime::Format::Flexible;
use	DateTimeX::Format::Excel;
use Type::Tiny;
use Data::Dumper;
use	Types::Standard qw(
		InstanceOf		Str			Num
		HasMethods		Bool		Enum
	);
use	MooseX::ShortCut::BuildInstance v1.8 qw( build_instance );#
use	lib
		'../../../../../../Log-Shiras/lib',
		$lib,
	;
#~ use Log::Shiras::Switchboard qw( :debug );
###LogSD	my	$operator = Log::Shiras::Switchboard->get_operator(#
#~ ###LogSD						name_space_bounds =>{
#~ ###LogSD							UNBLOCK =>{
#~ ###LogSD								log_file => 'trace',
#~ ###LogSD							},
#~ ###LogSD							main =>{
#~ ###LogSD								UNBLOCK =>{
#~ ###LogSD									log_file => 'info',
#~ ###LogSD								},
#~ ###LogSD							},
#~ ###LogSD						},
###LogSD						reports =>{
###LogSD							log_file =>[ Print::Log->new ],
###LogSD						},
###LogSD					);
###LogSD	use Log::Shiras::Telephone;
###LogSD	use Log::Shiras::UnhideDebug;
use	Spreadsheet::XLSX::Reader::LibXML::Error;
###LogSD	use Log::Shiras::UnhideDebug;
use	Spreadsheet::XLSX::Reader::LibXML::XMLReader::SharedStrings;
use	Spreadsheet::XLSX::Reader::LibXML::FmtDefault;
###LogSD	use Log::Shiras::UnhideDebug;
use	Spreadsheet::XLSX::Reader::LibXML::XMLReader::Styles;
use	Spreadsheet::XLSX::Reader::LibXML::Worksheet;
###LogSD	use Log::Shiras::UnhideDebug;
use	Spreadsheet::XLSX::Reader::LibXML::XMLReader::WorksheetToRow;
	$test_file = ( @ARGV ) ? $ARGV[0] : $test_file;
my	$shared_strings_file = $test_file;
my	$styles_file = $test_file;
	$shared_strings_file .= 'sharedStrings.xml';
	$styles_file .= 'styles.xml';
my	$test_file_2 = $test_file . 'worksheets/sheet2.xml';
	$test_file .= 'worksheets/sheet3.xml';
###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space => 'main', );
###LogSD		$phone->talk( level => 'trace', message => [ "Test file is: $test_file" ] );
my  ( 
			$test_instance, $error_instance, $styles_instance, $shared_strings_instance,
			$string_type, $date_time_type, $cell, $row_ref, $offset, $workbook_instance,
			$file_handle, $styles_handle, $shared_strings_handle, $format_instance
	);
my 			$row = 0;
my 			@class_attributes = qw(
				file						error_inst					is_hidden
				workbook_instance			sheet_type					sheet_rel_id
				sheet_id					sheet_position				sheet_name
				last_header_row				min_header_col				max_header_col
			);
my  		@class_methods = qw(
				is_sheet_hidden				is_empty_the_end			get_group_return_type
				has_min_col					has_min_row					has_max_col
				has_max_row					get_file					set_file
				has_file					clear_file					is_sheet_hidden
				sheet_id					position					get_name
				get_last_header_row			has_last_header_row			get_min_header_col
				set_min_header_col			clear_min_header_col		has_min_header_col
				get_max_header_col			set_max_header_col			clear_max_header_col
				has_max_header_col			min_row						max_row
				min_col						max_col						row_range
				col_range					get_merged_areas			is_column_hidden
				is_row_hidden				get_cell					get_next_value
				fetchrow_arrayref			fetchrow_array				set_headers
				fetchrow_hashref			set_custom_formats			has_custom_format
				get_custom_format			get_sheet_type				rel_id
			);
my			$answer_list =[
				[
					[
						{},{},{},{},'EOR','EOR','EOR',
					],
					[
						{ cell_id => 'A2', row => 1, col => 0, type => 'Text', unformatted => 'Hello', value => 'Hello' },{},{},
						{ cell_id => 'D2', row => 1, col => 3, type => 'Text', unformatted => 'my', value => 'my' },{},{},'EOR',
					],
					[
						{},{},{},{},{},{},'EOR',
					],
					[
						{},{},{ cell_id => 'C4', row => 3, col => 2, type => 'Text', unformatted => 'World', value => 'World', coercion_name => 'Excel_text_0' },{},{},{},'EOR',
					],
					[
						{},{},{},{},{},{},'EOR',
					],
					[
						{ cell_id => 'A6', row => 5, col => 0, type => 'Text', unformatted => 'Hello World', value => 'Hello World',
							get_rich_text =>[
								2, { color =>{ rgb => 'FFFF0000' }, sz => '11', b => undef, scheme => 'minor', rFont => 'Calibri', family => 2 },
								6, { color =>{ rgb => 'FF0070C0' }, sz => '20', b => undef, scheme => 'minor', rFont => 'Calibri', family => 2 },
							],
							merge_range => 'A6:B6',
						},
						{ cell_id => 'B6', row => 5, col => 1, type => 'Text', unformatted => undef, value => undef, merge_range => 'A6:B6' },{},{},{},{},'EOR',
					],
					[
						{},{ cell_id => 'B7', row => 6, col => 1, type => 'Numeric', unformatted => '69', value => '69' },{},{},{},{},'EOR',
					],
					[
						{},{ cell_id => 'B8', row => 7, col => 1, type => 'Numeric', unformatted => '27', value => '27' },{},{},
						{ cell_id => 'E8', row => 7, col => 4, type => 'Date', unformatted => '37145', value => '12-Sep-05', coercion_name => 'Excel_date_164' },{},'EOR',
					],
					[
						{},{ cell_id => 'B9', row => 8, col => 1, type => 'Numeric', unformatted => '42', value => '42', formula => 'B7-B8' },{},{},{},{},'EOR',
					],
					[
						{},{},{},{ cell_id => 'D10', row => 9, col => 3, type => 'Custom', unformatted => ' ', value => ' ', coercion_name => 'YYYYMMDD' },
						{ cell_id => 'E10', row => 9, col => 4, type => 'Custom', unformatted => '2/6/2011', value => '2011-02-06T00:00:00', coercion_name => 'Custom_date_type' },
						{ cell_id => 'F10', row => 9, col => 5, type => 'Custom', unformatted => '2/6/2011', value => '2011-02-06', coercion_name => 'YYYYMMDD' },'EOR',
					],
					[
						{ cell_id => 'A11', row => 10, col => 0, type => 'Numeric', unformatted => '2.1345678901', value => '2.13', coercion_name => 'Excel_number_2' },{},{},{},{},{},'EOR',
					],
					[
						{},{ cell_id => 'B12', row => 11, col => 1, type => 'Text', unformatted => undef, value => undef, formula => 'IF(B11>0,"Hello","")', },
						{},{ cell_id => 'D12', row => 11, col => 3, type => 'Date', unformatted => '39118', value => '6-Feb-11', coercion_name => 'Excel_date_164', merge_range => 'D12:E12' },
						{ cell_id => 'E12', row => 11, col => 4, type => 'Text', unformatted => undef, value => undef, coercion_name => 'Excel_date_164', merge_range => 'D12:E12' },{},'EOR',
					],
					[
						{},{},{},{},{},{},'EOR',
					],
					[
						{},{},{ cell_id => 'C14', row => 13, col => 2, type => 'Text', unformatted => ' ', value => ' ', has_coercion => '', },
						{ cell_id => 'D14', row => 13, col => 3, type => 'Custom', unformatted => '39118', value => '2011-2-06', coercion_name => 'Worksheet_Custom_0', },
						{ cell_id => 'E14', row => 13, col => 4, type => 'Date', unformatted => '39118', value => '6-Feb-11', coercion_name => 'Excel_date_164', },{},'EOF',
					],
				],
				[
					{ cell_id => 'A2', row => 1, col => 0, type => 'Text', unformatted => 'Hello', value => 'Hello' },
					{ cell_id => 'D2', row => 1, col => 3, type => 'Text', unformatted => 'my', value => 'my' },
					{ cell_id => 'C4', row => 3, col => 2, type => 'Text', unformatted => 'World', value => 'World', coercion_name => 'Excel_text_0' },
					{ cell_id => 'A6', row => 5, col => 0, type => 'Text', unformatted => 'Hello World', value => 'Hello World',
						get_rich_text =>[
							2, { color =>{ rgb => 'FFFF0000' }, sz => '11', b => undef, scheme => 'minor', rFont => 'Calibri', family => 2 },
							6, { color =>{ rgb => 'FF0070C0' }, sz => '20', b => undef, scheme => 'minor', rFont => 'Calibri', family => 2 },
						],
						merge_range => 'A6:B6',
					},
					{ cell_id => 'B7', row => 6, col => 1, type => 'Numeric', unformatted => '69', value => '69' },
					{ cell_id => 'B8', row => 7, col => 1, type => 'Numeric', unformatted => '27', value => '27' },
					{ cell_id => 'E8', row => 7, col => 4, type => 'Date', unformatted => '37145', value => '12-Sep-05', coercion_name => 'Excel_date_164' },
					{ cell_id => 'B9', row => 8, col => 1, type => 'Numeric', unformatted => '42', value => '42', formula => 'B7-B8' },
					{ cell_id => 'D10', row => 9, col => 3, type => 'Custom', unformatted => ' ', value => ' ', coercion_name => 'YYYYMMDD' },
					{ cell_id => 'E10', row => 9, col => 4, type => 'Custom', unformatted => '2/6/2011', value => '2011-02-06T00:00:00', coercion_name => 'Custom_date_type' },
					{ cell_id => 'F10', row => 9, col => 5, type => 'Custom', unformatted => '2/6/2011', value => '2011-02-06', coercion_name => 'YYYYMMDD' },
					{ cell_id => 'A11', row => 10, col => 0, type => 'Numeric', unformatted => '2.1345678901', value => '2.13', coercion_name => 'Excel_number_2' },
					#~ { cell_id => 'B12', row => 11, col => 1, type => 'Text', unformatted => undef, value => undef, formula => 'IF(B11>0,"Hello","")', },
					{ cell_id => 'D12', row => 11, col => 3, type => 'Date', unformatted => '39118', value => '6-Feb-11', coercion_name => 'Excel_date_164', merge_range => 'D12:E12' },
					{ cell_id => 'C14', row => 13, col => 2, type => 'Text', unformatted => ' ', value => ' ', has_coercion => '', },
					{ cell_id => 'D14', row => 13, col => 3, type => 'Custom', unformatted => '39118', value => '2011-2-06', coercion_name => 'Worksheet_Custom_0', },
					{ cell_id => 'E14', row => 13, col => 4, type => 'Date', unformatted => '39118', value => '6-Feb-11', coercion_name => 'Excel_date_164', },
					undef,
				],
				[
					[
						{},{},{},{},{},{},
					],
					[
						{ cell_id => 'A2', row => 1, col => 0, type => 'Text', unformatted => 'Hello', value => 'Hello' },{},{},
						{ cell_id => 'D2', row => 1, col => 3, type => 'Text', unformatted => 'my', value => 'my' },{},{},
					],
					[
						{},{},{},{},{},{},
					],
					[
						{},{},{ cell_id => 'C4', row => 3, col => 2, type => 'Text', unformatted => 'World', value => 'World', coercion_name => 'Excel_text_0' },{},{},{},
					],
					[
						{},{},{},{},{},{},
					],
					[
						{ cell_id => 'A6', row => 5, col => 0, type => 'Text', unformatted => 'Hello World', value => 'Hello World',
							get_rich_text =>[
								2, { color =>{ rgb => 'FFFF0000' }, sz => '11', b => undef, scheme => 'minor', rFont => 'Calibri', family => 2 },
								6, { color =>{ rgb => 'FF0070C0' }, sz => '20', b => undef, scheme => 'minor', rFont => 'Calibri', family => 2 },
							],
							merge_range => 'A6:B6',
						},
						{ cell_id => 'B6', row => 5, col => 1, type => 'Text', unformatted => undef, value => undef, merge_range => 'A6:B6' },{},{},{},{},
					],
					[
						{},{ cell_id => 'B7', row => 6, col => 1, type => 'Numeric', unformatted => '69', value => '69' },{},{},{},{},
					],
					[
						{},{ cell_id => 'B8', row => 7, col => 1, type => 'Numeric', unformatted => '27', value => '27' },{},{},
						{ cell_id => 'E8', row => 7, col => 4, type => 'Date', unformatted => '37145', value => '12-Sep-05', coercion_name => 'Excel_date_164' },{},
					],
					[
						{},{ cell_id => 'B9', row => 8, col => 1, type => 'Numeric', unformatted => '42', value => '42', formula => 'B7-B8' },{},{},{},{},
					],
					[
						{},{},{},{ cell_id => 'D10', row => 9, col => 3, type => 'Custom', unformatted => ' ', value => ' ', coercion_name => 'YYYYMMDD' },
						{ cell_id => 'E10', row => 9, col => 4, type => 'Custom', unformatted => '2/6/2011', value => '2011-02-06T00:00:00', coercion_name => 'Custom_date_type' },
						{ cell_id => 'F10', row => 9, col => 5, type => 'Custom', unformatted => '2/6/2011', value => '2011-02-06', coercion_name => 'YYYYMMDD' },
					],
					[
						{ cell_id => 'A11', row => 10, col => 0, type => 'Numeric', unformatted => '2.1345678901', value => '2.13', coercion_name => 'Excel_number_2' },{},{},{},{},{},
					],
					[
						{},{ cell_id => 'B12', row => 11, col => 1, type => 'Text', unformatted => undef, value => undef, formula => 'IF(B11>0,"Hello","")', },
						{},{ cell_id => 'D12', row => 11, col => 3, type => 'Date', unformatted => '39118', value => '6-Feb-11', coercion_name => 'Excel_date_164', merge_range => 'D12:E12' },
						{ cell_id => 'E12', row => 11, col => 4, type => 'Text', unformatted => undef, value => undef, coercion_name => 'Excel_date_164', merge_range => 'D12:E12' },{},
					],
					[
						{},{},{},{},{},{},
					],
					[
						{},{},{ cell_id => 'C14', row => 13, col => 2, type => 'Text', unformatted => ' ', value => ' ', has_coercion => '', },
						{ cell_id => 'D14', row => 13, col => 3, type => 'Custom', unformatted => '39118', value => '2011-2-06', coercion_name => 'Worksheet_Custom_0', },
						{ cell_id => 'E14', row => 13, col => 4, type => 'Date', unformatted => '39118', value => '6-Feb-11', coercion_name => 'Excel_date_164', },{},
					],
					[],
				],
				[
					[],
					['Hello',undef,undef,'my',],
					[],
					[undef,undef,'World'],
					[],
					['Hello World',undef],
					[undef,'69'],
					[undef,'27',undef,undef,'37145'],
					[undef,'42',],
					[undef,undef,undef,' ','2/6/2011','2/6/2011',],
					['2.1345678901'],
					[undef,undef,undef,'39118',undef],
					[],
					[undef,undef,' ','39118','39118'],
					'EOF',
				],
				[
					['Row Labels', '2016-02-06',  '2017-02-14', '2018-02-03', 'Grand Total', ],
					{ 'Row Labels' => 'Blue', '2016-02-06' => '10', '2017-02-14' => '7', },
					{ 'Row Labels' => 'Omaha', '2018-02-03' => '2', },
					{ 'Row Labels' => 'Red', '2016-02-06' => '30', '2017-02-14' => '5', '2018-02-03' => '3', },
					{ 'Row Labels' => 'Grand Total', '2016-02-06' => '40', '2017-02-14' => '12', '2018-02-03' => '5', },
					'EOF',
				],
				
			];
###LogSD		$phone->talk( level => 'info', message => [ "easy questions ..." ] );
lives_ok{
	
			my	@args_list	= ( system_type => 'apple_excel' );
			my	$converter	= DateTimeX::Format::Excel->new( @args_list );
			my	$string_via	= sub{ 
									my	$str = $_[0];
									return DateTime::Format::Flexible->parse_datetime( $str );#my	$dt	= 
									#~ return $dt->format_cldr( 'yyyy-M-d' );
								};
			my	$num_via	= sub{
									my	$num = $_[0];
									return $converter->parse_datetime( $num );#my	$dt = 
									#~ return $dt->format_cldr( 'yyyy-M-d' );
								};
			my	$date_time_from_value = Type::Coercion->new(
				type_coercion_map => [ Num, $num_via, Str, $string_via, ],
			);
			$date_time_type = Type::Tiny->new(
					name		=> 'Custom_date_type',
					constraint	=> sub{ ref($_) eq 'DateTime' },
					coercion	=> $date_time_from_value,
				);
			###LogSD	$phone->talk( level => 'trace', message => [
			###LogSD		"Check coercion:", $date_time_from_value->coerce( '2/6/2011' ) ] );
			$string_type = Type::Tiny->new(
					display_name => 'YYYYMMDD',
					constraint	=> sub{
						!$_ or (
						$_ =~ /^\d{4}\-(\d{2})-(\d{2})$/ and
						$1 > 0 and $1 < 13 and $2 > 0 and $2 < 32 )
					},
					coercion	=> Type::Coercion->new(
						type_coercion_map =>[
							$date_time_type->coercibles, sub{ my $tmp = $date_time_type->coerce( $_ ); $tmp->format_cldr( 'yyyy-MM-dd' ) },
						],
					),
			);
			$error_instance	= 	build_instance(
									package => 'ErrorInstance',
									superclasses =>[ 'Spreadsheet::XLSX::Reader::LibXML::Error' ],
									should_warn => 0,
				###LogSD				log_space	=> 'Test',
								);
			$format_instance	=  	Spreadsheet::XLSX::Reader::LibXML::FmtDefault->new(
										epoch_year	=> 1904,
										error_inst	=> $error_instance,
				###LogSD				log_space	=> 'Test',
									);
###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space => 'main', );
###LogSD		$phone->talk( level => 'trace', message => [ "Format instance:", $format_instance ] );
			$styles_instance	=	build_instance(
									package => 'StylesInstance',
									superclasses	=> [ 'Spreadsheet::XLSX::Reader::LibXML::XMLReader::Styles' ],
									format_inst => $format_instance,
									file		=> $styles_file,
									error_inst	=> $error_instance,
			###LogSD				log_space	=> 'Test',
								);
			$shared_strings_instance	=	Spreadsheet::XLSX::Reader::LibXML::XMLReader::SharedStrings->new(
											error_inst	=> $error_instance,
											file		=> $shared_strings_file,
			###LogSD						log_space	=> 'Test::SharedStrings',
										);
			$workbook_instance =	build_instance(
										package	=> 'WorkbookInstance',
										add_methods =>{
											get_epoch_year => sub{ return '1904' },
										},
										add_attributes =>{
											error_inst =>{
												isa			=> 	HasMethods[qw(
																	error set_error clear_error set_warnings if_warn
																) ],
												clearer		=> '_clear_error_inst',
												reader		=> 'get_error_inst',
												required	=> 1,
												handles =>[ qw(
													error set_error clear_error set_warnings if_warn
												) ],
											},
											format_inst =>{
												isa		=> HasMethods[qw( 
																set_error_inst				set_excel_region
																set_target_encoding			get_defined_excel_format
																set_defined_excel_formats	change_output_encoding
																set_epoch_year				set_cache_behavior
																set_date_behavior			get_defined_conversion
																parse_excel_format_string							 )],	
												writer	=> 'set_format_instance',
												reader	=> 'get_format_instance',
												handles =>[qw(
																get_defined_excel_format 	parse_excel_format_string
																change_output_encoding 		get_epoch_year
																)],
											},
											empty_is_end =>{
												isa		=> Bool,
												writer	=> 'set_empty_is_end',
												reader	=> 'is_empty_the_end',
												default	=> 0,
											},
											from_the_edge =>{
												isa		=> Bool,
												reader	=> '_starts_at_the_edge',
												writer	=> 'set_from_the_edge',
												default	=> 1,
											},
											shared_strings_instance =>{
												isa			=> HasMethods[ 'get_shared_string_position' ],
												predicate	=> '_has_shared_strings_file',
												handles		=>[ 'get_shared_string_position' ],
											},
											styles_instance =>{
												isa			=> HasMethods[ 'get_format_position' ],
												predicate	=> '_has_styles_file',
												handles =>[ 'get_format_position' ],
											},
											count_from_zero =>{
												isa		=> Bool,
												reader	=> 'counting_from_zero',
												writer	=> 'set_count_from_zero',
												default	=> 1,
											},
											file_boundary_flags =>{
												isa			=> Bool,
												reader		=> 'boundary_flag_setting',
												writer		=> 'change_boundary_flag',
												default		=> 1,
												required	=> 1,
											},
											group_return_type =>{
												isa		=> Enum[qw( unformatted value instance )],
												reader	=> 'get_group_return_type',
												writer	=> 'set_group_return_type',
												default	=> 'instance',
											},
											datetime_dates =>{
												isa		=> Bool,
												reader	=> 'get_date_behavior',
												writer	=> 'set_date_behavior',
												default	=> 0,
											},
											empty_return_type =>{
												isa		=> Enum[qw( empty_string undef_string )],
												reader	=> 'get_empty_return_type',
												writer	=> 'set_empty_return_type',
												default	=> 'undef_string',
											},
											values_only =>{
												isa		=> Bool,
												reader	=> 'get_values_only',
												writer	=> 'set_values_only',
												default	=> 0,
											},
										},
										error_inst => $error_instance,
										format_inst => $format_instance,
										styles_instance => $styles_instance,
										shared_strings_instance => $shared_strings_instance,
									);
			$test_instance	=	build_instance(
									package				=> 'WorksheetInterfaceTest',
									superclasses 		=>[ 'Spreadsheet::XLSX::Reader::LibXML::XMLReader::WorksheetToRow' ],
									roles 				=>[ 'Spreadsheet::XLSX::Reader::LibXML::Worksheet' ],
									file				=> $test_file,
									error_inst			=> $error_instance,
									workbook_instance	=> $workbook_instance,
			###LogSD				log_space			=> 'Test',
								);
			$test_instance->set_custom_formats(
								E10	=> $date_time_type,
								10	=> $string_type,
								D14	=> 'yyyy-m-dd',
							);
}										"Prep a test WorksheetInterfaceTest instance";
map{ 
has_attribute_ok
			$test_instance, $_,
										"Check that the WorksheetInterfaceTest instance has the -$_- attribute"
} 			@class_attributes;
map{
can_ok		$test_instance, $_,
} 			@class_methods;

###LogSD		$phone->talk( level => 'trace', message => [ 'Test instance:', $test_instance ] );
###LogSD		$phone->talk( level => 'info', message => [ "hardest questions ..." ] );
is			$test_instance->min_row, 0,
										"check that it knows what the lowest row number is";
is			$test_instance->min_col, 0,
										"check that it knows what the lowest column number is";
is			$test_instance->max_row, undef,
										"check that it knows what the highest row number is (not)";
is			$test_instance->max_col, undef,
										"check that it knows what the highest column number is (not)";
is_deeply	[$test_instance->row_range], [0,undef],
										"check for a correct row range";
is_deeply	[$test_instance->col_range], [0,undef],
										"check for a correct column range";

explain									"Test get_cell";
#~ ###LogSD		$operator->add_name_space_bounds( {
#~ ###LogSD				UNBLOCK =>{
#~ ###LogSD					log_file => 'trace',
#~ ###LogSD				},
#~ ###LogSD		} );
explain		$test_instance->get_cell( 1, 0 )->row;
is			$test_instance->get_cell( 1, 0 )->row, 1,
										"Check that you can call the same cell twice";
is			$test_instance->get_cell( 1, 0 )->row, 1,
										"...and again";
			#~ exit 1;
			my ( $row_min, $row_max ) = $test_instance->row_range();
			my ( $col_min, $col_max ) = $test_instance->col_range();
			my $test_group = 0;
			no warnings 'uninitialized';
			INITIALRUN: for my $row ( 0 .. 13 ) {
            for my $col ( 0 .. 6 ) {
				
###LogSD	my $expose_row = 13; my $expose_col = 6;
###LogSD	if( $row == $expose_row and $col == $expose_col ){
###LogSD		$operator->add_name_space_bounds( {
#~ ###LogSD			Test =>{
###LogSD				UNBLOCK =>{
###LogSD					log_file => 'trace',
###LogSD				},
#~ ###LogSD			},
###LogSD		} );
###LogSD	}

###LogSD	elsif( $row > $expose_row or ( $row == $expose_row and $col > $expose_col ) ){
###LogSD		exit 1;
###LogSD	}




lives_ok{	$cell = $test_instance->get_cell( $row, $col ) }
										"Get anything at the cell for row -$row- and col -$col-";
###LogSD	$phone->talk( level => 'debug', message => [ "cell:", $cell ] );
			if( !ref $answer_list->[$test_group]->[$row]->[$col] ){
is			$cell, $answer_list->[$test_group]->[$row]->[$col],
										"Check for the correct end flag: $answer_list->[$test_group]->[$row]->[$col]";
			#~ last INITIALRUN;
			}elsif( scalar( keys %{$answer_list->[$test_group]->[$row]->[$col]} ) == 0 ){
is			!$cell, 1,					"Check that an expected empty cell really is empty";
			}else{
			for my $key ( keys %{$answer_list->[$test_group]->[$row]->[$col]} ){
###LogSD	$phone->talk( level => 'debug', message => [ "checking method:", $key ] );
is_deeply	$cell->$key, $answer_list->[$test_group]->[$row]->[$col]->{$key},
										"Checking the method -$key- value at row -$row- and column -$col- is: $answer_list->[$test_group]->[$row]->[$col]->{$key}";
			}
			}
            }
			}
			#~ exit 1;
is			$test_instance->get_cell( 1, 6 ), 'EOR',
										"For row -1- and column -6- get a correct end of row flag: EOR";
is			$workbook_instance->change_boundary_flag( 0 ), 0,
										"Turn boundary flags off";
is			$test_instance->get_cell( 2, 7 ), undef,
										"Check row -2- and column -7- (end of row position) should return undef";
is			$test_instance->get_cell( 14, 0 ), undef,
										"Check row -14- and -0- (end of file position) should return undef";

ok			$test_instance->set_values_only( 1 ),
										'Turn values_only on';
explain									"Test get_next_value with values_only = 1";
			$test_group++;
			$cell = undef;
			my $x = 0;
			VALUERUN: while( !$cell or ref $cell eq 'Spreadsheet::XLSX::Reader::LibXML::Cell' ){#$x < 103 and ( )

###LogSD	my $expose_position = 17;
###LogSD	if( $x == $expose_position ){
###LogSD		$operator->add_name_space_bounds( {
###LogSD			main =>{
###LogSD				UNBLOCK =>{
###LogSD					log_file => 'debug',
###LogSD				},
###LogSD			},
###LogSD			Test =>{
###LogSD				Worksheet =>{
###LogSD					Interface =>{
###LogSD						UNBLOCK =>{
###LogSD							log_file => 'trace',
###LogSD						},
###LogSD						_hidden =>{
###LogSD							UNBLOCK =>{
###LogSD								log_file => 'fatal',
###LogSD							},
###LogSD						},
###LogSD					},
###LogSD				},
###LogSD			},
###LogSD		} );
###LogSD	}

###LogSD	elsif( $x > $expose_position){
###LogSD		exit 1;
###LogSD	}

lives_ok{	$cell = $test_instance->get_next_value }
										"Get the next cell with a value for position: $x";
###LogSD	$phone->talk( level => 'debug', message => [ "cell:", $cell ] );
			if( !ref $answer_list->[$test_group]->[$x] ){
is			$cell, $answer_list->[$test_group]->[$x],
										"Check for the correct end of file flag: undef";
			last VALUERUN;
			}else{
			for my $key ( keys %{$answer_list->[$test_group]->[$x]} ){
is_deeply	$cell->$key, $answer_list->[$test_group]->[$x]->{$key},
										"Checking the method -$key- value at position -$x- is: $answer_list->[$test_group]->[$x]->{$key}";
			}
			}
			$x++;
            }
			#~ exit 1;
is			$test_instance->set_values_only( 0 ), 0,
										'Turn values_only off';

explain									"Test fetchrow_arrayref";
			$test_group++;
			$row_ref = [];
			$x = 0;
			ROWREFRUN: while( ref $row_ref eq 'ARRAY' ){#

###LogSD	my $expose_array = 14;
###LogSD	if( $x == $expose_array ){
###LogSD		$operator->add_name_space_bounds( {
###LogSD			main =>{
###LogSD				UNBLOCK =>{
###LogSD					log_file => 'debug',
###LogSD				},
###LogSD			},
###LogSD			Test =>{
###LogSD				Worksheet =>{
###LogSD					Interface =>{
###LogSD						UNBLOCK =>{
###LogSD							log_file => 'trace',
###LogSD						},
###LogSD						_hidden =>{
###LogSD							UNBLOCK =>{
###LogSD								log_file => 'warn',
###LogSD							},
###LogSD						},
###LogSD					},
###LogSD				},
###LogSD			},
###LogSD		} );
###LogSD	}

###LogSD	elsif( $x > $expose_array ){
###LogSD		exit 1;
###LogSD	}

lives_ok{	$row_ref = $test_instance->fetchrow_arrayref }
										"Get the next fetchrow_arrayref for row: $x";
###LogSD	$phone->talk( level => 'debug', message => [ "row:", $row_ref ] );
###LogSD	$phone->talk( level => 'debug', message => [ "looking for answer:", $answer_list->[$test_group]->[$x] ] );
			if( !ref $answer_list->[$test_group]->[$x] ){###LogSD	$phone->talk( level => 'debug', message => [ "row:", $row ] );
###LogSD	$phone->talk( level => 'debug', message => [ "Found and -end- flag: $answer_list->[$test_group]->[$x]" ] );
is			$row_ref, $answer_list->[$test_group]->[$x],
										"Check for the correct end of file flag: undef";
			last ROWREFRUN;
			}else{
			my $col = 0;
			for my $cell ( @{$answer_list->[$test_group]->[$x]} ){
###LogSD	$phone->talk( level => 'debug', message => [ "Testing cell:", $row_ref->[$col] , "..against", $cell ] );
			if( scalar( keys %$cell ) > 0 ){
###LogSD	$phone->talk( level => 'debug', message => [ 'There is a value in the cell' ] );
is			ref( $row_ref->[$col] ), 'Spreadsheet::XLSX::Reader::LibXML::Cell',	
										"Check that row -$x- and column -$col- does have a value as expected";
			for my $key ( keys %$cell ){
is_deeply	$row_ref->[$col]->$key, $cell->{$key},
										"Checking the method -$key- value at row -$x- and column -$col- is: $cell->{$key}";
			}
			}else{
is			!$row_ref->[$col], 1,		"Check that an expected empty cell really is empty";
			}	
			$col++;
			}
			}
			$x++;
            }
			#~ exit 1;
explain									"Test fetchrow_array";
			$test_group++;
ok			$test_instance->change_boundary_flag( 1 ),
										"Turn boundary flag reporting back on";
ok			$workbook_instance->set_group_return_type( 'unformatted' ),
										"Return just the cell coerced values rather than a Cell instance";
			$x = 0;
			ROWARRAYRUN: while( !$row_ref or !$row_ref->[0] or $row_ref->[0] ne 'EOF' ){

###LogSD	my $expose_array = 20;
###LogSD	if( $x == $expose_array ){
###LogSD		$operator->add_name_space_bounds( {
###LogSD			main =>{
###LogSD				UNBLOCK =>{
###LogSD					log_file => 'debug',
###LogSD				},
###LogSD			},
###LogSD			Test =>{
###LogSD				Worksheet =>{
###LogSD					Interface =>{
###LogSD						UNBLOCK =>{
###LogSD							log_file => 'trace',
###LogSD						},
###LogSD						_hidden =>{
###LogSD							UNBLOCK =>{
###LogSD								log_file => 'warn',
###LogSD							},
###LogSD						},
###LogSD					},
###LogSD				},
###LogSD			},
###LogSD		} );
###LogSD	}

###LogSD	elsif( $x > $expose_array ){
###LogSD		exit 1;
###LogSD	}

lives_ok{	my @result = $test_instance->fetchrow_array;
			$row_ref = scalar( @result ) ? [@result] : []; }
										"Get the next fetchrow_array for row: $x";
###LogSD	$phone->talk( level => 'debug', message => [ "row:", $row_ref ] );
			if( !ref $answer_list->[$test_group]->[$x] ){
###LogSD	$phone->talk( level => 'debug', message => [ "Found and -end- flag: $answer_list->[$test_group]->[$x]" ] );
is			$row_ref->[0], $answer_list->[$test_group]->[$x],
										"Check for the correct end of file flag: EOF";
			last ROWARRAYRUN;
			}else{
is_deeply	$row_ref, $answer_list->[$test_group]->[$x],
										"..and validate the returned values";
			}
			$x++;
            }
###LogSD	$phone->talk( level => 'trace', message => [ "Master column data:", $test_instance->_get_column_formats ] );
is_deeply	[ $test_instance->is_column_hidden( 0 .. 5 ) ], [ 0, 0, 1, 1, 0, 0 ],
										'Check that the sheet knows which columns are hidden - by number';
###LogSD		$operator->add_name_space_bounds( {
###LogSD			Test =>{
###LogSD				Worksheet =>{
###LogSD					Interface =>{
###LogSD						is_column_hidden =>{
###LogSD							UNBLOCK =>{
###LogSD								log_file => 'trace',
###LogSD							},
###LogSD						},
###LogSD					},
###LogSD				},
###LogSD			},
#~ ###LogSD			main =>{
#~ ###LogSD				UNBLOCK =>{
#~ ###LogSD					log_file => 'trace',
#~ ###LogSD				},
#~ ###LogSD			},
###LogSD		} );
is_deeply	[ $test_instance->is_column_hidden( 'A', 'B', 'C', 'D', 'E', 'F' ) ], [ 0, 0, 1, 1, 0, 0 ],
										'Check that the sheet knows which columns are hidden - by letter';
#~ ###LogSD		$operator->add_name_space_bounds( {
#~ ###LogSD			Test =>{
#~ ###LogSD				Worksheet =>{
#~ ###LogSD					Interface =>{
#~ ###LogSD						is_row_hidden =>{
#~ ###LogSD							UNBLOCK =>{
#~ ###LogSD								log_file => 'trace',
#~ ###LogSD							},
#~ ###LogSD						},
#~ ###LogSD					},
#~ ###LogSD				},
#~ ###LogSD			},
#~ ###LogSD			main =>{
#~ ###LogSD				UNBLOCK =>{
#~ ###LogSD					log_file => 'trace',
#~ ###LogSD				},
#~ ###LogSD			},
#~ ###LogSD		} );
###LogSD	$phone->talk( level => 'trace', message => [ "Row range:", $test_instance->row_range ] );
is_deeply	[ $test_instance->is_row_hidden( 0 .. 15 ) ], [ 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 1, 0, 0, undef, undef, undef ],
										'Check that the sheet knows which rows are hidden - by number';
is			$test_instance->max_row, 13,
										"check that it knows what the highest row number is";
is			$test_instance->max_col, 5,
										"check that it knows what the highest column number is";
is_deeply	[$test_instance->row_range], [0,13],
										"check for a correct row range";
is_deeply	[$test_instance->col_range], [0,5],
										"check for a correct column range";

explain									"Test fetchrow_hashref";
			$test_group++;
ok			$workbook_instance->set_group_return_type( 'value' ),
										"Set the group_return_type to: value";
ok			$test_instance = WorksheetInterfaceTest->new(
								file			=> $test_file_2,
								error_inst		=> $error_instance,
								workbook_instance => $workbook_instance,
			###LogSD			log_space			=> 'Test',
							),			'Build another connection to a different worksheet';
ok			$test_instance->set_custom_formats(
								E10	=> $date_time_type,
								10	=> $string_type,
								D14	=> 'yyyy-mm-dd',
							),			'Add the custom formats';
is 			$test_instance->fetchrow_hashref( 1 ), undef,
										"Check that a fetchrow_hashref call returns undef without a set header";
is			$test_instance->error, 		"Headers must be set prior to calling fetchrow_hashref",
										"..and check for the correct error message";
###LogSD	my $test_header_load = 0;
###LogSD	if( $test_header_load ){
###LogSD		$operator->add_name_space_bounds( {
###LogSD			Test =>{
###LogSD				Worksheet =>{
###LogSD					Interface =>{
###LogSD						UNBLOCK =>{
###LogSD							log_file => 'trace',
###LogSD						},
#~ ###LogSD						_hidden =>{
#~ ###LogSD							UNBLOCK =>{
#~ ###LogSD								log_file => 'warn',
#~ ###LogSD							},
#~ ###LogSD						},
###LogSD					},
###LogSD				},
###LogSD				Styles =>{
###LogSD					UNBLOCK =>{
###LogSD						log_file => 'trace',
###LogSD					},
###LogSD				},
###LogSD			},
###LogSD		} );
###LogSD	}
is_deeply	$test_instance->set_headers( 1 ), $answer_list->[$test_group]->[0],
										"Set the headers for building a hashref";
###LogSD	if( $test_header_load ){
###LogSD		exit 1;
###LogSD	}
ok			$test_instance->set_max_header_col( 3 ),,
										"Set the maximum header column";
			$row_ref = undef;
			$x = 1;
			HASHREFRUN: while( !$row_ref or ref $row_ref eq 'HASH' ){

###LogSD	my $expose_ref = 7;
###LogSD	if( $x == $expose_ref ){
###LogSD		$operator->add_name_space_bounds( {
###LogSD			main =>{
###LogSD				UNBLOCK =>{
###LogSD					log_file => 'trace',
###LogSD				},
###LogSD			},
###LogSD			Test =>{
###LogSD				Worksheet =>{
###LogSD					Interface =>{
###LogSD						UNBLOCK =>{
###LogSD							log_file => 'trace',
###LogSD						},
###LogSD						_hidden =>{
###LogSD							UNBLOCK =>{
###LogSD								log_file => 'warn',
###LogSD							},
###LogSD						},
###LogSD					},
###LogSD				},
###LogSD			},
###LogSD		} );
###LogSD	}

###LogSD	elsif( $x > $expose_ref ){
###LogSD		exit 1;
###LogSD	}

lives_ok{	$row_ref = $test_instance->fetchrow_hashref }
										"Get the next fetchrow_hashref for row: $x";
###LogSD	$phone->talk( level => 'debug', message => [ "row:", $row_ref ] );
			if( !ref $answer_list->[$test_group]->[$x] ){
###LogSD	$phone->talk( level => 'debug', message => [ "Found and -end- flag: $answer_list->[$test_group]->[$x]" ] );
is			$row_ref, $answer_list->[$test_group]->[$x],
										"Check for the correct end of file flag: EOF";
			last HASHREFRUN;
			}else{
is_deeply	$row_ref, $answer_list->[$test_group]->[$x],
										"..and validate the returned values";# . Dumper( $answer_list->[$test_group]->[$x] )
			}
			$x++;
            }
is			$test_instance->fetchrow_hashref( 1 ), undef,
										"Check that calling for a row above or at the header in the table fails";
is			$test_instance->error, "The requested row -1- is at or above the bottom of the header rows ( 1 )",
										"..with the correct error message";
is_deeply	$test_instance->fetchrow_hashref( 3 ), $answer_list->[$test_group]->[$x + $offset - 3],
										"Get an arbitrary hashref row - and check the values";
explain 								"...Test Done";
done_testing();

###LogSD	package Print::Log;
###LogSD	use Data::Dumper;
###LogSD	sub new{
###LogSD		bless {}, shift;
###LogSD	}
###LogSD	sub add_line{
###LogSD		shift;
###LogSD		my @input = ( ref $_[0]->{message} eq 'ARRAY' ) ? 
###LogSD						@{$_[0]->{message}} : $_[0]->{message};
###LogSD		my ( @print_list, @initial_list );
###LogSD		no warnings 'uninitialized';
###LogSD		for my $value ( @input ){
###LogSD			push @initial_list, (( ref $value ) ? Dumper( $value ) : $value );
###LogSD		}
###LogSD		for my $line ( @initial_list ){
###LogSD			$line =~ s/\n$//;
###LogSD			$line =~ s/\n/\n\t\t/g;
###LogSD			push @print_list, $line;
###LogSD		}
###LogSD		printf( "| level - %-6s | name_space - %-s\n| line  - %04d   | file_name  - %-s\n\t:(\t%s ):\n", 
###LogSD					$_[0]->{level}, $_[0]->{name_space},
###LogSD					$_[0]->{line}, $_[0]->{filename},
###LogSD					join( "\n\t\t", @print_list ) 	);
###LogSD		use warnings 'uninitialized';
###LogSD	}

###LogSD	1;