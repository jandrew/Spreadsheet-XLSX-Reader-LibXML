#########1 Test File for Spreadsheet::XLSX::Reader::XMLReader::WorksheetToRow   8#########9
#!/usr/bin/env perl
my ( $lib, $test_file, $styles_file );
BEGIN{
	$ENV{PERL_TYPE_TINY_XS} = 0;
	my	$start_deeper = 1;
	$lib		= 'lib';
	$test_file	= 't/test_files/';
	for my $next ( <*> ){
		if( ($next eq 't') and -d $next ){
			$start_deeper = 0;
			last;
		}
	}
	if( $start_deeper ){
		$lib		= '../../../../../../' . $lib;
		$test_file	= '../../../../../test_files/';
	}
}
$| = 1;

use Test::Most tests => 57;
use Test::Moose;
use MooseX::ShortCut::BuildInstance qw( build_instance );
use Types::Standard qw( Bool HasMethods );
use	lib
		'../../../../../../../Log-Shiras/lib',
		$lib,
	;
use	Data::Dumper;
#~ use Log::Shiras::Switchboard qw( :debug );#
###LogSD	my	$operator = Log::Shiras::Switchboard->get_operator(
###LogSD						name_space_bounds =>{
###LogSD							UNBLOCK =>{
###LogSD								log_file => 'warn',
###LogSD							},
###LogSD							main =>{
###LogSD								UNBLOCK =>{
###LogSD									log_file => 'info',
###LogSD								},
###LogSD							},
###LogSD						reports =>{
###LogSD							log_file =>[ Print::Log->new ],
###LogSD						},
###LogSD					);
###LogSD	use Log::Shiras::Telephone;
###LogSD	use Log::Shiras::UnhideDebug;
use	Spreadsheet::XLSX::Reader::LibXML::Error;
use	Spreadsheet::XLSX::Reader::LibXML::XMLReader::WorksheetToRow;

	$test_file	= ( @ARGV ) ? $ARGV[0] : $test_file;
	$test_file .= 'string_in_worksheet.xml';
	
###LogSD	my	$log_space	= 'Test';
###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space => 'main', );
###LogSD		$phone->talk( level => 'trace', message => [ "Test file is: $test_file" ] );
my  ( 
			$test_instance, $error_instance, $workbook_instance, $file_handle, $format_instance, $shared_strings_instance
	);
my			$answer_ref = [
				[
					{ r => 'A828', cell_row => 828, cell_type => 'Text', cell_col => 1, s => '5', cell_xml_value => '2947'},
					{ r => 'B828', cell_col => 2, s => '5', cell_xml_value => 0, cell_type => 'Text', cell_row => 828 },
					{ r => 'C828', s => '7', cell_col => 3, cell_xml_value => '827', cell_type => 'Numeric', cell_row => 828 },
					{ r => 'D828', cell_type => 'Numeric', cell_row => 828, cell_xml_value => '40', cell_col => 4, s => '5' },
					{ r => 'E828', s => '5', cell_col => 5, cell_xml_value => '2493', cell_row => 828, cell_type => 'Text' },
					{ r => 'F828', cell_xml_value => '2428', s => '5', cell_col => 6, cell_row => 828, cell_type => 'Text' },
					{ r => 'G828', cell_type => 'Text', cell_row => 828, cell_col => 7, s => '5', cell_xml_value => '308' },
					{ r => 'H828', cell_xml_value => '311', s => '5', cell_col => 8, cell_type => 'Text', cell_row => 828 },
					{
					  'cell_col' => 9,
					  's' => '5',
					  'cell_formula' => 'A828',
					  'cell_xml_value' => '092-318',
					  'cell_row' => 828,
					  'cell_type' => 'Text',
					  'r' => 'I828'
					},
					{
					  'cell_col' => 10,
					  's' => '7',
					  'cell_xml_value' => '311',
					  'r' => 'J828',
					  'cell_row' => 828,
					  'cell_type' => 'Text'
					},
					 {
					  'cell_row' => 828,
					  'cell_type' => 'Text',
					  'r' => 'K828',
					  'cell_xml_value' => '308',
					  'cell_col' => 11,
					  's' => '5'
					},
					{
					  'r' => 'L828',
					  'cell_type' => 'Numeric',
					  'cell_row' => 828,
					  'cell_xml_value' => '42104',
					  's' => '6',
					  'cell_col' => 12
					},
					'EOF',
				],
			];
###LogSD	$phone->talk( level => 'info', message => [ "easy questions ..." ] );

lives_ok{
			$workbook_instance	= build_instance(
									package		=> 'WorkbookInstance',
									add_methods =>{
										counting_from_zero			=> sub{ return 0 },
										boundary_flag_setting		=> sub{},
										change_boundary_flag		=> sub{},
										_has_shared_strings_file	=> sub{ return 0 },
										_has_styles_file			=> sub{},
										get_format_position			=> sub{},
										get_epoch_year				=> sub{ return 1904 },
										get_group_return_type		=> sub{},
										set_group_return_type		=> sub{},
										get_date_behavior			=> sub{},
										set_date_behavior			=> sub{},
										get_empty_return_type		=> sub{ return 'undef_string' },
										get_values_only				=> sub{ 1 },
										set_values_only				=> sub{},
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
											default => sub{ Spreadsheet::XLSX::Reader::LibXML::Error->new},
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
										format_instance =>{
											isa		=> HasMethods[qw( 
															set_error_inst				set_excel_region
															set_target_encoding			get_defined_excel_format
															set_defined_excel_formats	change_output_encoding
															set_epoch_year				set_cache_behavior
															set_date_behavior			get_defined_conversion		
															parse_excel_format_string							)],	
											writer	=> 'set_format_instance',
											reader	=> 'get_format_instance',
											handles =>[qw(
															get_defined_excel_format 	parse_excel_format_string
															change_output_encoding		)],
										},
										_shared_strings_instance =>{
											isa			=> HasMethods[ 'get_shared_string_position' ],
											predicate	=> '_has_shared_strings_file',
											writer		=> '_set_shared_strings_instance',
											reader		=> '_get_shared_strings_instance',
											clearer		=> '_clear_shared_strings',
											handles		=>{
												'get_shared_string_position' => 'get_shared_string_position',
												_demolish_shared_strings => 'DEMOLISH',
											},
										},
									},
								);
			$test_instance	= Spreadsheet::XLSX::Reader::LibXML::XMLReader::WorksheetToRow->new(
				file				=> $test_file,
				workbook_instance	=> $workbook_instance,
			###LogSD	log_space	=> 'Test',
			);
			###LogSD	$phone->talk( level => 'info', message =>[ "Loaded test instance" ] );
}										"Prep a new WorksheetToRow instance";

###LogSD		$phone->talk( level => 'debug', message => [ "Max row is:" . $test_instance->_max_row ] );
is			$test_instance->_min_row, 1,
										"check that it knows what the lowest row number is";
is			$test_instance->_min_col, 1,
										"check that it knows what the lowest column number is";
is			$test_instance->_max_row, 1208,
										"check that it knows what the highest row number is";
is			$test_instance->_max_col, 50,
										"check that it knows what the highest column number is";

explain									"read through value cells ...";
			my $test = 0;
			for my $y (1..2){
			my $result;
explain									"Running cycle: $y";
			my $x = 0;
			while( !$result or $result ne 'EOF' ){
				
###LogSD	my $expose = 0;
###LogSD	if( $x == $expose and $y == 1 ){
###LogSD		$operator->add_name_space_bounds( {
#~ ###LogSD			Test =>{
#~ ###LogSD				_get_next_value_cell =>{
###LogSD					UNBLOCK =>{
###LogSD						log_file => 'trace',
###LogSD					},
#~ ###LogSD				},
#~ ###LogSD			},
###LogSD		} );
###LogSD	}

###LogSD	elsif( $x > ($expose + 0) and $y > 0 ){
###LogSD		exit 1;
###LogSD	}

lives_ok{	$result = $test_instance->_get_next_value_cell }
										"_get_next_value_cell test -$test- iteration -$y- from sheet position: $x";
			#~ print Dumper( $result );
###LogSD	$phone->talk( level => 'debug', message => [ "result at position -$x- is:", $result,
###LogSD		'Against answer:', $answer_ref->[$x], ] );
is_deeply	$result, $answer_ref->[$test]->[$x],"..........and see if test -$test- iteration -$y- from sheet position -$x- has good info";
			$x++;
			}
			}
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
###LogSD		printf( "| level - %-8s | name_space - %-s\n| line  - %07d | file_name  - %-s\n\t:(\t%s ):\n", 
###LogSD					$_[0]->{level}, $_[0]->{name_space},
###LogSD					$_[0]->{line}, $_[0]->{filename},
###LogSD					join( "\n\t\t", @print_list ) 	);
###LogSD		use warnings 'uninitialized';
###LogSD	}

###LogSD	1;