#########1 Test File for Spreadsheet::XLSX::Reader::XMLReader::Worksheet        8#########9
#!/usr/bin/env perl
my ( $lib, $test_file, $test_file2, $styles_file, $worksheet );
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

use	Test::Most tests => 38;
use	Test::Moose;
use Data::Dumper;
use	MooseX::ShortCut::BuildInstance qw( build_instance );
use Types::Standard qw( Bool HasMethods );
use	lib
		'../../../../../../../Log-Shiras/lib',
		$lib,
	;
#~ use Log::Shiras::Switchboard qw( :debug );#
###LogSD	my	$operator = Log::Shiras::Switchboard->get_operator(
###LogSD						name_space_bounds =>{
###LogSD							UNBLOCK =>{
###LogSD								log_file => 'warn',
###LogSD							},
###LogSD							main =>{
###LogSD								UNBLOCK =>{
###LogSD									log_file => 'debug',
###LogSD								},
###LogSD							},
#~ ###LogSD							Test =>{
#~ ###LogSD								get_merged_areas =>{
#~ ###LogSD									UNBLOCK =>{
#~ ###LogSD										log_file => 'trace',
#~ ###LogSD									},
#~ ###LogSD								},
#~ ###LogSD							},
###LogSD						},
###LogSD						reports =>{
###LogSD							log_file =>[ Print::Log->new ],
###LogSD						},
###LogSD					);
###LogSD	use Log::Shiras::Telephone;
###LogSD	use Log::Shiras::UnhideDebug;
use	Spreadsheet::XLSX::Reader::LibXML::Error;
#~ use	Spreadsheet::XLSX::Reader::LibXML::XMLReader;
use	Spreadsheet::XLSX::Reader::LibXML::XMLReader::WorksheetToRow;
use	Spreadsheet::XLSX::Reader::LibXML::FmtDefault;
use	DateTimeX::Format::Excel;
use	DateTime::Format::Flexible;
use	Type::Coercion;
use	Type::Tiny;

	$test_file	= ( @ARGV ) ? $ARGV[0] : $test_file;
	$test_file2 = $test_file . 'merged.xlsx';
	$test_file .= 'xl/worksheets/sheet3.xml';
	
###LogSD	my	$log_space	= 'Test';
###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space => 'main', );
###LogSD		$phone->talk( level => 'trace', message => [ "Test file is: $test_file" ] );
my  ( 
			$test_instance, $error_instance, $workbook_instance, $file_handle, $format_instance,
	);
my			$answer_ref = [
				[ [ 5, 0, 5, 1 ], [ 11, 3, 11, 4 ] ],
				[ [ 0, 1, 1, 2 ], [  1, 0,  2, 0 ] ],
				[ 'A1', '',],
				[ 'B1', 1, 'B1:C2' ],
				[ 'C1', 1, 'B1:C2' ],
				[ 'A2', 1, 'A2:A3' ],
				[ 'B2', 1, 'B1:C2' ],
				[ 'C2', 1, 'B1:C2' ],
				[ 'A3', 1, 'A2:A3' ],
				[ 'B3', '',],
				[ 'C3', '',],
			];

lives_ok{
			$error_instance		= 	Spreadsheet::XLSX::Reader::LibXML::Error->new( should_warn => 0 );
			$format_instance	=  	Spreadsheet::XLSX::Reader::LibXML::FmtDefault->new(
										error_inst	=> $error_instance,
				###LogSD				log_space	=> 'Test',
									);
			$workbook_instance	= build_instance(
									package		=> 'WorkbookInstance',
									add_methods =>{
										counting_from_zero			=> sub{ return 1 },
										boundary_flag_setting		=> sub{},
										change_boundary_flag		=> sub{},
										_has_shared_strings_file	=> sub{ return 1 },
										get_shared_string_position	=> sub{},
										_has_styles_file			=> sub{},
										get_format_position			=> sub{},
										get_epoch_year				=> sub{ return 1904 },
										get_group_return_type		=> sub{},
										set_group_return_type		=> sub{},
										get_date_behavior			=> sub{},
										set_date_behavior			=> sub{},
										get_empty_return_type		=> sub{ return 'undef_string' },
										get_values_only				=> sub{},
										set_values_only				=> sub{},
									},
									add_roles => [ 'Spreadsheet::XLSX::Reader::LibXML::CellToColumnRow' ],
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
									},
									error_inst => $error_instance,
									format_instance => $format_instance,
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
			#~ $test_instance	= Spreadsheet::XLSX::Reader::LibXML::XMLReader::WorksheetToRow->new(
				#~ file				=> $test_file,
				#~ error_inst			=> $error_instance,
				sheet_name			=> 'Sheet3',
				#~ workbook_instance	=> $workbook_instance,
			#~ ###LogSD	log_space	=> 'Test',
			#~ );
			###LogSD	$phone->talk( level => 'info', message =>[ "Loaded test instance" ] );
}										"Prep a new Worksheet instance";
###LogSD		$phone->talk( level => 'debug', message => [ "Max row is:" . $test_instance->max_row ] );
			my $x = 0;
is_deeply	$test_instance->get_merged_areas, $answer_ref->[$x++],
				'Check that get_merged_areas works';
use Spreadsheet::XLSX::Reader::LibXML ':like_ParseExcel';
lives_ok{
			$test_instance	= Spreadsheet::XLSX::Reader::LibXML->new(
				file_name => $test_file2,
			###LogSD	log_space	=> 'Test',
			);
			###LogSD	$phone->talk( level => 'info', message =>[ "Loaded test instance" ] );
}										"Prep a workbook instance";
###LogSD		$operator->add_name_space_bounds( {
#~ ###LogSD			Test =>{
#~ ###LogSD				Worksheet =>{
#~ ###LogSD					_build_out_the_cell =>{
###LogSD						UNBLOCK =>{
###LogSD							log_file => 'trace',
###LogSD						},
#~ ###LogSD					},
#~ ###LogSD					_get_next_value_cell =>{
#~ ###LogSD						UNBLOCK =>{
#~ ###LogSD							log_file => 'trace',
#~ ###LogSD						},
#~ ###LogSD					},
#~ ###LogSD				},
#~ ###LogSD			},
###LogSD		} );
lives_ok{ 	$worksheet = $test_instance->worksheet( 'Blad1' ) }
										'..and pull the worksheet in question (Blad1)';
###LogSD		$operator->add_name_space_bounds( {
#~ ###LogSD			Test =>{
#~ ###LogSD				Worksheet =>{
#~ ###LogSD					_build_out_the_cell =>{
###LogSD						UNBLOCK =>{
###LogSD							log_file => 'warn',
###LogSD						},
#~ ###LogSD					},
#~ ###LogSD					_get_next_value_cell =>{
#~ ###LogSD						UNBLOCK =>{
#~ ###LogSD							log_file => 'trace',
#~ ###LogSD						},
#~ ###LogSD					},
#~ ###LogSD				},
#~ ###LogSD			},
###LogSD		} );
is_deeply	$worksheet->get_merged_areas, $answer_ref->[$x++],
				'Check that get_merged_areas works';

			for my $row ( 0.. 2 ){
			for my $column ( 0..2 ){
###LogSD	my $reveal_row = 0;
###LogSD	my $reveal_col = 1;
###LogSD	my $revealed = 0;
###LogSD	if( $row == $reveal_row and $column == $reveal_col ){
###LogSD		$operator->add_name_space_bounds( {
###LogSD			Test =>{
###LogSD				Worksheet =>{
###LogSD					_build_out_the_cell =>{
###LogSD						UNBLOCK =>{
###LogSD							log_file => 'trace',
###LogSD						},
###LogSD					},
###LogSD					_get_next_value_cell =>{
###LogSD						UNBLOCK =>{
###LogSD							log_file => 'trace',
###LogSD						},
###LogSD					},
###LogSD				},
###LogSD			},
###LogSD		} );
###LogSD		$revealed = 1;
###LogSD	}
###LogSD	elsif( $revealed and ( $row != $reveal_row or $column != $reveal_col) ){
###LogSD		exit 1;
###LogSD		$operator->add_name_space_bounds( {
###LogSD			Test =>{
###LogSD				Worksheet =>{
###LogSD					_build_out_the_cell =>{
###LogSD						UNBLOCK =>{
###LogSD							log_file => 'warn',
###LogSD						},
###LogSD					},
###LogSD					_get_next_value_cell =>{
###LogSD						UNBLOCK =>{
###LogSD							log_file => 'warn',
###LogSD						},
###LogSD					},
###LogSD				},
###LogSD			},
###LogSD		} );
###LogSD		$revealed = 0;
###LogSD	}
ok			my $cell = $worksheet->get_cell( $row, $column ),
										"get the cell for row -$row- and column -$column-";
is			$cell->cell_id, $answer_ref->[$x]->[0],
										"Confirm the cellId: $answer_ref->[$x]->[0]";
is			$cell->is_merged, $answer_ref->[$x]->[1],
										"..and test for the correct merge state: $answer_ref->[$x]->[1]";
			if( $cell->is_merged ){
is			$cell->merge_range, $answer_ref->[$x]->[2],
										"..and check for the participating merge range: $answer_ref->[$x]->[2]";
			}
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
###LogSD		printf( "| level - %-6s | name_space - %-s\n| line  - %04d   | file_name  - %-s\n\t:(\t%s ):\n", 
###LogSD					$_[0]->{level}, $_[0]->{name_space},
###LogSD					$_[0]->{line}, $_[0]->{filename},
###LogSD					join( "\n\t\t", @print_list ) 	);
###LogSD		use warnings 'uninitialized';
###LogSD	}

###LogSD	1;