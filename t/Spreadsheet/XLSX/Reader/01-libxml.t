#########1 Test File for Spreadsheet::XLSX::Reader::LibXML  6#########7#########8#########9
#!/usr/bin/env perl
my ( $lib, $test_file );
#~ $Module::Load::Conditional::CACHE = undef;
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
		$lib		= '../../../../' . $lib;
		$test_file	= '../../../test_files/'
	}
}
$| = 1;

use	Test::Most tests => 125;
use	Test::Moose;
use Data::Dumper;
use	lib	'../../../../../Log-Shiras/lib',
		'../../../../../MooseX-ShortCut-BuildInstance/lib',
		$lib,
	;
#~ use Log::Shiras::Switchboard v0.21 qw( :debug );#
###LogSD	my	$operator = Log::Shiras::Switchboard->get_operator(
###LogSD						name_space_bounds =>{
###LogSD							UNBLOCK =>{
###LogSD								log_file => 'info',
###LogSD							},
###LogSD							build_class =>{
###LogSD								UNBLOCK =>{
###LogSD									log_file => 'warn',
###LogSD								},
###LogSD							},
###LogSD							build_instance =>{
###LogSD								UNBLOCK =>{
###LogSD									log_file => 'warn',
###LogSD								},
###LogSD							},
###LogSD							Test =>{
###LogSD								StylesInstance =>{
#~ ###LogSD									XMLReader =>{
#~ ###LogSD										DEMOLISH =>{
###LogSD											UNBLOCK =>{
###LogSD												log_file => 'warn',
###LogSD											},
#~ ###LogSD										},
#~ ###LogSD									},
###LogSD								},
###LogSD								SharedStringsInstance =>{
#~ ###LogSD									XMLReader =>{
#~ ###LogSD										DEMOLISH =>{
###LogSD											UNBLOCK =>{
###LogSD												log_file => 'warn',
###LogSD											},
#~ ###LogSD										},
#~ ###LogSD									},
###LogSD								},
#~ ###LogSD								Worksheet =>{
#~ ###LogSD									XMLReader =>{
#~ ###LogSD										DEMOLISH =>{
#~ ###LogSD											UNBLOCK =>{
#~ ###LogSD												log_file => 'trace',
#~ ###LogSD											},
#~ ###LogSD										},
#~ ###LogSD									},
#~ ###LogSD								},
###LogSD								Workbook =>{
###LogSD									worksheet =>{
###LogSD										UNBLOCK =>{
###LogSD											log_file => 'warn',
###LogSD										},
###LogSD									},
###LogSD									_build_file =>{
###LogSD										UNBLOCK =>{
###LogSD											log_file => 'warn',
###LogSD										},
###LogSD									},
###LogSD									_build_dom =>{
###LogSD										UNBLOCK =>{
###LogSD											log_file => 'warn',
###LogSD										},
###LogSD									},
###LogSD									_build_reader =>{
###LogSD										UNBLOCK =>{
###LogSD											log_file => 'warn',
###LogSD										},
###LogSD									},
###LogSD								},
###LogSD							},
###LogSD						},
###LogSD						reports =>{
###LogSD							log_file =>[ Print::Log->new ],
###LogSD						},
###LogSD					);
###LogSD	use Log::Shiras::Telephone;
###LogSD	use Log::Shiras::UnhideDebug;
use Spreadsheet::XLSX::Reader::LibXML;
$test_file = ( @ARGV ) ? $ARGV[0] : $test_file;
$test_file .= 'TestBook.xlsx';
my  ( 
		$error_instance, $parser, $workbook, $row_ref,
	);
my	$answer_ref = [
		[qw( Category Total Date )],
		[qw( Red 5 2017-2-14 )],
		[qw( Blue 7 2017-2-14 )],
		[qw( Omaha 2 2018-2-3 )],
		[qw( Red 3 2018-2-3 )],
		[qw( Red 30 2016-2-6 )],
		[qw( Blue 10 2016-2-6 )],
		'EOF',
		[ 'Superbowl Audibles', 'Column Labels', undef, undef, undef, ],
		[ 'Row Labels', '2016-2-6', '2017-2-14', '2018-2-3', 'Grand Total' ],
		[ 'Blue', 10, 7, undef, 17 ,],
		[ 'Omaha', undef, undef, 2, 2, ],
		[ 'Red', 30, 5, 3, 38, ],
		[ 'Grand Total', 40, 12, 5, 57, ],
		'EOF',
		[undef,undef,undef,undef,undef,undef,],
		['Hello',undef,undef,'my',undef,undef,],
		[undef,undef,undef,undef,undef,undef,],
		[undef,undef,'World',undef,undef,undef,],
		[undef,undef,undef,undef,undef,undef,],
		['Hello World',undef,undef,undef,undef,undef,],
		[undef,'69',undef,undef,undef,undef,],
		[undef,'27',undef,undef,'12-Sep-05',undef,],
		[undef,'42',,undef,undef,undef,undef,],
		[undef,undef,undef,undef,'2/6/2011','6-Feb-11',],
		['2.13',undef,undef,undef,undef,undef,],
		[undef,'',undef,'6-Feb-11',undef,undef,],
		[undef,undef,undef,undef,undef,undef,],
		[undef,undef,' ','39118','6-Feb-11',undef,],
		'EOF',
	];
my 			@class_attributes = qw(
				error_inst					file_name					file_handle
				sheet_parser				count_from_zero				file_boundary_flags
				empty_is_end				from_the_edge				default_format_list
				format_string_parser		group_return_type
				empty_return_type
			);
my  		@class_methods = qw(
				new							parse						worksheet
				worksheets					get_error_inst				error
				set_error					clear_error					set_warnings
				if_warn						set_file_name				has_file_name
				set_file_name				has_file_name				creator
				modified_by					date_created				date_modified
				set_parser_type				get_parser_type				counting_from_zero
				set_count_from_zero			boundary_flag_setting		change_boundary_flag
				set_empty_is_end			is_empty_the_end			set_from_the_edge
				set_default_format_list		get_default_format_list		set_format_string_parser
				get_format_string_parser	get_group_return_type		set_group_return_type
				get_epoch_year				get_shared_string_position	get_format_position
				get_worksheet_names			sheet_count					start_at_the_beginning
				worksheet_count				chartsheet_count			get_chartsheet_names
				get_sheet_names				worksheet_name				chartsheet_name
				in_the_list					get_empty_return_type		set_empty_return_type
			);
###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space => 'main', );
###LogSD		$phone->talk( level => 'info', message => [ "easy questions ..." ] );
map{ 
has_attribute_ok
			'Spreadsheet::XLSX::Reader::LibXML', $_,
										"Check that Spreadsheet::XLSX::Reader::LibXML has the -$_- attribute"
} 			@class_attributes;
map{
can_ok		'Spreadsheet::XLSX::Reader::LibXML', $_,
} 			@class_methods;

###LogSD		$phone->talk( level => 'info', message => [ "harder questions ..." ] );
lives_ok{
			$error_instance	= 	Spreadsheet::XLSX::Reader::LibXML::Error->new(
			###LogSD				log_space	=> 'ErrorInstance',
									should_warn => 0,
								);
			$parser =	Spreadsheet::XLSX::Reader::LibXML->new(
							#~ file_name			=> $test_file,
							count_from_zero		=> 0,
							error_inst			=> $error_instance,
							group_return_type	=> 'value',
							empty_return_type	=> 'undef_string',
			###LogSD		log_space			=> 'Test',
						);
}										"Prep a test parser instance";
###LogSD	$phone->talk( level => 'info', message => [ "parser only loaded" ] );

lives_ok{ 	
			$workbook = $parser->parse( $test_file );
}										"Attempt to unzip the file and prepare to read data";
			#~ print Dumper( $workbook );
			if ( !defined $workbook ) {
				# the test version of "die $parser->error()";
is			$parser->error(), 'Workbook failed to load',
										"Write any error messages from the file load";
			}else{
ok			1,							"The file unzipped and the parser set up without issues";
			}

			my	$offset_ref = [ 0, 8, 15 ];
			my	$y = 0;
			for my $worksheet ( $workbook->worksheets() ) {
explain		'testing worksheet: ' . $worksheet->get_name;
				$row_ref = undef;
			my	$x = 0;
			SHEETDATA: while( $x < 50 and !$row_ref or $row_ref ne 'EOF' ){
###LogSD	if( $x == 0 ){
###LogSD		$operator->add_name_space_bounds( {
###LogSD			Test =>{
###LogSD				Worksheet =>{
###LogSD					UNBLOCK =>{
###LogSD						log_file => 'trace',
###LogSD					},
#~ ###LogSD					_set_file_name =>{
#~ ###LogSD						UNBLOCK =>{
#~ ###LogSD							log_file => 'warn',
#~ ###LogSD						},
#~ ###LogSD					},
#~ ###LogSD					Types =>{
#~ ###LogSD						UNBLOCK =>{
#~ ###LogSD							log_file => 'warn',
#~ ###LogSD						},
#~ ###LogSD					},
###LogSD				},
###LogSD			},
###LogSD		} );
###LogSD	}elsif( $x > 0 ){
###LogSD		exit 1;
###LogSD	}
###LogSD	$phone->talk( level => 'debug', message => [ "getting position: $x" ] );
 
lives_ok{	$row_ref = $worksheet->fetchrow_arrayref }
										'Get the cell value for row: ' . ($x);
#~ explain		"Checking answer position: " . ($offset_ref->[$y] + $x);
			if( !ref $row_ref ){
is			$row_ref, $answer_ref->[$offset_ref->[$y] + $x++],
										"Check for expected EOF";
			last SHEETDATA;
			}else{
is_deeply	$row_ref, $answer_ref->[$offset_ref->[$y] + $x],
										"..and check that the correct values were returned";
			}
			$x++;
			}
			$y++;
			}
is			$workbook->parse( 'badfile.not' ), undef,
										"Check that a bad file will not load";
like		$workbook->error, qr/Attribute \(file_name\) does not pass the type constraint because: The string \-badfile\.not\- does not have an xlsx file extension/,
										"Confirm that the correct error is passed";
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