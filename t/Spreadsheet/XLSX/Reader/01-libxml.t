#########1 Test File for Spreadsheet::XLSX::Reader          6#########7#########8#########9
#!env perl
BEGIN{ $ENV{PERL_TYPE_TINY_XS} = 0; }
$| = 1;

use	Test::Most;
use	Test::Moose;
use	Data::Dumper;
use	MooseX::ShortCut::BuildInstance v1.26 qw( build_instance );
use	lib
		'../../../../Log-Shiras/lib',
		'../../../lib',;
use Log::Shiras::Switchboard v0.21 qw( :debug );#
###LogSD	my	$operator = Log::Shiras::Switchboard->get_operator(#
###LogSD						name_space_bounds =>{
###LogSD							UNBLOCK =>{
###LogSD								log_file => 'trace',
###LogSD							},
###LogSD							calcChain =>{
###LogSD								UNBLOCK =>{
###LogSD									log_file => 'warn',
###LogSD								},
###LogSD							},
###LogSD							sharedStrings =>{
###LogSD								UNBLOCK =>{
###LogSD									log_file => 'warn',
###LogSD								},
###LogSD							},
###LogSD							styles =>{
###LogSD								UNBLOCK =>{
###LogSD									log_file => 'warn',
###LogSD								},
###LogSD							},
###LogSD							Test =>{
###LogSD								Worksheet =>{
###LogSD									UNBLOCK =>{
###LogSD										log_file => 'trace',
###LogSD									},
###LogSD									_set_file_name =>{
###LogSD										UNBLOCK =>{
###LogSD											log_file => 'warn',
###LogSD										},
###LogSD									},
###LogSD									_load_unique_bits =>{
###LogSD										UNBLOCK =>{
###LogSD											log_file => 'trace',
###LogSD										},
###LogSD									},
###LogSD									parse_column_row =>{
###LogSD										UNBLOCK =>{
###LogSD											log_file => 'warn',
###LogSD										},
###LogSD									},
###LogSD								},
###LogSD								Workbook =>{
###LogSD									worksheets =>{
###LogSD										UNBLOCK =>{
###LogSD											log_file => 'warn',
###LogSD										},
###LogSD									},
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
###LogSD									_load_workbook_file =>{
###LogSD										UNBLOCK =>{
###LogSD											log_file => 'warn',
###LogSD										},
###LogSD									},
###LogSD									_load_rels_workbook_file =>{
###LogSD										UNBLOCK =>{
###LogSD											log_file => 'warn',
###LogSD										},
###LogSD									},
###LogSD									_load_doc_props_file =>{
###LogSD										UNBLOCK =>{
###LogSD											log_file => 'warn',
###LogSD										},
###LogSD									},
###LogSD									_set_shared_worksheet_files =>{
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
use Spreadsheet::XLSX::Reader;
my	$test_file = ( @ARGV ) ? $ARGV[0] : '../../test_files/';
	$test_file .= 'TestBook.xlsx';
	#~ print "Test file is: $test_file\n";
my  ( 
			$test_instance, $capture, $x, @answer, $error_instance,
	);
my 			$row = 0;
my 			@class_attributes = qw(
				file_name
				creator
				modified_by
				date_created
				date_modified
				sheet_parser
				sheet_parser_modules
			);
my  		@class_methods = qw(
				new
				parse
				worksheet
				worksheets
				set_file_name
				has_file_name
				set_sheet_parser_modules
				get_sheet_parser_modules
				get_epoch_year
				get_shared_string_position
				get_calc_chain_position
				get_worksheet_names
				worksheet_name
				number_of_sheets
				start_at_the_beginning
			);
###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space => 'main', );
###LogSD		$phone->talk( level => 'info', message => [ "easy questions ..." ] );
map{ 
has_attribute_ok
			'Spreadsheet::XLSX::Reader', $_,
										"Check that Spreadsheet::XLSX::Reader has the -$_- attribute"
} 			@class_attributes;
map{
can_ok		'Spreadsheet::XLSX::Reader', $_,
} 			@class_methods;


my	$parser =	Spreadsheet::XLSX::Reader->new(
					#~ file_name 	=> $test_file,
					count_from_zero	=> 0,
					log_space		=> 'Test',
					should_warn		=> 1,
				);
###LogSD	$phone->talk( level => 'trace', message => [ "parser only loaded" ] );
my	$workbook = $parser->parse( $test_file );
 
if ( !defined $workbook ) {
    #~ die $parser->error();
	die $parser->error()
}
 
for my $worksheet ( $workbook->worksheets() ) {
	print '-----  ' . $worksheet->name . "\n";
    my ( $row_min, $row_max ) = $worksheet->row_range();
    my ( $col_min, $col_max ) = $worksheet->col_range();
 
    for my $row ( $row_min .. $row_max ) {
        for my $col ( $col_min .. $col_max ) {
 
            my $cell = $worksheet->get_cell( $row, $col );
            next unless $cell;
 
            print "Row, Col    = ($row, $col)\n";
            print "Value       = ", $cell->value(),       "\n";
            print "Unformatted = ", $cell->unformatted(), "\n";
            print "\n";
			#~ my	$wait = <>;
        }
    }
}


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
###LogSD			$line =~ s/\n/\n\t\t/g;
###LogSD			push @print_list, $line;
###LogSD		}
###LogSD		printf( "name_space - %-50s | level - %-6s |\nfile_name  - %-50s | line  - %04d   |\n\t:(\t%s ):\n", 
###LogSD					$_[0]->{name_space}, $_[0]->{level},
###LogSD					$_[0]->{filename}, $_[0]->{line},
###LogSD					join( "\n\t\t", @print_list ) 	);
###LogSD		use warnings 'uninitialized';
###LogSD	}

1;