#########1 Test File for Spreadsheet::XLSX::Reader          6#########7#########8#########9
#!perl
$| = 1;

use	Test::Most;
use	Test::Moose;
use	Data::Dumper;
use	MooseX::ShortCut::BuildInstance qw( build_instance );
use	lib
		'../../../../Log-Shiras/lib',
		'../../../lib',;
use Log::Shiras::Switchboard qw( :debug );#
my	$operator = Log::Shiras::Switchboard->get_operator(#
					name_space_bounds =>{
						UNBLOCK =>{
							log_file => 'trace',
						},
						Test =>{
							Workbook =>{
								_build_file =>{
									UNBLOCK =>{
										log_file => 'debug',
									},
								},
							},
						},
					},
					reports =>{
						log_file =>[ Print::Log->new ],
					},
				);
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
### <where> - easy questions ...
map{ 
has_attribute_ok
			'Spreadsheet::XLSX::Reader', $_,
										"Check that Spreadsheet::XLSX::Reader has the -$_- attribute"
} 			@class_attributes;
map{
can_ok		'Spreadsheet::XLSX::Reader', $_,
} 			@class_methods;







my	$workbook =	Spreadsheet::XLSX::Reader->new(
					file_name 	=> $test_file,
					log_space	=> 'Test'
				);
 
if ( !defined $workbook ) {
    #~ die $parser->error();
	die $workbook->error()
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
            #~ print "Value       = ", $cell->value(),       "\n";
            print "Unformatted = ", $cell->unformatted(), "\n";
            print "\n";
			#~ my	$wait = <>;
        }
    }
}


package Print::Log;
use Data::Dumper;
sub new{
	bless {}, shift;
}
sub add_line{
	shift;
	#~ print Dumper( $_[0] );
	#~ exit 1;
	my @input = ( ref $_[0]->{message} eq 'ARRAY' ) ? 
					@{$_[0]->{message}} : $_[0]->{message};
	my ( @print_list, @initial_list );
	no warnings 'uninitialized';
	for my $value ( @input ){
		push @initial_list, (( ref $value ) ? Dumper( $value ) : $value );
	}
	for my $line ( @initial_list ){
		$line =~ s/\n/\n\t\t/g;
		push @print_list, $line;
	}
	printf( "name_space - %-50s | level - %-6s |\nfile_name  - %-50s | line  - %04d   |\n\t:(\t%s ):\n", 
				$_[0]->{name_space}, $_[0]->{level},
				$_[0]->{filename}, $_[0]->{line},
				join( "\n\t\t", @print_list ) 	);
	use warnings 'uninitialized';
}

1;