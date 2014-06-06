#########1 Test File for Spreadsheet::XLSX::Reader::XMLDOM::Cell      7#########8#########9
#!perl
$| = 1;

use	Test::Most;
use	Test::Moose;
use	Data::Dumper;
use	MooseX::ShortCut::BuildInstance v1.8 qw( build_instance );
use	lib
		'../../../../../../Log-Shiras/lib',
		'../../../../../lib',;
#~ use Log::Shiras::Switchboard qw( :debug );#
###LogSD	my	$operator = Log::Shiras::Switchboard->get_operator(#
###LogSD						name_space_bounds =>{
###LogSD							UNBLOCK =>{
###LogSD								log_file => 'trace',
###LogSD							},
###LogSD						},
###LogSD						reports =>{
###LogSD							log_file =>[ Print::Log->new ],
###LogSD						},
###LogSD					);
###LogSD	use Log::Shiras::Telephone;
###LogSD	use Log::Shiras::UnhideDebug;
use	Spreadsheet::XLSX::Reader::XMLDOM::Cell;
use	Spreadsheet::XLSX::Reader::Error;
use	Spreadsheet::XLSX::Reader::Types qw( ZeroFromNum FourteenFromWinExcelNum );
my	$test_file = ( @ARGV ) ? $ARGV[0] : '../../../../test_files/';
	$test_file .= 'styles.xml';
###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space => 'main', );
###LogSD		$phone->talk( level => 'trace', message => [ "Test file is: $test_file" ] );
my  ( 
			$test_instance, $capture, $x, @answer,# $error_instance,
	);
my 			$row = 0;
my 			@class_attributes = qw(
				error_inst
				value_encoding
				value_type
				cell_column
				cell_row
				number_format
				cell_element
			);
				#~ _string_encoding
				#~ _cell_element
				#~ _stored_type
				#~ _cell_column
				#~ _cell_row
				#~ _format
my  		@class_methods = qw(
				new
				formula
				is_merged
				encoding
				get_xml
				get_cell_position
				value
				unformatted
				type
				column
				row
				set_format
				get_format
				clear_format
				has_format
				format_name
				error
				set_error
				clear_error
				set_warnings
				if_warn
			);
###LogSD		$phone->talk( level => 'info', message => [ "easy questions ..." ] );
map{ 
has_attribute_ok
			'Spreadsheet::XLSX::Reader::XMLDOM::Cell', $_,
										"Check that Spreadsheet::XLSX::Reader::XMLDOM::Cell has the -$_- attribute"
} 			@class_attributes;
map{
can_ok		'Spreadsheet::XLSX::Reader::XMLDOM::Cell', $_,
} 			@class_methods;

###LogSD		$phone->talk( level => 'info', message => [ "harder questions ..." ] );
lives_ok{
			$test_instance	=	Spreadsheet::XLSX::Reader::XMLDOM::Cell->new(
									error_inst		=> Spreadsheet::XLSX::Reader::Error->new,
									value_encoding	=> 'UTF-8',
									value_type		=> 'number',
									cell_column		=> 1,
									cell_row		=> 1,
									number_format	=> ZeroFromNum,
									#~ cell_element		=> 
								);
}										"Prep a new Cell instance";

###LogSD		$phone->talk( level => 'info', message => [ "hardest questions ..." ] );
#~ is			$test_instance->formula, '????',
										#~ "Check that the 'formula' method returns: ????";
#~ is_deeply	$test_instance->is_merged, ['????', '????'],4
										#~ "Check that the 'is_merged' method returns: [????, ????]";
#~ is_deeply	$test_instance->is_merged( 'array' ), [['?','?'] ['?','?']],
										#~ "Check that the 'is_merged( 'array' )' method returns: [['?','?'] ['?','?']]";
is			$test_instance->encoding, 'UTF-8',
										"Check that the 'encoding' method returns: UTF-8";
#~ is			$test_instance->get_xml->textContent, '?????',
										#~ "Pull the whole cell as an XML node, test a node method, and check the result: ????";
#~ is			$test_instance->get_cell_position, '????',
										#~ "Check that the 'get_cell_position' method returns: ????";
#~ is			$test_instance->unformatted, '????',
										#~ "Check that the 'unformatted' method returns: ????";
is			$test_instance->type, 'number',
										"Check that the 'type' method returns: number";
is			$test_instance->column, 1,
										"Check that the 'column' method returns: 1";
is			$test_instance->row, 1,
										"Check that the 'row' method returns: 1";
is			$test_instance->has_format, 1,
										"Check that the 'has_format' method returns: TRUE";
is			$test_instance->format_name, 'ZeroFromNum',
										"Check that the 'format_name' method returns: ZeroFromNum";
is			$test_instance->get_format->display_name, 'ZeroFromNum',
										"Get the full Type::Coercion object and call a Type::Coercion method on it returning: ZeroFromNum";
lives_ok{	$test_instance->clear_format }
										"Clear the format";
is			$test_instance->has_format, '',
										"... and check that the 'has_format' method returns: FALSE";
is			$test_instance->set_format( FourteenFromWinExcelNum ), FourteenFromWinExcelNum,
										"Set the format object to: FourteenFromWinExcelNum";
is			$test_instance->format_name, 'FourteenFromWinExcelNum',
										"... and check that the 'format_name' method returns: FourteenFromWinExcelNum";
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
###LogSD			$line =~ s/\n/\n\t\t/g;
###LogSD			push @print_list, $line;
###LogSD		}
###LogSD		printf( "name_space - %-50s | level - %-6s |\nfile_name  - %-50s | line  - %04d   |\n\t:(\t%s ):\n", 
###LogSD					$_[0]->{name_space}, $_[0]->{level},
###LogSD					$_[0]->{filename}, $_[0]->{line},
###LogSD					join( "\n\t\t", @print_list ) 	);
###LogSD		use warnings 'uninitialized';
###LogSD	}

###LogSD	1;