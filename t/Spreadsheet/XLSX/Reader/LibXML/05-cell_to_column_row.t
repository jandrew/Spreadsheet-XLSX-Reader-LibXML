#########1 Test File for Spreadsheet::XLSX::Reader::LibXML::CellToColumnRow     8#########9
#!env perl
BEGIN{ $ENV{PERL_TYPE_TINY_XS} = 0; };
$| = 1;

use	Test::Most tests => 79;
use	Test::Moose;
use	MooseX::ShortCut::BuildInstance qw( build_instance should_re_use_classes );
should_re_use_classes( 1 );
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
use	Spreadsheet::XLSX::Reader::LibXML::CellToColumnRow;
use	Spreadsheet::XLSX::Reader::LibXML::Error;
use	Spreadsheet::XLSX::Reader::LibXML::LogSpace;
my  ( 
			$test_instance,
	);
my  		@class_attributes = qw(
			);
my  		@class_methods = qw(
				parse_column_row
				build_cell_label
			);
my			$question_ref =[
				'A1', 'B1','C1','D1','E1', 'F1','G1','H1',
				'I1', 'J1','K1','L1','M1', 'N1','O1','P1',
				'Q1', 'R1','S1','T1','U1', 'V1','W1','X1',
				'Y1', 'Z1','AA1','AB1','AC1', 'AD1','AE1',
				'XFD1048576', 'XFE1', 'A1048577', 'A0', '10',
				'A', 'Z1.1',
				];
my			$answer_ref = [
				[ 1, 1 ],[ 2, 1 ],[ 3, 1 ],[ 4, 1 ],[ 5, 1 ],
				[ 6, 1 ],[ 7, 1 ],[ 8, 1 ],[ 9, 1 ],[ 10, 1 ],
				[ 11, 1 ],[ 12, 1 ],[ 13, 1 ],[ 14, 1 ],[ 15, 1 ],
				[ 16, 1 ],[ 17, 1 ],[ 18, 1 ],[ 19, 1 ],[ 20, 1 ],
				[ 21, 1 ],[ 22, 1 ],[ 23, 1 ],[ 24, 1 ],[ 25, 1 ],
				[ 26, 1 ],[ 27, 1 ],[ 28, 1 ],[ 29, 1 ],[ 30, 1 ],
				[ 31, 1 ],[ 16384, 1048576 ],
				[ undef, 1 ], [ 1, undef ], [ 1, undef ],
				[ undef, 10 ], [ 1, undef ], [ undef, undef ],
				[ 0, 0 ],[ 1, 0 ],[ 2, 0 ],[ 3, 0 ],[ 4, 0 ],
				[ 5, 0 ],[ 6, 0 ],[ 7, 0 ],[ 8, 0 ],[ 9, 0 ],
				[ 10, 0 ],[ 11, 0 ],[ 12, 0 ],[ 13, 0 ],[ 14, 0 ],
				[ 15, 0 ],[ 16, 0 ],[ 17, 0 ],[ 18, 0 ],[ 19, 0 ],
				[ 20, 0 ],[ 21, 0 ],[ 22, 0 ],[ 23, 0 ],[ 24, 0 ],
				[ 25, 0 ],[ 26, 0 ],[ 27, 0 ],[ 28, 0 ],[ 29, 0 ],
				[ 30, 0 ],[ 16383, 1048575 ],
				[ undef, 0 ], [ 0, undef ], [ 0, undef ],
				[ undef, 9 ], [ 0, undef ], [ undef, undef ],
			];
my			$error_ref =[
				undef,undef,undef,undef,undef,undef,undef,undef,undef,undef,
				undef,undef,undef,undef,undef,undef,undef,undef,undef,undef,
				undef,undef,undef,undef,undef,undef,undef,undef,undef,undef,
				undef,undef,
				qr/\QThe column text -XFE- points to a position at -16385- past the excel limit of: 16,384\E/,
				qr/\QThe requested row cannot be greater than 1,048,576 - you requested: 1048577\E/,
				qr/\QThe requested row cannot be less than one - you requested: 0\E/,
				qr/\QCould not parse the column component from -10-\E/,
				qr/\QCould not parse the row component from -A-\E/,
				qr/\QThe regex (?^:^([A-Z])?([A-Z])?([A-Z])?([0-9]*)$) could not match -Z1.1-\E/,
				undef,undef,undef,undef,undef,undef,undef,undef,undef,undef,
				undef,undef,undef,undef,undef,undef,undef,undef,undef,undef,
				undef,undef,undef,undef,undef,undef,undef,undef,undef,undef,
				undef,undef,
				qr/\QThe column text -XFE- points to a position at -16385- past the excel limit of: 16,384\E/,
				qr/\QThe requested row cannot be greater than 1,048,576 - you requested: 1048577\E/,
				qr/\QThe requested row cannot be less than one - you requested: 0\E/,
				qr/\QCould not parse the column component from -10-\E/,
				qr/\QCould not parse the row component from -A-\E/,
				qr/\QThe regex (?^:^([A-Z])?([A-Z])?([A-Z])?([0-9]*)$) could not match -Z1.1-\E/,
			];
###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space => 'main', );
###LogSD		$phone->talk( level => 'info', message => [ "initial questions ..." ] );
lives_ok{
			$test_instance = build_instance(
				package => 'Spreadsheet::XLSX::Reader::LibXML::CellToColumnRow::TestClass',
				add_roles_in_sequence =>[ 
					'Spreadsheet::XLSX::Reader::LibXML::LogSpace',
					'Spreadsheet::XLSX::Reader::LibXML::CellToColumnRow',
				],
				add_attributes =>{ 
					error_inst =>{
						handles =>[ qw( error set_error clear_error set_warnings if_warn ) ],
						default	=>	sub{ Spreadsheet::XLSX::Reader::LibXML::Error->new(
										#~ should_warn => 1,
										should_warn => 0,# to turn off cluck when the error is set
									) },
					},
					
				},
				name_space		=> 'Test',
				should_warn		=> 0,
				count_from_zero	=> 0,
			);
}										"Prep a new CellToColumnRow instance";
map{ 
has_attribute_ok
			$test_instance, $_,			"Check that Spreadsheet::XLSX::Reader::LibXML::CellToColumnRow has the -$_- attribute"
} 			@class_attributes;
map{
can_ok		$test_instance, $_,
} 			@class_methods;

###LogSD		$phone->talk( level => 'info', message => [ "harder questions ..." ] );
			no warnings 'uninitialized';
map{
is_deeply	[ $test_instance->parse_column_row( $question_ref->[$_] ) ], $answer_ref->[$_],
										"Convert the Excel cell ID -" . $question_ref->[$_] . "- to column, row: (" .
										$answer_ref->[$_]->[0] . ', ' . $answer_ref->[$_]->[1] . ')';
if( $error_ref->[$_] ){
like		$test_instance->error, $error_ref->[$_],
										"... and check for the correct error message";
}
}(0 .. 37);
map{
is			$test_instance->build_cell_label( @{$answer_ref->[ $_]} ), $question_ref->[$_],#Reverse the polarity flow through the gate
										"Convert the column, row: (" . $answer_ref->[ $_]->[0] . 
										', ' . $answer_ref->[ $_]->[1] . ') - to Excel cell ID -' . $question_ref->[$_] . '-'
										;
}(0 .. 31);
#~ lives_ok{
			#~ $test_instance = build_instance(
				#~ package => 'Spreadsheet::XLSX::Reader::LibXML::CellToColumnRow::TestClass',# Just call the class since it is already built
				#~ error_inst => Spreadsheet::XLSX::Reader::LibXML::Error->new(
					# should_warn => 1,
					#~ should_warn => 0,# to turn off cluck when the error is set
				#~ ),
				#~ name_space		=> 'Test',
				#~ should_warn		=> 0,
				#~ count_from_zero	=> 1,
			#~ );
#~ }										"Build a new test instance to count rows and columns from zero";
#~ map{
#~ is_deeply	[ $test_instance->parse_column_row( $question_ref->[$_] ) ], $answer_ref->[38 + $_],
										#~ "Convert the Excel cell ID -" . $question_ref->[$_] . "- to column, row: (" .
										#~ $answer_ref->[38 + $_]->[0] . ', ' . $answer_ref->[38 + $_]->[1] . ')';
#~ if( $error_ref->[38 + $_] ){
#~ like		$test_instance->error, $error_ref->[38 + $_],
										#~ "... and check for the correct error message";
#~ }
#~ }(0 .. 37);
#~ map{
#~ is			$test_instance->build_cell_label( @{$answer_ref->[38 + $_]} ), $question_ref->[$_],#Reverse the polarity flow through the gate
										#~ "Convert the column, row: (" . $answer_ref->[38 + $_]->[0] . 
										#~ ', ' . $answer_ref->[38 + $_]->[1] . ') - to Excel cell ID -' . $question_ref->[$_] . '-'
										#~ ;
#~ }(0 .. 31);
			use warnings 'uninitialized';
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