#########1 Test File for Spreadsheet::XLSX::Reader::LibXML::GetCell   7#########8#########9
#!evn perl
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

use	Test::Most tests => 729;
use	Test::Moose;
use	DateTime::Format::Flexible;
use	DateTimeX::Format::Excel;
use Type::Tiny;
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
###LogSD						reports =>{
###LogSD							log_file =>[ Print::Log->new ],
###LogSD						},
###LogSD					);
###LogSD	use Log::Shiras::Telephone;
###LogSD	use Log::Shiras::UnhideDebug;
use	Spreadsheet::XLSX::Reader::LibXML::Error;
###LogSD	use Log::Shiras::UnhideDebug;
use	Spreadsheet::XLSX::Reader::LibXML::XMLReader::SharedStrings v0.5;
use	Spreadsheet::XLSX::Reader::LibXML::FmtDefault v0.5;
use	Spreadsheet::XLSX::Reader::LibXML::ParseExcelFormatStrings v0.5;
###LogSD	use Log::Shiras::UnhideDebug;
use	Spreadsheet::XLSX::Reader::LibXML::XMLReader::Styles v0.5;
use	Spreadsheet::XLSX::Reader::LibXML::GetCell v0.5;
###LogSD	use Log::Shiras::UnhideDebug;
use	Spreadsheet::XLSX::Reader::LibXML::XMLReader::Worksheet v0.5;
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
	);
my 			$row = 0;
my 			@class_attributes = qw(
				file_name					error_inst					log_space
				sheet_rel_id				sheet_id					sheet_position
				sheet_name					custom_formats
			);
my  		@class_methods = qw(
				new							counting_from_zero			boundary_flag_setting
				change_boundary_flag		_has_shared_strings_file	get_shared_string_position
				_has_styles_file			get_format_position			get_cell
				get_next_value				fetchrow_arrayref			fetchrow_array
				set_headers					fetchrow_hashref			has_custom_format
				get_custom_format			set_custom_format			set_custom_formats
			);
my			$answer_list =[
				{},{},{},{},{},{},
				{ cell_id => 'A2', row => 1, col => 0, type => 'Text', unformatted => 'Hello', value => 'Hello' },{},{},
				{ cell_id => 'D2', row => 1, col => 3, type => 'Text', unformatted => 'my', value => 'my' },{},{},
				{},{},{},{},{},{},
				{},{},{ cell_id => 'C4', row => 3, col => 2, type => 'Text', unformatted => 'World', value => 'World', coercion_name => 'Excel_number_0' },{},{},{},
				{},{},{},{},{},{},
				{ cell_id => 'A6', row => 5, col => 0, type => 'Text', unformatted => 'Hello World', value => 'Hello World',
					get_rich_text =>[
						2, { color =>{ rgb => 'FFFF0000' }, sz => '11', b => 1, scheme => 'minor', rFont => 'Calibri', family => 2 },
						6, { color =>{ rgb => 'FF0070C0' }, sz => '20', b => 1, scheme => 'minor', rFont => 'Calibri', family => 2 },
					],
					merge_range => 'A6:B6',
				},
				{ cell_id => 'B6', row => 5, col => 1, type => 'Text', unformatted => undef, value => undef, merge_range => 'A6:B6' },{},{},{},{},
				{},{ cell_id => 'B7', row => 6, col => 1, type => 'Numeric', unformatted => '69', value => '69' },{},{},{},{},
				{},{ cell_id => 'B8', row => 7, col => 1, type => 'Numeric', unformatted => '27', value => '27' },{},{},
				{ cell_id => 'E8', row => 7, col => 4, type => 'Date', unformatted => '37145', value => '12-Sep-05', coercion_name => 'Excel_date_164' },{},
				{},{ cell_id => 'B9', row => 8, col => 1, type => 'Numeric', unformatted => '42', value => '42', formula => 'B7-B8' },{},{},{},{},
				{},{},{},{ cell_id => 'D10', row => 9, col => 3, type => 'Custom', unformatted => undef, value => undef, coercion_name => 'YYYYMMDD' },
				{ cell_id => 'E10', row => 9, col => 4, type => 'Custom', unformatted => '2/6/2011', value => '2011-02-06T00:00:00', coercion_name => 'Custom_date_type' },
				{ cell_id => 'F10', row => 9, col => 5, type => 'Custom', unformatted => '2/6/2011', value => '2011-02-06', coercion_name => 'YYYYMMDD' },
				{ cell_id => 'A11', row => 10, col => 0, type => 'Numeric', unformatted => '2.1345678901', value => '2.13', coercion_name => 'Excel_number_2' },{},{},{},{},{},
				{},{},{},{ cell_id => 'D12', row => 11, col => 3, type => 'Date', unformatted => '39118', value => '6-Feb-11', coercion_name => 'Excel_date_164', merge_range => 'D12:E12' },
				{ cell_id => 'E12', row => 11, col => 4, type => 'Text', unformatted => undef, value => undef, coercion_name => 'Excel_date_164', merge_range => 'D12:E12' },{},
				{},{},{},{},{},{},
				{},{},{ cell_id => 'C14', row => 13, col => 2, type => 'Text', unformatted => ' ', value => ' ', has_coercion => '', },
				{ cell_id => 'D14', row => 13, col => 3, type => 'Custom', unformatted => '39118', value => '2011-02-06', coercion_name => 'YYYYMMDD', },
				{ cell_id => 'E14', row => 13, col => 4, type => 'Date', unformatted => '39118', value => '6-Feb-11', coercion_name => 'Excel_date_164', },{},
				'EOF',
				{ cell_id => 'A2', row => 1, col => 0, type => 'Text', unformatted => 'Hello', value => 'Hello' },
				{ cell_id => 'D2', row => 1, col => 3, type => 'Text', unformatted => 'my', value => 'my' },
				{ cell_id => 'C4', row => 3, col => 2, type => 'Text', unformatted => 'World', value => 'World', coercion_name => 'Excel_number_0' },
				{ cell_id => 'A6', row => 5, col => 0, type => 'Text', unformatted => 'Hello World', value => 'Hello World',
					get_rich_text =>[
						2, { color =>{ rgb => 'FFFF0000' }, sz => '11', b => 1, scheme => 'minor', rFont => 'Calibri', family => 2 },
						6, { color =>{ rgb => 'FF0070C0' }, sz => '20', b => 1, scheme => 'minor', rFont => 'Calibri', family => 2 },
					],
					merge_range => 'A6:B6',
				},
				{ cell_id => 'B6', row => 5, col => 1, type => 'Text', unformatted => undef, value => undef, merge_range => 'A6:B6' },
				{ cell_id => 'B7', row => 6, col => 1, type => 'Numeric', unformatted => '69', value => '69' },
				{ cell_id => 'B8', row => 7, col => 1, type => 'Numeric', unformatted => '27', value => '27' },
				{ cell_id => 'E8', row => 7, col => 4, type => 'Date', unformatted => '37145', value => '12-Sep-05', coercion_name => 'Excel_date_164' },
				{ cell_id => 'B9', row => 8, col => 1, type => 'Numeric', unformatted => '42', value => '42', formula => 'B7-B8' },
				{ cell_id => 'D10', row => 9, col => 3, type => 'Custom', unformatted => undef, value => undef, coercion_name => 'YYYYMMDD' },
				{ cell_id => 'E10', row => 9, col => 4, type => 'Custom', unformatted => '2/6/2011', value => '2011-02-06T00:00:00', coercion_name => 'Custom_date_type' },
				{ cell_id => 'F10', row => 9, col => 5, type => 'Custom', unformatted => '2/6/2011', value => '2011-02-06', coercion_name => 'YYYYMMDD' },
				{ cell_id => 'A11', row => 10, col => 0, type => 'Numeric', unformatted => '2.1345678901', value => '2.13', coercion_name => 'Excel_number_2' },
				{ cell_id => 'D12', row => 11, col => 3, type => 'Date', unformatted => '39118', value => '6-Feb-11', coercion_name => 'Excel_date_164', merge_range => 'D12:E12' },
				{ cell_id => 'E12', row => 11, col => 4, type => 'Text', unformatted => undef, value => undef, coercion_name => 'Excel_date_164', merge_range => 'D12:E12' },
				{ cell_id => 'C14', row => 13, col => 2, type => 'Text', unformatted => ' ', value => ' ', has_coercion => '', },
				{ cell_id => 'D14', row => 13, col => 3, type => 'Custom', unformatted => '39118', value => '2011-02-06', coercion_name => 'YYYYMMDD', },
				{ cell_id => 'E14', row => 13, col => 4, type => 'Date', unformatted => '39118', value => '6-Feb-11', coercion_name => 'Excel_date_164', },
				undef,
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
					{},{},{ cell_id => 'C4', row => 3, col => 2, type => 'Text', unformatted => 'World', value => 'World', coercion_name => 'Excel_number_0' },{},{},{},
				],
				[
					{},{},{},{},{},{},
				],
				[
					{ cell_id => 'A6', row => 5, col => 0, type => 'Text', unformatted => 'Hello World', value => 'Hello World',
						get_rich_text =>[
							2, { color =>{ rgb => 'FFFF0000' }, sz => '11', b => 1, scheme => 'minor', rFont => 'Calibri', family => 2 },
							6, { color =>{ rgb => 'FF0070C0' }, sz => '20', b => 1, scheme => 'minor', rFont => 'Calibri', family => 2 },
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
					{},{},{},{ cell_id => 'D10', row => 9, col => 3, type => 'Custom', unformatted => undef, value => undef, coercion_name => 'YYYYMMDD' },
					{ cell_id => 'E10', row => 9, col => 4, type => 'Custom', unformatted => '2/6/2011', value => '2011-02-06T00:00:00', coercion_name => 'Custom_date_type' },
					{ cell_id => 'F10', row => 9, col => 5, type => 'Custom', unformatted => '2/6/2011', value => '2011-02-06', coercion_name => 'YYYYMMDD' },
				],
				[
					{ cell_id => 'A11', row => 10, col => 0, type => 'Numeric', unformatted => '2.1345678901', value => '2.13', coercion_name => 'Excel_number_2' },{},{},{},{},{},
				],
				[
					{},{},{},{ cell_id => 'D12', row => 11, col => 3, type => 'Date', unformatted => '39118', value => '6-Feb-11', coercion_name => 'Excel_date_164', merge_range => 'D12:E12' },
					{ cell_id => 'E12', row => 11, col => 4, type => 'Text', unformatted => undef, value => undef, coercion_name => 'Excel_date_164', merge_range => 'D12:E12' },{},
				],
				[
					{},{},{},{},{},{},
				],
				[
					{},{},{ cell_id => 'C14', row => 13, col => 2, type => 'Text', unformatted => ' ', value => ' ', has_coercion => '', },
					{ cell_id => 'D14', row => 13, col => 3, type => 'Custom', unformatted => '39118', value => '2011-02-06', coercion_name => 'YYYYMMDD', },
					{ cell_id => 'E14', row => 13, col => 4, type => 'Date', unformatted => '39118', value => '6-Feb-11', coercion_name => 'Excel_date_164', },{},
				],
				undef,
				[undef,undef,undef,undef,undef,undef,],
				['Hello',undef,undef,'my',undef,undef,],
				[undef,undef,undef,undef,undef,undef,],
				[undef,undef,'World',undef,undef,undef,],
				[undef,undef,undef,undef,undef,undef,],
				['Hello World',undef,undef,undef,undef,undef,],
				[undef,'69',undef,undef,undef,undef,],
				[undef,'27',undef,undef,'37145',undef,],
				[undef,'42',,undef,undef,undef,undef,],
				[undef,undef,undef,undef,'2/6/2011','2/6/2011',],
				['2.1345678901',undef,undef,undef,undef,undef,],
				[undef,undef,undef,'39118',undef,undef,],
				[undef,undef,undef,undef,undef,undef,],
				[undef,undef,' ','39118','39118',undef,],
				'EOF',
				['Row Labels', '2016-2-6',  '2017-2-14', '2018-2-3', 'Grand Total', ],
				{ 'Row Labels' => 'Blue', '2016-2-6' => '10', '2017-2-14' => '7', },
				{ 'Row Labels' => 'Omaha', '2018-2-3' => '2', },
				{ 'Row Labels' => 'Red', '2016-2-6' => '30', '2017-2-14' => '5', '2018-2-3' => '3', },
				{ 'Row Labels' => 'Grand Total', '2016-2-6' => '40', '2017-2-14' => '12', '2018-2-3' => '5', },
				'EOF',
				
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
					name		=> 'YYYYMMDD',
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
								);
			$styles_instance =	build_instance(
									package => 'StylesInstance',
									superclasses	=> [ 'Spreadsheet::XLSX::Reader::LibXML::XMLReader::Styles' ],
									add_roles_in_sequence => [qw(
										Spreadsheet::XLSX::Reader::LibXML::FmtDefault
										Spreadsheet::XLSX::Reader::LibXML::ParseExcelFormatStrings
									)],
									file_name	=> $styles_file,
									log_space	=> 'Test::Styles',
									error_inst	=> $error_instance,
									epoch_year	=> 1904,
								);
			$shared_strings_instance =	Spreadsheet::XLSX::Reader::LibXML::XMLReader::SharedStrings->new(
											error_inst	=> $error_instance,
											file_name	=> $shared_strings_file,
											log_space	=> 'Test::SharedStrings',
										);
			$workbook_instance =	build_instance(
										package	=> 'WorkbookInstance',
										add_methods =>{
											get_epoch_year => sub{ return '1904' },
										},
										add_attributes =>{
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
												handles		=>[ 'get_format_position' ],
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
										},
										styles_instance => $styles_instance,
										shared_strings_instance => $shared_strings_instance,
									);
			$test_instance	=	build_instance(
									package	=> 'GetCellTest',
									superclasses =>[ 'Spreadsheet::XLSX::Reader::LibXML::XMLReader::Worksheet' ],
									file_name			=> $test_file,
									log_space			=> 'Test',
									error_inst			=> $error_instance,
									custom_formats		=> {
																E10	=> $date_time_type,
																10	=> $string_type,
																D14	=> $string_type,
															},
									workbook_instance	=> $workbook_instance,
								);
}										"Prep a test GetCellTest instance";
map{ 
has_attribute_ok
			$test_instance, $_,
										"Check that the GetCellTest instance has the -$_- attribute"
} 			@class_attributes;
map{
can_ok		$test_instance, $_,
} 			@class_methods;

###LogSD		$phone->talk( level => 'trace', message => [ 'Test instance:', $test_instance ] );
###LogSD		$phone->talk( level => 'info', message => [ "hardest questions ..." ] );

explain									"Test get_cell";
			my ( $row_min, $row_max ) = $test_instance->row_range();
			my ( $col_min, $col_max ) = $test_instance->col_range();
			my $x = 0;
			no warnings 'uninitialized';
			INITIALRUN: for my $row ( $row_min .. ($row_max + 1) ) {
            for my $col ( $col_min .. $col_max ) {

#~ ###LogSD	if( $row == 3 and $col == 2 ){
#~ ###LogSD		$operator->add_name_space_bounds( {
#~ ###LogSD			main =>{
#~ ###LogSD				UNBLOCK =>{
#~ ###LogSD					log_file => 'debug',
#~ ###LogSD				},
#~ ###LogSD			},
#~ ###LogSD			Test =>{
#~ ###LogSD				UNBLOCK =>{
#~ ###LogSD					log_file => 'trace',
#~ ###LogSD				},
#~ ###LogSD			},
#~ ###LogSD		} );
#~ ###LogSD	}

lives_ok{	$cell = $test_instance->get_cell( $row, $col ) }
										"Get anything at the cell for row -$row- and col -$col-";
###LogSD	$phone->talk( level => 'debug', message => [ "cell:", $cell ] );
			if( !ref $answer_list->[$x] ){
is			$cell, $answer_list->[$x],
										"Check for the correct end of file flag: EOF";
			last INITIALRUN;
			}elsif( scalar( keys %{$answer_list->[$x]} ) == 0 ){
is			!$cell, 1,					"Check that an expected empty cell really is empty";
			}else{
			for my $key ( keys %{$answer_list->[$x]} ){
###LogSD	$phone->talk( level => 'debug', message => [ "checking method:", $key ] );
is_deeply	$cell->$key, $answer_list->[$x]->{$key},
										"Checking the method -$key- value at row -$row- and column -$col- is: $answer_list->[$x]->{$key}";
			}
			}
			$x++;
            }
			}
is			$test_instance->get_cell( 1, 6 ), 'EOR',
										"Get a correct end of row flag: EOR";
is			$workbook_instance->change_boundary_flag( 0 ), 0,
										"Turn boundary flags off";
is			$test_instance->get_cell( 2, 6 ), undef,
										"Check an end of row position (should be undef)";
is			$test_instance->get_cell( 14, 0 ), undef,
										"Check and end of file position (should be undef)";

explain									"Test get_next_value";
			$cell = undef;
			$x = 0;
			VALUERUN: while( $x < 103 and (!$cell or ref $cell eq 'Spreadsheet::XLSX::Reader::LibXML::Cell' ) ){
				my $position = $x + 85;

#~ ###LogSD	if( $position == 85 ){
#~ ###LogSD		$operator->add_name_space_bounds( {
#~ ###LogSD			main =>{
#~ ###LogSD				UNBLOCK =>{
#~ ###LogSD					log_file => 'debug',
#~ ###LogSD				},
#~ ###LogSD			},
#~ ###LogSD			Test =>{
#~ ###LogSD				UNBLOCK =>{
#~ ###LogSD					log_file => 'trace',
#~ ###LogSD				},
#~ ###LogSD			},
#~ ###LogSD		} );
#~ ###LogSD	}

lives_ok{	$cell = $test_instance->get_next_value }
										"Get the next cell with a value for position: $x";
###LogSD	$phone->talk( level => 'debug', message => [ "cell:", $cell ] );
			if( !ref $answer_list->[$position] ){
is			$cell, $answer_list->[$position],
										"Check for the correct end of file flag: undef";
			last VALUERUN;
			}else{
			for my $key ( keys %{$answer_list->[$position]} ){
is_deeply	$cell->$key, $answer_list->[$position]->{$key},
										"Checking the method -$key- value at position -$x- is: $answer_list->[$position]->{$key}";
			}
			}
			$x++;
            }

explain									"Test fetchrow_arrayref";
			$row_ref = undef;
			$x = 0;
			$offset = 104;
			ROWREFRUN: while( $x < 119 and ( !$row_ref or ref $row_ref eq 'ARRAY' ) ){
				my $position = $x + $offset;

#~ ###LogSD	if( $x == 1 ){
#~ ###LogSD		$operator->add_name_space_bounds( {
#~ ###LogSD			main =>{
#~ ###LogSD				UNBLOCK =>{
#~ ###LogSD					log_file => 'debug',
#~ ###LogSD				},
#~ ###LogSD			},
#~ ###LogSD			Test =>{
#~ ###LogSD				UNBLOCK =>{
#~ ###LogSD					log_file => 'trace',
#~ ###LogSD				},
#~ ###LogSD				parse_element =>{
#~ ###LogSD					UNBLOCK =>{
#~ ###LogSD						log_file => 'warn',
#~ ###LogSD					},
#~ ###LogSD				},
#~ ###LogSD				Styles =>{
#~ ###LogSD					UNBLOCK =>{
#~ ###LogSD						log_file => 'warn',
#~ ###LogSD					},
#~ ###LogSD				},
#~ ###LogSD			},
#~ ###LogSD		} );
#~ ###LogSD	}

lives_ok{	$row_ref = $test_instance->fetchrow_arrayref }
										"Get the next fetchrow_arrayref for row: $x";
###LogSD	$phone->talk( level => 'debug', message => [ "row:", $row_ref ] );
			if( !ref $answer_list->[$position] ){###LogSD	$phone->talk( level => 'debug', message => [ "row:", $row ] );
###LogSD	$phone->talk( level => 'debug', message => [ "Found and -end- flag: $answer_list->[$position]" ] );
is			$row_ref, $answer_list->[$position],
										"Check for the correct end of file flag: undef";
			last ROWREFRUN;
			}else{
			my $col = 0;
			for my $cell ( @{$answer_list->[$position]} ){
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

explain									"Test fetchrow_array";
ok			$test_instance->change_boundary_flag( 1 ),
										"Turn boundary flag reporting back on";
ok			$workbook_instance->set_group_return_type( 'unformatted' ),
										"Return just the cell coerced values rather than a Cell instance";
			$x = 0;
			$offset = 119;
			ROWARRAYRUN: while( $x < 134 and ( !$row_ref or !$row_ref->[0] or $row_ref->[0] ne 'EOF' ) ){
				my $position = $x + $offset;

#~ ###LogSD	if( $x == 1 ){
#~ ###LogSD		$operator->add_name_space_bounds( {
#~ ###LogSD			main =>{
#~ ###LogSD				UNBLOCK =>{
#~ ###LogSD					log_file => 'debug',
#~ ###LogSD				},
#~ ###LogSD			},
#~ ###LogSD			Test =>{
#~ ###LogSD				UNBLOCK =>{
#~ ###LogSD					log_file => 'trace',
#~ ###LogSD				},
#~ ###LogSD				parse_element =>{
#~ ###LogSD					UNBLOCK =>{
#~ ###LogSD						log_file => 'warn',
#~ ###LogSD					},
#~ ###LogSD				},
#~ ###LogSD				Styles =>{
#~ ###LogSD					UNBLOCK =>{
#~ ###LogSD						log_file => 'warn',
#~ ###LogSD					},
#~ ###LogSD				},
#~ ###LogSD				SharedStrings =>{
#~ ###LogSD					UNBLOCK =>{
#~ ###LogSD						log_file => 'warn',
#~ ###LogSD					},
#~ ###LogSD				},
#~ ###LogSD			},
#~ ###LogSD		} );
#~ ###LogSD	}

lives_ok{	$row_ref = [$test_instance->fetchrow_array] }
										"Get the next fetchrow_array for row: $x";
###LogSD	$phone->talk( level => 'debug', message => [ "row:", $row_ref ] );
			if( !ref $answer_list->[$position] ){
###LogSD	$phone->talk( level => 'debug', message => [ "Found and -end- flag: $answer_list->[$position]" ] );
is			$row_ref->[0], $answer_list->[$position],
										"Check for the correct end of file flag: EOF";
			last ROWARRAYRUN;
			}else{
is_deeply	$row_ref, $answer_list->[$position],
										"..and validate the returned values";
			}
			$x++;
            }

explain									"Test fetchrow_hashref";
ok			$workbook_instance->set_group_return_type( 'value' ),
										"Set the group_return_type to: value";
ok			$test_instance = GetCellTest->new(
								file_name			=> $test_file_2,
								log_space			=> 'Test',
								error_inst			=> $error_instance,
								custom_formats		=> {
															E10	=> $date_time_type,
															10	=> $string_type,
															D14	=> $string_type,
														},
								workbook_instance	=> $workbook_instance,
							),			'Build another connection to a different worksheet';
is 			$test_instance->fetchrow_hashref( 1 ), undef,
										"Check that a fetchrow_hashref call returns undef without a set header";
is			$test_instance->error, "Headers must be set prior to calling fetchrow_hashref",
										"..and check for the correct error message";
is_deeply	$test_instance->set_headers( 1 ), $answer_list->[134],
										"Set the headers for building a hashref";
ok			$test_instance->set_max_header_col( 3 ),,
										"Set the maximum header column";
			$row_ref = undef;
			$x = 0;
			$offset = 135;
			HASHREFRUN: while( $x < 140 and ( !$row_ref or ref $row_ref eq 'HASH' ) ){
				my $position = $x + $offset;

###LogSD	if( $x == 0 ){
###LogSD		$operator->add_name_space_bounds( {
###LogSD			main =>{
###LogSD				UNBLOCK =>{
###LogSD					log_file => 'debug',
###LogSD				},
###LogSD			},
###LogSD			Test =>{
###LogSD				UNBLOCK =>{
###LogSD					log_file => 'trace',
###LogSD				},
###LogSD				SharedStrings =>{
###LogSD					UNBLOCK =>{
###LogSD						log_file => 'warn',
###LogSD					},
###LogSD				},
###LogSD				Styles =>{
###LogSD					UNBLOCK =>{
###LogSD						log_file => 'warn',
###LogSD					},
###LogSD				},
###LogSD				parse_element =>{
###LogSD					UNBLOCK =>{
###LogSD						log_file => 'warn',
###LogSD					},
###LogSD				},
###LogSD				_get_next_value_cell =>{
###LogSD					UNBLOCK =>{
###LogSD						log_file => 'warn',
###LogSD					},
###LogSD				},
###LogSD				fetchrow_arrayref =>{
###LogSD					UNBLOCK =>{
###LogSD						log_file => 'warn',
###LogSD					},
###LogSD				},
###LogSD				_build_out_the_cell =>{
###LogSD					UNBLOCK =>{
###LogSD						log_file => 'warn',
###LogSD					},
###LogSD				},
###LogSD				_get_row_all =>{
###LogSD					UNBLOCK =>{
###LogSD						log_file => 'warn',
###LogSD					},
###LogSD				},
###LogSD				_get_col_row =>{
###LogSD					UNBLOCK =>{
###LogSD						log_file => 'warn',
###LogSD					},
###LogSD				},
###LogSD				_parse_column_row =>{
###LogSD					UNBLOCK =>{
###LogSD						log_file => 'warn',
###LogSD					},
###LogSD				},
###LogSD			},
###LogSD		} );
###LogSD	}

lives_ok{	$row_ref = $test_instance->fetchrow_hashref }
										"Get the next fetchrow_hashref for row: $x";
###LogSD	$phone->talk( level => 'debug', message => [ "row:", $row_ref ] );
			if( !ref $answer_list->[$position] ){
###LogSD	$phone->talk( level => 'debug', message => [ "Found and -end- flag: $answer_list->[$position]" ] );
is			$row_ref, $answer_list->[$position],
										"Check for the correct end of file flag: EOF";
			last HASHREFRUN;
			}else{
is_deeply	$row_ref, $answer_list->[$position],
										"..and validate the returned values";
			}
			$x++;
            }
is			$test_instance->fetchrow_hashref( 1 ), undef,
										"Check that calling for a row above or at the header in the table fails";
is			$test_instance->error, "The requested row -1- is at or above the bottom of the header rows ( 1 )",
										"..with the correct error message";
is_deeply	$test_instance->fetchrow_hashref( 3 ), $answer_list->[$x + $offset - 3],
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