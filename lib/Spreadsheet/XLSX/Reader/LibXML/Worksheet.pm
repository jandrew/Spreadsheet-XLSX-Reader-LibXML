package Spreadsheet::XLSX::Reader::LibXML::Worksheet;
use version; our $VERSION = qv('v0.38.14');
###LogSD	warn "You uncovered internal logging statements for Spreadsheet::XLSX::Reader::LibXML::Worksheet-$VERSION";

use Carp 'confess';
use	Moose::Role;
requires qw(
	_min_row					_max_row					_min_col
	_max_col					_get_col_row				_get_next_value_cell		
	_get_row_all				_get_merge_map
);
###LogSD	requires 'get_log_space', 'get_all_space';
use Types::Standard qw(
	Bool 						HasMethods					Enum
	Int							is_Int						ArrayRef
	is_ArrayRef					HashRef						is_HashRef
	is_Object					Str							is_Str
);# Int
use lib	'../../../../../lib',;
###LogSD	use Log::Shiras::Telephone;
###LogSD	use Log::Shiras::UnhideDebug;
#~ use	Spreadsheet::XLSX::Reader::LibXML::Cell;
use	Spreadsheet::XLSX::Reader::LibXML::Types qw(
		SpecialDecimal					SpecialZeroScientific
		SpecialOneScientific			SpecialTwoScientific
		SpecialThreeScientific			SpecialFourScientific
		SpecialFiveScientific
	);

#########1 Dispatch Tables    3#########4#########5#########6#########7#########8#########9

my $format_headers =[ qw(
		cell_font		cell_border			cell_style
		cell_fill		cell_coercion		cell_alignment
	) ];

#########1 Public Attributes  3#########4#########5#########6#########7#########8#########9

has sheet_type =>(
		isa		=> Enum[ 'worksheet' ],
		default	=> 'worksheet',
		reader	=> 'get_sheet_type',
	);

has sheet_rel_id =>(
		isa		=> Str,
		reader	=> 'rel_id',
	);

has sheet_id =>(
		isa		=> Int,
		reader	=> 'sheet_id',#This feels like it might be broken but never tested?
	);

has sheet_position =>(# XML position
		isa		=> Int,
		reader	=> 'position',
	);

has sheet_name =>(
		isa		=> Str,
		reader	=> 'get_name',
	);

has last_header_row =>(
		isa			=> Int,
		reader		=> 'get_last_header_row',
		writer		=> '_set_last_header_row',
		clearer		=> '_clear_last_header_row',
		predicate	=> 'has_last_header_row'
	);

has min_header_col =>(
		isa			=> Int,
		reader		=> 'get_min_header_col',
		writer		=> 'set_min_header_col',
		clearer		=> 'clear_min_header_col',
		predicate	=> 'has_min_header_col'
	);

has max_header_col =>(
		isa			=> Int,
		reader		=> 'get_max_header_col',
		writer		=> 'set_max_header_col',
		clearer		=> 'clear_max_header_col',
		predicate	=> 'has_max_header_col'
	);

#########1 Public Methods     3#########4#########5#########6#########7#########8#########9

sub min_row{
	my( $self ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD			$self->get_all_space . '::Interface::row_bound::min_row', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"Returning the minimum row: " . $self->_min_row ] );
	return $self->_min_row;
}

sub max_row{
	my( $self ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD			$self->get_all_space . '::Interface::row_bound::max_row', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"Returning the maximum row: " . $self->_max_row ] );
	return $self->_max_row;
}

sub min_col{
	my( $self ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD			$self->get_all_space . '::Interface::row_bound::min_col', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"Returning the minimum column: " . $self->_min_col ] );
	return $self->_min_col;
}

sub max_col{
	my( $self ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD			$self->get_all_space . '::Interface::row_bound::max_col', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"Returning the maximum column: " . $self->_max_col ] );
	return $self->_max_col;
}

sub row_range{
	my( $self ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD			$self->get_all_space . '::Interface::row_bound::row_range', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"Returning row range( " . $self->_min_row . ", " . $self->_max_row . " )" ] );
	return( $self->_min_row, $self->_max_row );
}

sub col_range{
	my( $self ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD			$self->get_all_space . '::Interface::row_bound::col_range', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"Returning col range( " . $self->_min_col . ", " . $self->_max_col . " )" ] );
	return( $self->_min_col, $self->_max_col );
}

sub get_merged_areas{
	my( $self, ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD			$self->get_all_space . '::Interface::get_merged_areas', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			'Pulling the merge map ParseExcel style' ] );
	
	# Get the raw merge map
	my $raw_map = $self->_get_merge_map;
	###LogSD	$phone->talk( level => 'trace', message =>[
	###LogSD		"Raw merge row map;", $raw_map] );
	my ( $new_map, $dup_ref );
	#parse out the empty rows
	for my $row ( @$raw_map ){
		next if !$row;
		###LogSD	$phone->talk( level => 'trace', message =>[
		###LogSD		"Processing the merge row data:", $row] );
		for my $merge_cell ( @$row ){
			next if !$merge_cell;
			next if exists $dup_ref->{$merge_cell};
			###LogSD	$phone->talk( level => 'trace', message =>[
			###LogSD		"Processing the merge row data: $merge_cell"] );
			my $merge_ref;
			for my $cell ( split /:/, $merge_cell ){
				my ( $column, $row ) = $self->parse_column_row( $cell );
				push @$merge_ref, $row, $column;
				###LogSD	$phone->talk( level => 'trace', message =>[
				###LogSD		"Updated merge ref:", $merge_ref] );
			}
			$dup_ref->{$merge_cell} = 1;
			push @$new_map, $merge_ref;
			###LogSD	$phone->talk( level => 'trace', message =>[
			###LogSD		"Updated merge areas:", $new_map] );
		}
	}
	###LogSD	$phone->talk( level => 'info', message =>[
	###LogSD		"Final merge areas:", $new_map] );
	return $new_map;
}

sub is_column_hidden{
	my( $self, @column_requests ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD			$self->get_all_space . '::Interface::is_column_hidden', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			'Pulling the hidden state for the columns:', @column_requests ] );
	
	my @number_list;
	for my $item ( @column_requests ){
		if( is_Int( $item ) ){
			push @number_list, $item;
		}else{
			my ( $column, $dummy_row ) =  $self->parse_column_row( $item );
			###LogSD	$phone->talk( level => 'trace', message => [
			###LogSD		"Parsed -$item- to column number: $column" ] );
			push @number_list, $self->_get_excel_position( $column );# Convert to excel since 'around' didn't
		}
	}
	my @tru_dat = $self->_is_column_hidden( @number_list );
	my $true = 0;
	map{ $true = 1 if $_ } @tru_dat;
	###LogSD	$phone->talk( level => 'info', message =>[
	###LogSD		"Final column hidden state is -$true- with list:", @tru_dat] );
	return wantarray ? @tru_dat : $true;
}

sub is_row_hidden{
	my( $self, @row_requests ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD			$self->get_all_space . '::Interface::is_row_hidden', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			'Pulling the hidden state for the rows:', [@row_requests] ] );
	
	my @tru_dat;
	my $true = 0;
	for my $row ( @row_requests ){
		my $answer = $self->_get_row_hidden( $row );
		###LogSD	$phone->talk( level => 'debug', message =>[
		###LogSD		"Row answer is:", $answer ] );
		$true = 1 if $answer;
		push @tru_dat, ( ($self->min_row > $row or $self->max_row < $row) ? undef :  $answer ? 1 : 0 );
	}
	###LogSD	$phone->talk( level => 'info', message =>[
	###LogSD		"Final row hidden state is -$true- with list:", @tru_dat] );
	return wantarray ? @tru_dat : $true;
}

sub get_cell{
    my ( $self, $requested_row, $requested_column ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD			$self->get_all_space . '::Interface::get_cell', );
	###LogSD		$phone->talk( level => 'info', message =>[
	###LogSD			"Arrived at get_cell with: ",
	###LogSD			"Requested row: " . (defined( $requested_row ) ? $requested_row : ''),
	###LogSD			"Requested column: " . (defined( $requested_column ) ? $requested_column : '' ),] );
	
	# Ensure we have a good column and row to work from
	if( !defined $requested_row ){
		$self->set_error( "No row provided" );
		return undef;
	}
	if( !defined $requested_column ){
		$self->set_error( "No column provided" );
		return undef;
	}
	
	# Get information
	my $result = $self->_get_col_row( $requested_column, $requested_row );
	
	# Handle EOF EOR flags
	my $return = undef;
	if( $result ){
		if( is_HashRef( $result ) ){
			$return = $self->_build_out_the_cell( $result );
		}else{
			$return = $result if $self->boundary_flag_setting;
			( $requested_column, $requested_row ) = 
				( $result eq 'EOR' ) ? ( 0, $requested_row + 1 ) : ( 0, 0 );
		}
	}
	$self->_set_reported_row_col( [ $requested_row, $requested_column ] );
	###LogSD	$phone->talk( level => 'debug', message =>[
	###LogSD		"Set the reported [ row, col ] to: [ $requested_row, $requested_column ]", ] );
	###LogSD	$phone->talk( level => 'trace', message =>[ 'Final return:', $return ] );
	
	return $return;
}

sub get_next_value{
    my ( $self, ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD			$self->get_all_space . '::Interface::get_next_value', );
	###LogSD		$phone->talk( level => 'info', message =>[ 'Arrived at get_next_value', ] );
	
	my $result = $self->_get_next_value_cell;
	###LogSD	$phone->talk( level => 'info', message =>[ 'Next value;', $result ] );
	
	# Handle EOF EOR flags
	my $return = undef;
	my ( $reported_row, $reported_col ) = ( 0, 0 );
	if( $result ){
		if( is_HashRef( $result ) ){
			( $reported_row, $reported_col ) = ( $result->{cell_row}, $result->{cell_col} );
			$return = $self->_build_out_the_cell( $result );
		}elsif( $self->boundary_flag_setting ){
			$return = $result;
		}
	}
	$self->_set_reported_row_col( [ $reported_row, $reported_col ] );
	###LogSD	$phone->talk( level => 'debug', message =>[
	###LogSD		"Set the reported [ row, col ] to: [ $reported_row, $reported_col ]",
	###LogSD		"Returning the ref:", $return ] );
	return $return;
}

sub fetchrow_arrayref{
    my ( $self, $row ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD			$self->get_all_space . '::Interface::fetchrow_arrayref', );
	###LogSD		$phone->talk( level => 'info', message =>[
	###LogSD			"Arrived at fetchrow_arrayref for row: " . ((defined $row) ? $row : ''), ] );
	
	# Handle an implied next
	if( !defined $row ){
		my $last_row = $self->_get_reported_row;# Even if a cell is not at the end was last reported
		$row = $last_row + 1;
		###LogSD	$phone->talk( level => 'debug', message =>[
		###LogSD			"Resolved an implied 'next' row request to row: $row", ] );
	}
	
	my $result = $self->_get_row_all( $row );
	###LogSD	$phone->talk( level => 'debug', message =>[ 'Returned row ref;', $result ] );
	my $return = [];
	my ( $reported_row, $reported_col ) = ( $row, ($self->_max_col + 1) );
	if( $result ){
		if( is_ArrayRef( $result ) ){
			for my $cell ( @$result ){
				if( is_HashRef( $cell ) ){
					###LogSD	$phone->talk( level => 'debug', message =>[
					###LogSD		'Building out the cell:', $cell ] );
					push @$return, $self->_build_out_the_cell( $cell, );
				}else{
					push @$return, $cell;
				}
			}
		}else{
			if( $self->boundary_flag_setting ){
				###LogSD	$phone->talk( level => 'debug', message =>[
				###LogSD		'Package requests return of boundary flags' ] );
				$return = $result;
			}else{
				###LogSD	$phone->talk( level => 'debug', message =>[
				###LogSD		'Returning undef value for end of file state since boundary flags are off' ] );
				$return = undef;
			}
			# Handle EOF flags;
			( $reported_row, $reported_col ) = ( 0, 0 );
		}
	}
	
	# Handle full rows with empty_is_end = 0
	$self->_set_reported_row_col( [ $reported_row, $reported_col ] );
	###LogSD	$phone->talk( level => 'debug', message =>[
	###LogSD		"Set the reported [ row, col ] to: [ $reported_row, $reported_col ]", ] );
	###LogSD	$phone->talk( level => 'trace', message =>[ 'Final return:', $return ] );
	
	return $return;
}

sub fetchrow_array{
    my ( $self, $row ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD			$self->get_all_space . '::Interface::fetchrow_array', );
	###LogSD		$phone->talk( level => 'info', message =>[
	###LogSD			"Arrived at fetchrow_array for row: " . ((defined $row) ? $row : ''), ] );
	my $array_ref = $self->fetchrow_arrayref( $row );
	###LogSD	$phone->talk( level => 'trace', message =>[ 'Initial return:', $array_ref ] );
	my @return = 
		is_ArrayRef( $array_ref ) ? @$array_ref :
		is_Str( $array_ref ) ? $array_ref : ();
	###LogSD	$phone->talk( level => 'trace', message =>[ 'Final return:', @return ] );
	return @return;
}

sub set_headers{
    my ( $self, @header_row_list ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD			$self->get_all_space . '::Interface::set_headers', );
	###LogSD		$phone->talk( level => 'info', message =>[
	###LogSD			"Arrived at set_headers for row: ", @header_row_list, ] );
	my $header_ref;
	$self->_clear_last_header_row;
	$self->_clear_header_ref;
	my $old_output = $self->get_group_return_type;
	###LogSD		$phone->talk( level => 'info', message =>[
	###LogSD			"Old output type: $old_output", ] );
	my $new_output = $old_output;
	if( $old_output eq 'instance' ){
		$self->set_group_return_type( 'value' );
		$new_output = 'value';
	}else{
		$old_output = undef;
	}
	###LogSD		$phone->talk( level => 'info', message =>[
	###LogSD			"New output type: $new_output", ] );
	if( scalar( @header_row_list ) == 0 ){
		$self->set_error( "No row numbers passed to use as headers" );
		return undef;
	}
	my $last_header_row = 0;
	my $code_ref;
	for my $row ( @header_row_list ){
		if( ref( $row ) ){
			$code_ref = $row;
			###LogSD	$phone->talk( level => 'info', message =>[
			###LogSD		"Found header manipulation code: ", $code_ref, ] );
			next;
		}
		$last_header_row = $row if $row > $last_header_row;
		my $array_ref = $self->fetchrow_arrayref( $row );
		###LogSD	$phone->talk( level => 'info', message =>[
		###LogSD		"Retreived header row -$row- with values: ", $array_ref, ] );
		for my $x ( 0..$#$array_ref ){
			$header_ref->[$x] = $array_ref->[$x] if !defined $header_ref->[$x];
		}
		###LogSD	$phone->talk( level => 'info', message =>[
		###LogSD		"Updated header ref: ", $header_ref, ] );
	}
	if( $code_ref ){
		my $scrubbed_headers;
		for my $header ( @$header_ref ){
			push @$scrubbed_headers, $code_ref->( $header );
		}
		###LogSD	$phone->talk( level => 'info', message =>[
		###LogSD		"scrubbed header ref: ", $scrubbed_headers, ] );
		$header_ref = $scrubbed_headers;
	}
	$self->_set_last_header_row( $last_header_row );
	$self->_set_header_ref( $header_ref );
	$self->set_group_return_type( $old_output ) if $old_output;
	return $header_ref;
}

sub fetchrow_hashref{
    my ( $self, $row ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD			$self->get_all_space . '::Interface::fetchrow_hashref', );
	###LogSD		$phone->talk( level => 'info', message =>[
	###LogSD			"Arrived at fetchrow_hashref for row: " . ((defined $row) ? $row : ''), ] );
	# Check that the headers are set
	if( !$self->_has_header_ref ){
		$self->set_error( "Headers must be set prior to calling fetchrow_hashref" );
		return undef;
	}elsif( defined $row and $row <= $self->get_last_header_row ){
		$self->set_error(
			"The requested row -$row- is at or above the bottom of the header rows ( " .
			$self->get_last_header_row . ' )'
		);
		return undef;
	}
	my $array_ref = $self->fetchrow_arrayref( $row );
	return $array_ref if !$array_ref or $array_ref eq 'EOF';
	my $header_ref = $self->_get_header_ref;
	my ( $start, $end ) = ( $self->min_col, $self->max_col );
	my ( $min_col, $max_col ) = ( $self->get_min_header_col, $self->get_max_header_col );
	###LogSD	$phone->talk( level => 'info', message =>[
	###LogSD		((defined $min_col) ? "Minimum header column: $min_col" : undef),
	###LogSD		((defined $max_col) ? "Maximum header column: $max_col" : undef), ] );
	$min_col = ($min_col and $min_col>$start) ? $min_col - $start : 0;
	$max_col = ($max_col and $max_col<$end) ? $end - $max_col : 0;
	###LogSD	$phone->talk( level => 'info', message =>[
	###LogSD		((defined $min_col) ? "Minimum header column offset: $min_col" : undef),
	###LogSD		((defined $max_col) ? "Maximum header column offset: $max_col" : undef), ] );
	
	# Build the ref
	my $return;
	my $blank_count = 0;
	for my $x ( (0+$min_col)..($self->max_col-$max_col) ){
		my $header = defined( $header_ref->[$x] ) ? $header_ref->[$x] : 'blank_header_' . $blank_count++;
		###LogSD	$phone->talk( level => 'info', message =>[
		###LogSD		"Possibly adding value for header: $header" ] );
		if( defined $array_ref->[$x] ){
			###LogSD	$phone->talk( level => 'info', message =>[
			###LogSD		"Adding value: $array_ref->[$x]" ] );
			$return->{$header} = $array_ref->[$x];
		}
	}
	
	return $return;
}

sub set_custom_formats{
    my ( $self, @input_args ) = @_;
	my $args; 
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD			$self->get_all_space . '::Interface::set_custom_formats', );
	###LogSD		$phone->talk( level => 'info', message =>[
	###LogSD			"Arrived at set_custom_formats with: ", @input_args, ] );
	my $worksheet_custom = 0;
	if( !@input_args ){
		$self->( "The input args to 'set_custom_formts' are empty - no op" );
		return undef;
	}elsif( is_HashRef( $input_args[0] ) and @input_args == 1 ){
		$args = $input_args[0];
	}elsif( @input_args % 2 == 0 ){
		$args = { @input_args };
	}else{
		$self->set_error( "Unable to coerce input args to a hashref: " . join( '~|~', @input_args ) );
		return undef;
	}
	###LogSD	$phone->talk( level => 'info', message =>[
	###LogSD			"Now acting on: ", $args ] );
	my $final_load;
	for my $key ( keys %$args ){
		my $new_coercion;
		if( $key eq '' or $key !~ /[A-Z]{0,3}(\d*)/ ){
			$self->set_error( "-$key- is not an allowed custom format key" );
			next;
		}elsif( is_Object( $args->{$key} ) ){
			###LogSD	$phone->talk( level => 'info', message =>[
			###LogSD			"Key -$key- already has an object" ] );
			$new_coercion = $args->{$key};
		}else{
			###LogSD	$phone->talk( level => 'info', message =>[
			###LogSD			"Trying to build a new coercion for -$key- with: $args->{$key}" ] );
			$new_coercion = $self->parse_excel_format_string( $args->{$key}, "Worksheet_Custom_" . $worksheet_custom++ );
			if( !$new_coercion ){
				$self->set_error( "No custom coercion could be built for -$key- with: $args->{$key}" );
				next;
			}
			###LogSD	$phone->talk( level => 'info', message =>[
			###LogSD			"Built possible new coercion for -$key-" ] );
		}
		if( !$new_coercion->can( 'assert_coerce' ) ){
			$self->set_error( "The identified coercion for -$key- cannot 'assert_coerce'" );
		}elsif( !$new_coercion->can( 'display_name' ) ){
			$self->set_error( "The custom coercion for -$key- cannot 'display_name'" );
		}else{
			###LogSD	$phone->talk( level => 'info', message =>[
			###LogSD			"Loading -$key- with coercion: " . $new_coercion->display_name ] );
			$final_load->{$key} = $new_coercion;
		}
	}
	$self->_set_custom_format( %$final_load );
}

#########1 Private Attributes 3#########4#########5#########6#########7#########8#########9

has _custom_formats =>(
		isa		=> HashRef[ HasMethods[ 'assert_coerce', 'display_name' ] ],
		traits	=> ['Hash'],
		reader	=> 'get_custom_formats',
		default => sub{ {} },
		clearer	=> '_clear_custom_formats',
		handles	=>{
			has_custom_format => 'exists',
			get_custom_format => 'get',
			_set_custom_format => 'set',
		},
	);

has _header_ref =>(
		isa			=> ArrayRef,
		writer		=> '_set_header_ref',
		reader		=> '_get_header_ref',
		clearer		=> '_clear_header_ref',
		predicate	=> '_has_header_ref',
	);

has _reported_row_col =>(# Manage (store and retreive) in count from 1 mode
		isa		=> ArrayRef[Int],
		traits	=> ['Array'],
		writer	=> '_set_reported_row_col',
		reader	=> '_get_reported_row_col',
		default	=> sub{ [ 0, 0 ] },# Pre-row and pre-col
		handles =>{
			_get_reported_row => [ get => 0 ],
			_get_reported_col => [ get => 1 ],
		},
	);

#########1 Private Methods    3#########4#########5#########6#########7#########8#########9

sub BUILD {
###LogSD	my	$self = shift;
###LogSD		$self->set_class_space( 'Worksheet' );
			require Spreadsheet::XLSX::Reader::LibXML::Cell;
###LogSD		Spreadsheet::XLSX::Reader::LibXML::Cell->import( $self->get_log_space );
}

sub _build_out_the_cell{
	my ( $self, $result, ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD			$self->get_all_space . '::Interface::_hidden::_build_out_the_cell', );
	###LogSD		$phone->talk( level => 'debug', message =>[  
	###LogSD			 "Building out the cell ref:", $result, "..with results as: ". $self->get_group_return_type ] );
	my ( $return, $hidden_format );
	$return->{cell_xml_value} = $result->{cell_xml_value} if defined $result->{cell_xml_value};
	if( is_HashRef( $result ) ){
		###LogSD	$phone->talk( level => 'debug', message =>[
		###LogSD		"processing cell object from cell ref:", $result ] );
		my $scientific_format;
		if( exists $return->{cell_xml_value} and defined $result->{cell_xml_value} and $result->{cell_xml_value} =~ /^(\-)?((\d+)?\.?(\d+)?)[Ee](\-)?(\d+)$/ and $2 and $6 ){#Implement implied output formatting intrensic to Excel for scientific notiation
			###LogSD	$phone->talk( level => 'trace', message =>[
			###LogSD		"Found special scientific notation case were stored values and visible values possibly differ" ] );
			my	$dec_sign = $1 ? $1 : '';
			my	$decimal = $2;
			my	$exp_sign = $5 ? $5 : '';
			my	$exponent = $6;
				$decimal = sprintf '%.14f', $decimal;
			$decimal =~ /([1-9])?\.(.*[1-9])?(0*)$/;
			my	$last_sig_digit = 
					!$2         ? 0 :
					defined $3 ? 14 - length( $3 ) : 14 ;
			my $initial_significant_digits = length( $exp_sign ) > 0 ? ($last_sig_digit + $exponent) : ($last_sig_digit - $exponent);
			#~ my $sig_digit_delta = 
			###LogSD	$phone->talk( level => 'debug', message =>[
			###LogSD		"Processing decimal                          : $decimal",
			###LogSD		"Final significant digit of the decimal is at: $last_sig_digit",
			###LogSD		"Total significant digits                    : $initial_significant_digits", ] );
			if( $initial_significant_digits > 19 ){
				###LogSD	$phone->talk( level => 'debug', message =>[
				###LogSD		"Attempting to use sprintf: %.${last_sig_digit}f", ] );
				$return->{cell_unformatted}  = $dec_sign . sprintf "%.${last_sig_digit}f", $decimal;
				$return->{cell_unformatted} .= 'E' . $exp_sign . sprintf '%02d', $exponent;
				###LogSD	$phone->talk( level => 'debug', message =>[
				###LogSD		"Found the unformatted scientific notation case with result: $return->{cell_unformatted}"] );
			}else{
				###LogSD	$phone->talk( level => 'trace', message =>[
				###LogSD		"Attempting to use sprintf: %.${initial_significant_digits}f", ] );
				$return->{cell_unformatted} = sprintf "%.${initial_significant_digits}f", $return->{cell_xml_value};
				###LogSD	$phone->talk( level => 'debug', message =>[
				###LogSD		"Found the unformatted decimal case with output: $return->{cell_unformatted}"] );
			}
			
			my	$short_decimal = sprintf '%.5f', $decimal;
				$short_decimal =~ /([1-9])?\.(.*[1-9])?(0*)$/;
			my	$short_sig_digit = 
					!$2         ? 0 :
					defined $3 ? 5 - length( $3 ) : 5 ;
			
			$scientific_format =
				( $initial_significant_digits < 10  ) ? SpecialDecimal :
				( $short_sig_digit == 0 ) ? SpecialZeroScientific :
				( $short_sig_digit == 1 ) ? SpecialOneScientific :
				( $short_sig_digit == 2 ) ? SpecialTwoScientific :
				( $short_sig_digit == 3 ) ? SpecialThreeScientific :
				( $short_sig_digit == 4 ) ? SpecialFourScientific :
				( $short_sig_digit == 5 ) ? SpecialFiveScientific :
					SpecialZeroScientific ;
			###LogSD	$phone->talk( level => 'debug', message =>[
			###LogSD		"Resolved the final formatted output to formatter: " . $scientific_format->display_name ] );
		}
		$return->{cell_type} = $result->{cell_type};
		$return->{r} = $result->{r};
		$return->{cell_merge} = $result->{cell_merge} if exists $result->{cell_merge};
		$return->{cell_hidden} = $result->{cell_hidden} if exists $result->{cell_hidden};
		if( !exists $return->{cell_unformatted} and exists $result->{cell_xml_value} ){
			@$return{qw( cell_unformatted rich_text )} = @$result{qw( cell_xml_value rich_text )};
			delete $return->{rich_text} if !$return->{rich_text};
		}
		
		#Implement user defined changes in encoding
		if( $return->{cell_unformatted} and length( $return->{cell_unformatted} ) > 0 ){
			$return->{cell_unformatted} = $self->change_output_encoding( $return->{cell_unformatted} );
			$return->{cell_xml_value} = $self->change_output_encoding( $return->{cell_xml_value} );
			###LogSD	$phone->talk( level => 'debug', message =>[
			###LogSD		"Unformatted with output encoding changed: " . $return->{cell_unformatted} ] );# if defined $return->{cell_unformatted};
		}
		if( $return->{cell_xml_value} and length( $return->{cell_xml_value} ) > 0 ){#Implement user defined changes in encoding
			$return->{cell_xml_value} = $self->change_output_encoding( $return->{cell_xml_value} );
			###LogSD	$phone->talk( level => 'debug', message =>[
			###LogSD		"XML with output encoding changed: " . $return->{cell_xml_value} ] );# if defined $return->{cell_unformatted};
		}
		if( $self->get_group_return_type eq 'unformatted' ){
			###LogSD	$phone->talk( level => 'debug', message =>[
			###LogSD		"Sending back just the unformatted value: " . ($return->{cell_unformatted}//'') ] ) ;
			return $return->{cell_unformatted};
		}elsif( $self->get_group_return_type eq 'xml_value' ){
			###LogSD	$phone->talk( level => 'debug', message =>[
			###LogSD		"Sending back just the unformatted value: " . ($return->{cell_xml_value}//'') ] ) ;
			return $return->{cell_xml_value};
		}
		
		# Get any relevant custom format
		my	$custom_format;
		if( $self->has_custom_format( $result->{r} ) ){
			###LogSD	$phone->talk( level => 'debug', message =>[
			###LogSD		"Custom format exists for: $result->{r}",] );
			$custom_format = $self->get_custom_format( $result->{r} );
		}else{
			$result->{r} =~ /([A-Z]+)(\d+)/;
			my ( $col_letter, $excel_row ) = ( $1, $2 );
			if( $self->has_custom_format( $col_letter ) ){
				###LogSD	$phone->talk( level => 'debug', message =>[
				###LogSD		"Custom format exists for column: $col_letter",] );
				$custom_format = $self->get_custom_format( $col_letter );
			}elsif( $self->has_custom_format( $excel_row ) ){
				###LogSD	$phone->talk( level => 'debug', message =>[
				###LogSD		"Custom format exists for row: $excel_row",] );
				$custom_format = $self->get_custom_format( $excel_row );
			}
		}
		
		# Initial check for return of value only (custom format case)
		if( $custom_format ){
			###LogSD	$phone->talk( level => 'debug', message =>[
			###LogSD		'Cell custom_format is:', $custom_format ] );
			if( $self->get_group_return_type eq 'value' ){
				###LogSD	$phone->talk( level => 'debug', message =>[
				###LogSD		'Applying custom format to: ' .  $return->{cell_unformatted} ] );
				###LogSD	$phone->talk( level => 'trace', message =>[
				###LogSD		'Returning value coerced by custom format:', $custom_format ] );
				return	Spreadsheet::XLSX::Reader::LibXML::Cell->_return_value_only(
							$return->{cell_unformatted}, 
							$custom_format,
							$self->_get_error_inst,
				###LogSD	$self->get_log_space,
						);
			}
			$return->{cell_coercion} = $custom_format;
			$return->{cell_type} = 'Custom';
		}
		
		# handle the formula
		if( exists $result->{cell_formula}  ){
			$return->{cell_formula} = $result->{cell_formula};
		}
		
		# convert the row column to user defined
		$return->{cell_row} = $self->_get_used_position( $result->{cell_row} );
		$return->{cell_col} = $self->_get_used_position( $result->{cell_col} );
		###LogSD	$phone->talk( level => 'debug', message =>[
		###LogSD		"Cell args to this point:", $return] );
		
		if( exists $result->{s} ){
			my $header = ($self->get_group_return_type eq 'value') ? 'cell_coercion' : undef;
			my $exclude_header = ($custom_format) ? 'cell_coercion' : undef;
			my $format;
			if( $header and $exclude_header and $header eq $exclude_header ){
				###LogSD	$phone->talk( level => 'info', message =>[
				###LogSD		"It looks like you just want to just return the formatted value but there is already a custom format" ] );
			}elsif( $self->_has_styles_file ){
				###LogSD	$phone->talk( level => 'debug', message =>[
				###LogSD		"Pulling formats with:", $result->{s}, $header, $exclude_header ] );
				$format = $self->get_format_position( $result->{s}, $header, $exclude_header );
				###LogSD	$phone->talk( level => 'trace', message =>[
				###LogSD		"format position is:", $format ] );
			}else{
				confess "'s' element called out but the style file is not available!";
			}
			###LogSD	$phone->talk( level => 'debug', message =>[
			###LogSD		"Checking if the defined number format needs replacing with:", $custom_format, $scientific_format] );
			if( $custom_format ){
				###LogSD	$phone->talk( level => 'debug', message =>[
				###LogSD		"Custom formats override this cell", $custom_format->display_name] );
				$return->{cell_coercion} = $custom_format;
				delete $format->{cell_coercion};
			}elsif( $scientific_format and
						(	!exists $format->{cell_coercion} or 
							$format->{cell_coercion}->display_name eq 'Excel_number_0' or 
							$format->{cell_coercion}->display_name eq 'Excel_text_0'		) ){
				###LogSD	$phone->talk( level => 'debug', message =>[
				###LogSD		"The generic number case will implement a hidden scientific format", $scientific_format] );
				$return->{cell_coercion} = $scientific_format;
				delete $format->{cell_coercion};
			}
			# Second check for value only - for the general number case not just custom formats
			if( $self->get_group_return_type eq 'value' ){
				###LogSD	$phone->talk( level => 'debug', message =>[
				###LogSD		'Applying (a possible) regular format to: ' .  ($return->{cell_unformatted}//''), $return, $format ] );
				return	Spreadsheet::XLSX::Reader::LibXML::Cell->_return_value_only(
							$return->{cell_unformatted}, 
							$return->{cell_coercion} // $format->{cell_coercion},
							$self->_get_error_inst,
				###LogSD	$self->get_log_space,
						);
			}
			if( $self->_has_styles_file ){
				###LogSD	$phone->talk( level => 'debug', message =>[
				###LogSD		"Format headers are:", $format_headers ] );
				for my $header ( @$format_headers ){
					if( exists $format->{$header} ){
						###LogSD	$phone->talk( level => 'trace', message =>[
						###LogSD		"Transferring styles header -$header- to the cell", ] );
						$return->{$header} = $format->{$header};
						if( $header eq 'cell_coercion' ){
							if(	$return->{cell_type} eq 'Numeric' and
								$format->{$header}->name =~ /date/i ){
								###LogSD	$phone->talk( level => 'trace', message =>[
								###LogSD		"Found a -Date- cell", ] );
								$return->{cell_type} = 'Date';
							}
						}
					}
				}
				###LogSD	$phone->talk( level => 'trace', message =>[
				###LogSD		"Practice special old spreadsheet magic here as needed - for now only single quote in the formula bar",  ] );
				if( exists $format->{quotePrefix} ){
					###LogSD	$phone->talk( level => 'debug', message =>[
					###LogSD		"Found the single quote in the formula bar case",  ] );# Other similar cases include carat and double quote in the formula bar (middle and right justified)
					$return->{cell_alignment}->{horizontal} = 'left';
					$return->{cell_formula} = $return->{cell_formula} ? ("'" . $return->{cell_formula}) : "'";
				}
					
			}
		}
		###LogSD	$phone->talk( level => 'debug', message =>[
		###LogSD		"Checking if a scientific format should be used", $scientific_format] );
		if( $scientific_format and !exists $return->{cell_coercion} ){
			###LogSD	$phone->talk( level => 'debug', message =>[
			###LogSD		"The generic number case will implement a hidden scientific format", $scientific_format] );
			$return->{cell_coercion} = $scientific_format;
		}
			
		###LogSD	$phone->talk( level => 'trace', message =>[
		###LogSD		"Checking return type: " . $self->get_group_return_type,  ] );
		# Final check for value only - for the text case
		if( $self->get_group_return_type eq 'value' ){
			###LogSD	$phone->talk( level => 'debug', message =>[
			###LogSD		'Applying (a possible) regular format to: |' .  ($return->{cell_unformatted}//'') . '|' ] );
			return	Spreadsheet::XLSX::Reader::LibXML::Cell->_return_value_only(
						$return->{cell_unformatted}, 
						$return->{cell_coercion},
						$self->_get_error_inst,
			###LogSD	$self->get_log_space,
					);
		}
		$return->{error_inst} = $self->_get_error_inst;
		###LogSD	$result->{log_space} = $self->get_log_space;
		###LogSD	$phone->talk( level => 'debug', message =>[
		###LogSD		"Final args ref is:", $return] );
	}elsif( $result ){
		# This is where you return 'EOF' and 'EOR' flags
		return ( $self->boundary_flag_setting ) ? $result : undef;
	}

	# build a cell
	my $cell = Spreadsheet::XLSX::Reader::LibXML::Cell->new( %$return );
	###LogSD		$phone->talk( level => 'debug', message =>[
	###LogSD			"Cell is:", $cell ] );
	return $cell;
}

around [ qw( build_cell_label get_cell fetchrow_arrayref is_column_hidden is_row_hidden ) ] => sub{ #
	my ( $method, $self, @input_list ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD			$self->get_all_space . '::Interface::_hidden::scrubbing_input', );
	###LogSD		$phone->talk( level => 'debug', message =>[ 'Received:', @input_list ] );
	my @new_list;
	for my $input ( @input_list ){
		push @new_list, ( is_Int( $input ) ? $self->_get_excel_position( $input ) : $input );
	}
	###LogSD	$phone->talk( level => 'debug', message =>[ 'executing list:', @new_list ] );
	my @output_list = $self->$method( @new_list );
	###LogSD	$phone->talk( level => 'debug', message =>[ 'returning list:', @output_list ] );
	return  wantarray ? @output_list : $output_list[0];
};

around [ qw( parse_column_row min_col max_col min_row max_row row_range col_range ) ] => sub{
	my ( $method, $self, @input_list ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD			$self->get_all_space . '::Interface::_hidden::scrubbing_output', );
	###LogSD		$phone->talk( level => 'debug', message =>[ 'Executing ' . (caller( 2 ))[3] . '(' . join( ', ', @input_list) . ')'  ] );
	my @new_list = $self->$method( @input_list );
	###LogSD	$phone->talk( level => 'debug', message =>[ 'method result:', @new_list ] );
	my @output_list = ();
	for my $output ( @new_list ){
		push @output_list, ( is_Int( $output ) ? $self->_get_used_position( $output ) : $output );
	}
	###LogSD	$phone->talk( level => 'debug', message =>[ 'returning list:', @output_list ] );
	return wantarray ? @output_list : $output_list[0];
};

sub _get_excel_position{
	my ( $self, $used_int ) = @_;
	return undef if !defined $used_int;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD			$self->get_all_space . '::Interface::_hidden::scrubbing_input::_get_excel_position', );
	###LogSD		$phone->talk( level => 'debug', message =>[
	###LogSD			"Converting used number  -$used_int- to Excel" ] );
	my	$excel_position = $used_int;
	$excel_position += 1 if $self->counting_from_zero;
	###LogSD		$phone->talk( level => 'debug', message =>[
	###LogSD			"New position is: $excel_position" ] );
	return $excel_position;
}

sub _get_used_position{
	my ( $self, $excel_int ) = @_;
	return undef if !defined $excel_int;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD			$self->get_all_space . '::Interface::_hidden::scrubbing_output::_get_used_position', );
	###LogSD		$phone->talk( level => 'debug', message =>[
	###LogSD			"Converting Excel  -$excel_int- to the used number" ] );
	my	$used_position = $excel_int;
	$used_position -= 1 if $self->counting_from_zero;
	###LogSD		$phone->talk( level => 'debug', message =>[
	###LogSD			"The used position is: $used_position" ] );
	return $used_position;
}

#########1 Phinish            3#########4#########5#########6#########7#########8#########9

no Moose::Role;
1;

#########1 Documentation      3#########4#########5#########6#########7#########8#########9
__END__

=head1 NAME

Spreadsheet::XLSX::Reader::LibXML::Worksheet - Top level xlsx Worksheet interface

=head1 SYNOPSIS

If you are looking for the synopsis for the package see 
L<Spreadsheet::XLSX::Reader::LibXML/SYNOPSIS>.  Otherwise the best example for use of 
this module alone is the test file in this package 
t/Spreadsheet/XLSX/Reader/LibXML/10-get_cell.t
    
=head1 DESCRIPTION

This documentation is written to explain ways to use this module when writing your 
own excel parser.  To use the general package for excel parsing out of the box please 
review the documentation for L<Workbooks|Spreadsheet::XLSX::Reader::LibXML>,
L<Worksheets|Spreadsheet::XLSX::Reader::LibXML::Worksheet>, and 
L<Cells|Spreadsheet::XLSX::Reader::LibXML::Cell>

This is the extracted L<Role|Moose::Manual::Roles> to be used as a top level worksheet 
interface.  This is the place where all the various details in each sub XML sheet are 
coallated into a set of data representing all the necessary information for a requested 
cell.  Since this is the center of data coallation all elements that may be customized 
should reside outside of this role.  This includes any specific elements that would be 
different between each of the sheet parser types and any element of Excel data presentation 
that may lend itself to customization.  For instance all the XML parser 
methods, (Reader, DOM, and possibly SAX) should exist outside and preferebly below this 
role.

This role is L<also|Spreadsheet::XLSX::Reader::LibXML::CellToColumnRow/DESCRIPTION> 
contains a layer of abstraction to allow for run time setting of count-from-one or 
count-from-zero mode.  The layer of abstraction is use with the Moose 
L<around|Moose::Manual::MethodModifiers/AROUND modifiers> modifier.

=head2 requires

These are method(s) used by this Role but not provided by the role.  Any class consuming this 
role will not build without first providing these methods prior to loading this role.  
I<Since this is the center of data collation the list is long>.

=head3 min_row

=over

B<Definition:> Used to get the minimum row with data in the worksheet.

=back

=head3 max_row

=over

B<Definition:> Used to get the maximum row with data in the worksheet.

=back

=head3 min_col

=over

B<Definition:> Used to get the minimum column with data in the worksheet.

=back

=head3 max_col

=over

B<Definition:> Used to get the maximum column with data in the worksheet.

=back

=head3 row_range

=over

B<Definition:> Used to return a list of the L<$minimum_row|/min_row> and 
L<$maximum_row|/max_row> values

=back

=head3 col_range

=over

B<Definition:> Used to return a list of the L<$minimum_column|/min_col> and 
L<$maximum_column|/max_col> values.

=back

=head3 _get_next_value_cell

=over

B<Definition:> This should return the next cell data from the worksheet file 
that contains unique formatting or information.  The data is expected in a 
perl hash ref.  This method should collect data left to right and top to 
bottom.  I<The styles.xml, sharedStrings.xml, and calcChain.xml etc. sheet data 
are coallated into the cell information at this point>.  An 'EOF' string should be 
returned when the file has reached the end and then the method should wrap 
back to the beginning.

B<Example of expected return data set>

   {
      'r'          => 'A6',    # The cell ID
      'cell_merge' => 'A6:B6', # The merge range
      'row'        => 6,       # count by 1 (no 'around' performed on leading '_' methods)
      'col'        => 1,       # count by 1 (no 'around' performed on leading '_' methods)
      's'          => '11',    # Styles type (position 11 in the styles sheet)
      't'          => 's'      # Cell data type (string)
      'v' =>{                  # Cell data (since this cell is string 
         'raw_text' => '15'    # data this actually points to position 
      }                        #    15 in the sharedStrings.xml file )
   }

=back

=head3 _get_next_cell

=over

B<Definition:> Like L<_get_next_value_cell|/_get_next_value_cell> this method should 
return the next cell.  The difference is it should return undef for empty cells 
rather than skipping them.  This method should collect data left to right and top 
to bottom.  I<The styles.xml, sharedStrings.xml, and calcChain.xml etc. sheet data 
are coallated into the cell information at this point>.  An 'EOF' string should be 
returned when the file has reached the end and then the method should wrap back to 
the beginning.

=back

=head3 _get_col_row

=over

B<Definition:> This method should provide a targeted way to return the worksheet 
file information on a cell.  It should only accept count-from-one column and row 
numbers and the column should be required before the row.  If the request is made 
for an out of row bounds position the method should provide an 'EOR' string.  An 
'EOF' string should be returned when the file has reached the end and then the 
method should wrap back to the beginning.

=back

I<The attribute L<workbook_instance|/workbook_instance> must also be filled 
correctly since it exports all of the the workbook level functionality to this class.> 

=head2 Primary Methods

These are the various methods provided by this role.  Each of them calls a sub method 
to get the base cell data and then coallates that information into the proper return 
value(s) defined by L<Spreadsheet::XLSX::Reader::LibXML/group_return_type>.

=head3 get_cell( $row, $column )

=over

B<Definition:> This calls the supplied method L<_get_col_row|/_get_col_row>.

B<Accepts:> the list ( $row, $column ) both required (and modified as needed by the 
attribute state of L<Spreadsheet::XLSX::Reader::LibXML/count_from_zero>)

B<Returns:> if data to build a cell instance is provided then the instance is collated, 
built, and returned.  Otherwise the value from '_get_col_row' is returned unfiltered.

=back

=head3 get_next_value

=over

B<Definition:> This calls the supplied method L<_get_next_value_cell|/_get_next_value_cell>

B<Accepts:> nothing

B<Returns:> if data to build a cell instance is provided then the instance is collated, 
built, and returned.  Otherwise the value from '_get_next_value_cell' is returned 
unfiltered.

=back

=head3 fetchrow_arrayref( $row )

=over

B<Definition:>  This calls the supplied method L<_get_row_all|/_get_row_all>.  It will 
return 'EOF' once instead of an array reference for the end of the file before resetting 
to the first row..

B<Accepts:> undef = next|$row = a row integer indicating the desired row (modified as 
needed by the attribute state of L<Spreadsheet::XLSX::Reader::LibXML/count_from_zero>)

B<Returns:> an array ref of all possible column positions in that row with data filled in 
as appropriate. (or 'EOF')

=back

=head3 fetchrow_array( $row )

=over

B<Definition:> This function calls L<fetchrow_arrayref|/fetchrow_arrayref( $row )> 
except it returns an array instead of an array ref

B<Accepts:> undef = next|$row = a row integer indicating the desired row

B<Returns:> an array of all possible column positions in that row with data filled in 
as appropriate.

=back

=head3 set_headers( @header_row_list [ \&header_scrubber ] )

=over

B<Definition:> This function is used to set headers used in the function 
L<fetchrow_hashref|/fetchrow_hashref( $row )>.  It accepts a list of row numbers that 
will be collated into a set of headers used to build the hashref for each row.
The header rows are coallated in sequence with the first number taking precedence.  
The list is also used to set the lowest row of the headers in the table.  All rows 
at that level and higher will be considered out of the table and will return undef 
while setting the error instance.  If some of the columns do not have values then 
the instance will auto generate unique headers for each empty header column to fill 
out the header ref.  [ optionally: it is possible to pass a coderef to scrub the 
headers so they make some sence. for example; ]

	my $scrubber = sub{
		my $input = $_[0];
		$input =~ s/\n//g if $input;
		$input =~ s/\s/_/g if $input;
		return $input;
	};
	$self->set_headers( 2, 1, $scrubber ); # Called internally as $new_value = $scrubber->( $old_value );
	# Returns/stores the headers set at row 2 and 1 with values from row 2 taking precedence
	#  Then it scrubs the values by removing newlines and replacing spaces with underscores.

B<Accepts:> a list of row numbers (modified as needed by the attribute state of 
L<Spreadsheet::XLSX::Reader::LibXML/count_from_zero>) and an optional L<closure
|http://www.perl.com/pub/2002/05/29/closure.html>.

B<Returns:> an array ref of the built headers for review

=back

=head3 fetchrow_hashref( $row )

=over

B<Definition:> This function is used to return a hashref representing the data in the 
specified row.  If no $row value is passed it will return the 'next' row of data.  A call 
to this function without L<setting|/set_headers( @header_row_list )> the headers first 
will return undef and set the error instance.  This function calls 
L<_get_row_all|/_get_row_all>.

B<Accepts:> a target $row number for return values or undef meaning 'next'

B<Returns:> a hash ref of the values for that row

=back

=head2 Attributes

Arguments that can be passed to new when creating a class instance or changed using 
one of the 'attribute methods'.   Where an attribute is delegating the 'attribute 
method' from a method in the instance stored in the attribute the documentation will 
indicate that the 'attribute method' is 'delegated'.  I<All 'delegated' methods are 
required for the instance to be accepted by the attribute.>  For more information on 
attributes see L<Moose::Manual::Attributes> and L<Moose::Manual::Delegation>.

=head3 last_header_row

=over

B<Definition:> This is generally set by the method L<set_headers( @header_row_list )
|/set_headers( @header_row_list )> method I<not during -E<gt>new> and is the largest row number 
of the @header_row_list I<not necessarily the last number in the sequence>.

B<Default:> undef

B<attribute methods> Methods provided to adjust this attribute
		
=over

B<get_last_header_row>

=over

B<Definition:> returns the value of the attribute

=back

B<has_last_header_row>

=over

B<Definition:> predicate for the attribute

=back

=back

=back

=head3 min_header_col

=over

B<Definition:> When the method L<fetchrow_hashref|/fetchrow_hashref( $row )> is 
called it is possible to only return a set of information between two defined 
columns.  This is the attribute that defines the start column.

B<Default:> undef

B<attribute methods> Methods provided to adjust this attribute
		
=over

B<set_min_header_col>

=over

B<Definition:> sets the value of the attribute

B<Range:> integer values I<Integers less than L<min_col|/min_col> will 
be ignored>

=back
		
=over

B<get_min_header_col>

=over

B<Definition:> returns the value of the attribute

=back

B<has_min_header_col>

=over

B<Definition:> predicate for the attribute

=back

B<clear_min_header_col>

=over

B<Definition:> sets min_header_col to 'undef'

=back

=back

=back

=back

=head3 max_header_col

=over

B<Definition:> When the method L<fetchrow_hashref|/fetchrow_hashref( $row )> 
is called it is possible to only collect a set of information between two defined 
columns.  This is the attribute that defines the end column.

B<Default:> undef

B<attribute methods> Methods provided to adjust this attribute
		
=over

B<set_max_header_col>

=over

B<Definition:> sets the value of the attribute

B<Range:> integer values I<Integers larger than L<max_col|/max_col> will 
be ignored>

=back
		
=over

B<get_max_header_col>

=over

B<Definition:> returns the value of the attribute

=back

B<has_max_header_col>

=over

B<Definition:> predicate for the attribute

=back

B<clear_max_header_col>

=over

B<Definition:> sets min_header_col to 'undef'

=back

=back

=back

=back

=head3 custom_formats

=over

B<Definition:> When this role is coallating data about a cell it will check this 
attribute before it checks the styles sheet to see if there is a format defined 
by the user for converting the L<unformatted
|Spreadsheet::XLSX::Reader::LibXML::Cell/unformatted> data.  The formats stored 
must have two methods 'assert_coerce' and 'display_name'.  The cell instance 
builder will consult this attribute by first checking the cellID as a key, then 
it checks for just the column letter(s) as a key, and finally it checks the row 
number as a key.  For an easy way to build custom conversion review the 
documentation for L<Type::Tiny|Type::Tiny::Manual::Libraries> and 
L<Type::Coercions|Type::Tiny::Manual::Coercions>. I<the Chained Coercions are 
very cool!.>

B<Default:> undef

B<attribute methods> Methods provided to adjust this attribute
		
=over

B<set_custom_formats( { $key =E<gt> $conversion } )>

=over

B<Definition:> a way to set all $key => $conversion pairs at once

B<Accepts:> a hashref of $key => $conversion pairs

=back
		
=over

B<has_custom_format( $key )>

=over

B<Definition:> checks if the specific $key for a format is registered

=back

B<get_custom_format( $key )>

=over

B<Definition:> get the custom format for the requested $key

B<Returns:> the $conversion registered to the $key

=back

B<set_custom_format( $key =E<gt> $conversion )>

=over

B<Definition:> set the custom format $conversion for the identified $key

=back

=back

=back

=back

=head3 workbook_instance

=over

B<Definition:> This is where the workbook level methods are accessed by 
the worksheet.  Because the workbook class is complex and I don't wan't to 
maintain duplicate documentation I request that you review the 
L<documentation|Spreadsheet::XLSX::Reader::LibXML> for that class 
there.  This attribute can/should only be set at ->new, however, it 
delegates to this class a number of methods that will update the workbook 
instance and therefore have universal effect when the other sheets are read.

B<Default:> none

B<Required:> yes

B<attribute methods> Methods provided to adjust this attribute
		
=over

B<counting_from_zero> - delegated

B<boundary_flag_setting> - delegated

B<change_boundary_flag> - delegated

B<_has_shared_strings_file> - delegated

B<get_shared_string_position> - delegated

B<_has_styles_file> - delegated

B<get_format_position> - delegated

B<set_empty_is_end> - delegated

B<is_empty_the_end> - delegated

B<_starts_at_the_edge> - delegated

B<get_group_return_type> - delegated

B<set_group_return_type> - delegated

B<get_epoch_year> - delegated

B<change_output_encoding> - delegated

B<get_date_behavior> - delegated

B<set_date_behavior> - delegated

B<get_empty_return_type> - delegated

B<set_error> - delegated

B<set_values_only> - delegated

B<get_values_only> - delegated

=back

=back

=head1 SUPPORT

=over

L<github Spreadsheet::XLSX::Reader::LibXML/issues
|https://github.com/jandrew/Spreadsheet-XLSX-Reader-LibXML/issues>

B<1.> Add the workbook attributute to the documentation

=back

=head1 TODO

=over

B<1.> Eliminate the min / max row / col calls from this role 
(and requireds) if possible.  

=back

=head1 AUTHOR

=over

=item Jed Lund

=item jandrew@cpan.org

=back

=head1 COPYRIGHT

This program is free software; you can redistribute
it and/or modify it under the same terms as Perl itself.

The full text of the license can be found in the
LICENSE file included with this module.

This software is copyrighted (c) 2014, 2015 by Jed Lund

=head1 DEPENDENCIES

=over

L<version> - 0.77

L<Type::Tiny> - 1.000

L<Spreadsheet::XLSX::Reader::LibXML::Cell>

L<Moose::Role>

=over

B<requires>

=over

min_row

max_row

min_col

max_col

row_range

col_range

_get_col_row

_get_next_value_cell

_get_row_all

=back

=back

=back

=head1 SEE ALSO

=over

L<Spreadsheet::ParseExcel> - Excel 2003 and earlier

L<Spreadsheet::XLSX> - 2007+

L<Spreadsheet::ParseXLSX> - 2007+

L<Log::Shiras|https://github.com/jandrew/Log-Shiras>

=over

All lines in this package that use Log::Shiras are commented out

=back

=back

=cut

#########1#########2 main pod documentation end  5#########6#########7#########8#########9