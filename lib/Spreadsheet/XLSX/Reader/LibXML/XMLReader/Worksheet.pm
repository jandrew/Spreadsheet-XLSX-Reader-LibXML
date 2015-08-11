package Spreadsheet::XLSX::Reader::LibXML::XMLReader::Worksheet;
use version; our $VERSION = qv('v0.38.10');
###LogSD	warn "You uncovered internal logging statements for Spreadsheet::XLSX::Reader::LibXML::XMLReader::Worksheet-$VERSION";

use	5.010;
use	Moose;
use	MooseX::StrictConstructor;
use	MooseX::HasDefaults::RO;
use Carp qw( confess );
use Types::Standard qw(
		Int				Str				ArrayRef
		HashRef			HasMethods		Bool
		Enum
    );
use lib	'../../../../../../lib';
###LogSD	use Log::Shiras::Telephone;
###LogSD	use Log::Shiras::UnhideDebug;
extends	'Spreadsheet::XLSX::Reader::LibXML::XMLReader';

#########1 Dispatch Tables & Package Variables    5#########6#########7#########8#########9

my	$cell_name_translation = {
		f => 'cell_formula',
		v => 'raw_value',
	};

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
	
has workbook_instance =>(
		isa		=> HasMethods[qw(
						counting_from_zero			boundary_flag_setting
						change_boundary_flag		_has_shared_strings_file
						get_shared_string_position	_has_styles_file
						get_format_position			set_empty_is_end
						is_empty_the_end			_starts_at_the_edge
						get_group_return_type		set_group_return_type
						get_epoch_year				change_output_encoding
						get_date_behavior			set_date_behavior
						get_empty_return_type		set_error
						get_values_only				set_values_only
						parse_excel_format_string
					)],
		handles	=> [qw(
						counting_from_zero			boundary_flag_setting
						change_boundary_flag		_has_shared_strings_file
						get_shared_string_position	_has_styles_file
						get_format_position			set_empty_is_end
						is_empty_the_end			_starts_at_the_edge
						get_group_return_type		set_group_return_type
						get_epoch_year				change_output_encoding
						get_date_behavior			set_date_behavior
						get_empty_return_type		set_error
						get_values_only				set_values_only
						parse_excel_format_string
					)],
		required => 1,
	);
###LogSD	use Log::Shiras::UnhideDebug;
with	'Spreadsheet::XLSX::Reader::LibXML::CellToColumnRow',
		'Spreadsheet::XLSX::Reader::LibXML::XMLReader::XMLToPerlData',
		;
with	'Spreadsheet::XLSX::Reader::LibXML::GetCell';

#########1 Public Methods     3#########4#########5#########6#########7#########8#########9

sub min_row{
	my( $self ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD					name_space 	=> ($self->get_log_space .  '::row_bound::min_row' ), );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"Returning the minimum row: " . $self->_min_row ] );
	return $self->_min_row;
}

sub max_row{
	my( $self ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD					name_space 	=> ($self->get_log_space .  '::row_bound::max_row' ), );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"Returning the maximum row: " . $self->_max_row ] );
	return $self->_max_row;
}

sub min_col{
	my( $self ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD					name_space 	=> ($self->get_log_space .  '::row_bound::min_col' ), );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"Returning the minimum column: " . $self->_min_col ] );
	return $self->_min_col;
}

sub max_col{
	my( $self ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD					name_space 	=> ($self->get_log_space .  '::row_bound::max_col' ), );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"Returning the maximum column: " . $self->_max_col ] );
	return $self->_max_col;
}

sub row_range{
	my( $self ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD					name_space 	=> ($self->get_log_space .  '::row_bound::row_range' ), );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"Returning row range( " . $self->_min_row . ", " . $self->_max_row . " )" ] );
	return( $self->_min_row, $self->_max_row );
}

sub col_range{
	my( $self ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD					name_space 	=> ($self->get_log_space .  '::row_bound::col_range' ), );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"Returning col range( " . $self->_min_col . ", " . $self->_max_col . " )" ] );
	return( $self->_min_col, $self->_max_col );
}


#########1 Private Attributes 3#########4#########5#########6#########7#########8#########9

has _sheet_min_col =>(
		isa			=> Int,
		writer		=> '_set_min_col',
		reader		=> '_min_col',
		predicate	=> 'has_min_col',
	);

has _sheet_min_row =>(
		isa			=> Int,
		writer		=> '_set_min_row',
		reader		=> '_min_row',
		predicate	=> 'has_min_row',
	);

has _sheet_max_col =>(
		isa			=> Int,
		writer		=> '_set_max_col',
		reader		=> '_max_col',
		predicate	=> 'has_max_col',
	);

has _sheet_max_row =>(
		isa			=> Int,
		writer		=> '_set_max_row',
		reader		=> '_max_row',
		predicate	=> 'has_max_row',
	);

has _last_row_col =>(
		isa			=> ArrayRef[Int],
		reader		=> '_get_last_row_col',
		writer		=> '_set_last_row_col',
		predicate	=> '_has_last_row_col',
	);

has _last_cell_ref =>(
		isa			=> HashRef,
		reader		=> '_get_last_cell_ref',
		writer		=> '_set_last_cell_ref',
		clearer		=> '_clear_last_cell_ref',
		predicate	=> '_has_last_cell_ref',
	);

has _next_row_col =>(
		isa			=> ArrayRef[Int],
		reader		=> '_get_next_row_col',
		writer		=> '_set_next_row_col',
		predicate	=> '_has_next_row_col',
	);

has _next_cell_ref =>(
		isa			=> HashRef,
		reader		=> '_get_next_cell_ref',
		writer		=> '_set_next_cell_ref',
		clearer		=> '_clear_next_cell_ref',
		predicate	=> '_has_next_cell_ref',
	);

has	_merge_map =>(
		isa		=> ArrayRef,
		traits	=> ['Array'],
		writer	=> '_set_merge_map',
		handles	=>{
			_get_row_merge_map => 'get',
		},
	);

has _reported_col =>(
		isa			=> Int,
		writer		=> '_set_reported_col',
		reader		=> '_get_reported_col',
	);

has _reported_row =>(
		isa			=> Int,
		writer		=> '_set_reported_row',
		reader		=> '_get_reported_row',
	);

#########1 Private Methods    3#########4#########5#########6#########7#########8#########9

sub _load_unique_bits{
	my( $self, ) = @_;#, $new_file, $old_file
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD					name_space 	=> ($self->get_log_space . '::_load_unique_bits' ), );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"Setting the Worksheet unique bits", "Byte position: " . $self->byte_consumed ] );
	
	# Read the sheet dimensions
	if( $self->next_element( 'dimension' ) ){
		my $dimension = $self->parse_element;
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"parsed dimension value:", $dimension ] );
		my	( $start, $end ) = split( /:/, $dimension->{ref} );
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"Start position: $start", 
		###LogSD		( $end ? "End position: $end" : '' ), "Byte position: " . $self->byte_consumed ] );
		my ( $start_column, $start_row ) = ( $self->_starts_at_the_edge ) ?
												( 1, 1 ) : $self->_parse_column_row( $start );
		my ( $end_column, $end_row	) = $end ? 
				$self->_parse_column_row( $end ) : 
				( $start_column, $start_row ) ;
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"Start column: $start_column", "Start row: $start_row",
		###LogSD		"End column: $end_column", "End row: $end_row" ] );
		$self->_set_min_col( $start_column );
		$self->_set_min_row( $start_row );
		$self->_set_max_col( $end_column );
		$self->_set_max_row( $end_row );
		$self->_set_last_row_col( [$start_row, ($start_column - 1)] );
		$self->_clear_last_cell_ref;
		$self->_set_next_row_col( [$start_row, ($start_column - 1)] );
		$self->_clear_next_cell_ref;
		$self->_set_reported_row( $start_row );
		$self->_set_reported_col( $start_column - 1 );
	}else{
		confess "No sheet dimensions provided";# Shouldn't the error instance be loaded already?
	}
	
	#build a merge map
	my	$merge_ref = [];
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"Loading the mergeCell" ] );
	while( $self->node_name eq 'mergeCell' or $self->next_element('mergeCell') ){
		my $merge_range = $self->parse_element;
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"parsed merge element to:", $merge_range ] );
		my ( $start, $end ) = split /:/, $merge_range->{ref};
		my ( $start_col, $start_row ) = $self->_parse_column_row( $start );
		my ( $end_col, $end_row ) = $self->_parse_column_row( $end );
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"Start column: $start_col", "Start row: $start_row",
		###LogSD		"End column: $end_col", "End row: $end_row" ] );
		my 	$min_col = $start_col;
		while ( $start_row <= $end_row ){
			$merge_ref->[$start_row]->[$start_col] = $merge_range->{ref};
			$start_col++;
			if( $start_col > $end_col ){
				$start_col = $min_col;
				$start_row++;
			}
		}
	}
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"Final merge ref:", $merge_ref ] );
	$self->_set_merge_map( $merge_ref );
	$self->start_the_file_over;
	return 1;
}

sub _get_next_value_cell{
	my( $self, ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD					name_space 	=> ($self->get_log_space . '::_get_next_value_cell' ), );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			'Loading the next cell with value after [row, column]: [' .
	###LogSD			join( ', ', @{$self->_get_next_row_col} ) . ']'] );
	my	$result = 1;
		$result = $self->next_element( 'c' ) if !$self->node_name or $self->node_name ne 'c';
	my	$sub_ref = 'EOF';
	if( !$result ){
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD			'Reached the end of the file',] );
		$self->start_the_file_over;
	}else{
		$sub_ref = undef;
		CHECKVALUECELLS: while( !$sub_ref or $self->get_values_only ){
			$sub_ref = $self->parse_element;
			@$sub_ref{qw( col row )} = $self->_parse_column_row( $sub_ref->{r} );
			#~ if( exists $sub_ref->{t} and exists $sub_ref->{s} ){# special old magic for cell formatting from the formula bar
				#~ $sub_ref->{f}->{raw_text} =
					#~ $sub_ref->{s} eq '1' ? "'" :
					#~ $sub_ref->{s} eq '2' ? "^" :
					#~ $sub_ref->{s} eq '3' ? '"' : undef ;
				#~ delete $sub_ref->{s};
				#~ delete $sub_ref->{f}->{raw_text} if !$sub_ref->{f}->{raw_text};
			#~ }
			###LogSD	$phone->talk( level => 'trace', message => [
			###LogSD		'The next cell with data is:', $sub_ref,] );
			if( exists $sub_ref->{v} or !$self->get_values_only ){### Other text or value call outs here?
				###LogSD	$phone->talk( level => 'trace', message => [
				###LogSD		'Found a cell with a value no additional work is needed' ,] );
				last CHECKVALUECELLS;
			}
			$result = $self->next_element( 'c' ) if !$self->node_name or $self->node_name ne 'c';
			if( !$result ){
				###LogSD	$phone->talk( level => 'debug', message => [
				###LogSD			'Reached the end of the file',] );
				$sub_ref = 'EOF';
				$self->start_the_file_over;
				last CHECKVALUECELLS;
			}
		}
	}
	
	#Add merge value
	if( $sub_ref and ref( $sub_ref ) ){
		my $merge_row = $self->_get_row_merge_map( $sub_ref->{row} );
		if( ref( $merge_row ) and $merge_row->[$sub_ref->{col}] ){
			$sub_ref->{cell_merge} = $merge_row->[$sub_ref->{col}];
		}
	}
	###LogSD	$phone->talk( level => 'trace', message => [
	###LogSD		'Ref to this point:', $sub_ref,] );
	
	# move current to prior
	if( $self->_has_next_cell_ref ){
		###LogSD	$phone->talk( level => 'trace', message => [
		###LogSD		'Saving the next ref as the last ref:', $self->_get_next_cell_ref,] );
		$self->_set_last_cell_ref( $self->_get_next_cell_ref );
		$self->_set_last_row_col( $self->_get_next_row_col );
	}
	
	#load current
	if( ref $sub_ref ){
		###LogSD	$phone->talk( level => 'trace', message => [
		###LogSD		'Saving the next ref data in the attributes:', $sub_ref] );
		$self->_set_next_cell_ref( $sub_ref );
		$self->_set_next_row_col( [ @$sub_ref{qw( row col )} ] );
		$self->_set_reported_row( $sub_ref->{row} );
		$self->_set_reported_col( $sub_ref->{col} );
	}else{
		###LogSD	$phone->talk( level => 'trace', message => [
		###LogSD		'Past the EOF so saving the empty ref and position [row, col]: [' .
		###LogSD		($self->_max_row + 1) . ', ' . $self->_min_col . ']',, caller(1)] );
		$self->_clear_next_cell_ref;
		$self->_set_next_row_col( [($self->_max_row + 1), $self->_min_col ] );
		$self->_set_reported_row( $self->_min_row );
		$self->_set_reported_col( $self->_min_col - 1 );
	}
	
	return $sub_ref;
}

sub _get_next_cell{
	my( $self, ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD					name_space 	=> ($self->get_log_space . '::_get_next_cell' ), );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			'Loading the next cell after [row, column]: [' . $self->_get_reported_row .
	###LogSD			', ' . $self->_get_reported_col . ']',] );
	my ( $next_row, $next_col ) = @{$self->_get_next_row_col};
	my	$target_row = $self->_get_reported_row;
	my	$target_col = $self->_get_reported_col + 1;
	if( $target_col > $self->_max_col ){
		$target_row++;
		$target_col = $self->_min_col;
	}
	# check if an index reset is needed (transition case from a different parsing method)
	if(	$target_row < $self->_get_last_row_col->[0] ){
		###LogSD	$phone->talk( level => 'debug', message =>[
		###LogSD		'Starting the sheet over' ] );
		$self->_set_next_row_col( [ $self->_min_row, ($self->_min_col - 1) ] );
		$self->_set_last_row_col( [ @{$self->_get_next_row_col} ] );
		$self->_clear_last_cell_ref;
		( $next_row, $next_col ) = @{$self->_get_next_row_col} ;
	}
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		'Searching for [row, column]: [' . $target_row . ', ' . $target_col . ']',] );
	my	$result = 'NoAdvance';
	if( $target_row <= $self->_max_row ){
		while( ( $target_row > $next_row ) or
				( $target_row == $next_row and $target_col > $next_col ) ){
			$result = $self->_get_next_value_cell;
			if( $result eq 'EOF' ){
				( $next_row, $next_col ) = ( ($self->_max_row + 1), $self->_min_row );
				last;
			}
			( $next_row, $next_col ) = @$result{qw( row col )};
		}
	}
	$self->_set_reported_row( $target_row );
	$self->_set_reported_col( $target_col );
	###LogSD	$phone->talk( level => 'debug', message =>[ 'Advanced to:', $result ] );
	
	# check for EOF and empty cells(no EOR in a _next_xxx scenario) 
	if(	$target_row > $self->_max_row or # Maximum EOF
		$self->is_empty_the_end and $result eq 'EOF' ){ # Stop when empty EOF
		###LogSD	$phone->talk( level => 'debug', message =>[
		###LogSD		'Reached the end of EOF state for target [row, col]: [' .
		###LogSD		$target_row . ', ' . $target_col . ']', '..or got an earl EOF' ] );
		$self->_set_reported_row( $self->_min_row );
		$self->_set_reported_col( $self->_min_col - 1 );
		$self->_set_next_row_col( [ $self->_min_row, ($self->_min_col - 1) ] );
		return 'EOF';
	}elsif( !$self->is_empty_the_end and $next_row > $target_row ){
		###LogSD	$phone->talk( level => 'debug', message =>[
		###LogSD		'Found an empty cell at the end of the row for [row, col]: [' .
		###LogSD		$target_row . ', ' . $target_col . ']' ] );
		return undef;
	}elsif( $self->is_empty_the_end and $next_row > $target_row ){
		###LogSD	$phone->talk( level => 'debug', message =>[
		###LogSD		'Wrapping the row for [row, col]: [' . $target_row . ', ' . $target_col . ']' ] );
		$target_row = $self->_set_reported_row( $target_row + 1 );
		$target_col = $self->_set_reported_col( $self->_min_col );
		if( $next_row == ($target_row) and $next_col == $self->_min_col ){
			###LogSD	$phone->talk( level => 'debug', message =>[
			###LogSD		'Found a result at [row, col]: [' . $target_row . ', ' . $target_col . ']' ] );
			return $self->_get_next_cell_ref;
		}else{
			###LogSD	$phone->talk( level => 'debug', message =>[
			###LogSD		'Found an empty cell at the beginning of the next row' ] );
			return undef;
		}
	}elsif( $next_row == $target_row and $next_col > $target_col ){
		###LogSD	$phone->talk( level => 'debug', message =>[
		###LogSD		'Found and empty cell at [row, col]: [' . $target_row . ', ' . $target_col . ']' ] );
		return undef;
	}elsif( $result eq 'NoAdvance' ){
		###LogSD	$phone->talk( level => 'debug', message =>[
		###LogSD		'Retreiving the stored value for [row, col]: [' . $target_row . ', ' . $target_col . ']' ] );
		$result = $self->_get_next_cell_ref;
	}		
	###LogSD	$phone->talk( level => 'debug', message =>[
	###LogSD		'Found a result at [row, col]: [' . $target_row . ', ' . $target_col . ']' ] );
	return $result;
}

sub _get_col_row{
	my( $self, $column, $row ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD					name_space 	=> ($self->get_log_space . '::_get_col_row' ), );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			'Getting [column, row]: [' . (($column) ? $column : undef) .
	###LogSD			', ' . (($row) ? $row : undef) . ']',] );
	
	# Validate
	if( !$column or !$row ){
		$self->set_error( "Missing either a passed column or row" );
		return undef;
	}
	
	# See if you went too far
	if( $row > $self->_max_row ){
		###LogSD	$phone->talk( level => 'debug', message =>[
		###LogSD		"Requested row -$row- is greater than the max row: " . $self->_max_row ] );
		return 'EOF';
	}
	if( $column > $self->_max_col ){
		###LogSD	$phone->talk( level => 'debug', message =>[
		###LogSD		"Requested column -$column- is greater than the max column: " . $self->_max_col ] );
		return (($row == $self->_max_row) ? 'EOF' : 'EOR');
	}
	
	# check if an index reset is needed
	my	$result = 'NoAdvance';
	if(	$row < $self->_get_last_row_col->[0] or
		$row == $self->_get_last_row_col->[0] and $column < $self->_get_last_row_col->[1]){
		###LogSD	$phone->talk( level => 'debug', message =>[
		###LogSD		'Starting the sheet over' ] );
		$self->_set_next_row_col( [ $self->_min_row, ($self->_min_col - 1) ] );
		$self->_set_last_row_col( [ @{$self->_get_next_row_col} ] );
		$self->_clear_last_cell_ref;
		$self->_clear_next_cell_ref;
		$self->start_the_file_over;
		$self->_set_reported_row( $self->_min_row );
		$self->_set_reported_col( $self->_min_col - 1 );
	}
	my ( $next_row, $next_col ) = @{$self->_get_next_row_col};
	
	# Move to bracket the target value
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		'Searching for [column, row]: [' . $column . ', ' . $row . ']',] );
	while( ( $row > $next_row ) or
			( $row == $next_row and $column > $next_col ) ){
		$result = $self->_get_next_value_cell;
		if( $result eq 'EOF' ){
			( $next_row, $next_col ) = ( ($self->_max_row + 1), $self->_min_row );
			last;
		}
		( $next_row, $next_col ) = @$result{qw( row col )};
	}
	$self->_set_reported_row( $row );
	$self->_set_reported_col( $column );
	###LogSD	$phone->talk( level => 'debug', message =>[ 'Advanced to:', $result ] );
	
	# check for EOF, EOR, and empty cells
	if(	$row == $next_row and $column == $next_col ){
		###LogSD	$phone->talk( level => 'debug', message =>[
		###LogSD		'Found data at (next) [row, col]: [' .
		###LogSD		$row . ', ' . $column . ']', ] );
		return $self->_get_next_cell_ref;
	}elsif( $row == $self->_get_last_row_col->[0] and $column == $self->_get_last_row_col->[1] ){
		###LogSD	$phone->talk( level => 'debug', message =>[
		###LogSD		'Found data at the (last) [row, col]: [' .
		###LogSD		$row . ', ' . $column . ']', ] );
		$self->_set_reported_row( $self->_get_last_row_col->[0] );
		$self->_set_reported_col( $self->_get_last_row_col->[1] );
		return $self->_get_last_cell_ref;
	}elsif( $self->is_empty_the_end and $next_row > $self->_max_row ){
		###LogSD	$phone->talk( level => 'debug', message =>[
		###LogSD		'Reached the end of file (EOF) state for target [row, col]: [' .
		###LogSD		$row . ', ' . $column . ']', ] );
		$self->_set_reported_row( $self->_min_row );
		$self->_set_reported_col( $self->_min_col - 1 );
		$self->_set_next_row_col( [ $self->_min_row, ($self->_min_col - 1) ] );
		return 'EOF';
	}elsif( $self->is_empty_the_end and $next_row > $row ){
		###LogSD	$phone->talk( level => 'debug', message =>[
		###LogSD		'Reached the end of row (EOR) state for target [row, col]: [' .
		###LogSD		$row . ', ' . $column . ']', ] );
		return 'EOR';
	}
	###LogSD	$phone->talk( level => 'debug', message =>[
	###LogSD		'Found an empty cell at [row, col]: [' . $row . ', ' . $column . ']' ] );
	return undef;
}

sub _get_row_all{
	my( $self, $row ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD					name_space 	=> ($self->get_log_space . '::_get_row_all' ), );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			'Getting row: ' . (($row) ? $row : '') ] );
	
	# Get next row as needed
	if( !$row ){
		my $last_col = $self->_get_reported_col;
		###LogSD	$phone->talk( level => 'info', message =>[
		###LogSD		"No row requested - Determining what it should be with the last column: $last_col", ] );
		$row = $self->_get_reported_row + !!$last_col;
		###LogSD	$phone->talk( level => 'info', message =>[
		###LogSD		"Now requesting Excel row: $row", ] );
	}
	
	# See if you went too far
	if( $row > $self->_max_row ){
		###LogSD	$phone->talk( level => 'debug', message =>[
		###LogSD		"Requested row -$row- is greater than the max row: " . $self->_max_row ] );
		$self->_set_next_row_col( [ $self->_min_row, ($self->_min_col - 1) ] );
		$self->_set_last_row_col( [ @{$self->_get_next_row_col} ] );
		$self->_clear_last_cell_ref;
		$self->_clear_next_cell_ref;
		$self->start_the_file_over;
		$self->_set_reported_row( $self->_min_row );
		$self->_set_reported_col( $self->_min_col - 1 );
		return 'EOF';
	}
	
	# check if an index reset is needed
	my	$result = 'NoAdvance';
	if(	$row < ($self->_get_last_row_col->[0] - 1 )  ){
		###LogSD	$phone->talk( level => 'debug', message =>[
		###LogSD		'Starting the sheet over' ] );
		$self->_set_next_row_col( [ $self->_min_row, ($self->_min_col - 1) ] );
		$self->_set_last_row_col( [ @{$self->_get_next_row_col} ] );
		$self->_clear_last_cell_ref;
		$self->_clear_next_cell_ref;
		$self->start_the_file_over;
		$self->_set_reported_row( $self->_min_row );
		$self->_set_reported_col( $self->_min_col - 1 );
	}
	my ( $next_row, $next_col ) = @{$self->_get_next_row_col};
	
	# Move to bracket the target value
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"Searching for row: $row", "..against next row: $next_row" ] );
	while( $row > $next_row ){
		$result = $self->_get_next_value_cell;
		if( $result eq 'EOF' ){
			( $next_row, $next_col ) = ( ($self->_max_row + 1), $self->_min_row );
			last;
		}
		( $next_row, $next_col ) = @$result{qw( row col )};
	}
	###LogSD	$phone->talk( level => 'debug', message =>[ 'Advanced to:', $result ] );
	
	# check for EOF and empty rows
	if( $row > $self->_get_last_row_col->[0] and
		$row < $self->_get_next_row_col->[0]		){
		###LogSD	$phone->talk( level => 'debug', message =>[
		###LogSD		"Found an empty row at: $row", 'Need to determine if it is is an EOF', ] );
		$self->_set_reported_col( $self->_max_col );
		$self->_set_reported_row( $row );
		if( $self->_get_next_row_col->[0] > $self->_max_row ){
			###LogSD	$phone->talk( level => 'debug', message =>[
			###LogSD		'Found a weird end of file before max row', ] );
			$self->_set_next_row_col( [ $self->_min_row, ($self->_min_col - 1) ] );
			$self->_set_last_row_col( [ @{$self->_get_next_row_col} ] );
			$self->_clear_last_cell_ref;
			$self->_clear_next_cell_ref;
			$self->start_the_file_over;
			$self->_set_reported_row( $self->_min_row );
			$self->_set_reported_col( $self->_min_col - 1 );
			return 'EOF';
		}elsif( $self->is_empty_the_end ){
			###LogSD	$phone->talk( level => 'debug', message =>[
			###LogSD		"Don't fill in empty stuff", ] );
			return [];
		}else{
			my $answer;
			for( $self->_min_col .. $self->_max_col ){
				push @$answer, undef;
			}
			return $answer;
		}
	}
	
	# build the row ref
	my	$column = $self->_min_col;
		$result = undef;
	my	$x = 0;
	my	$answer_ref = [];
	while( $x < 17000 ){ #Excel 2013 goes to 16,384 columns
		$result = $self->_get_col_row( $column, $row );
		###LogSD	$phone->talk( level => 'debug', message =>[
		###LogSD		'collecting the data at [row, col]: [' .
		###LogSD		$row . ', ' . $column . ']', '..with result:', $result ] );
		last if ($result and ($result eq 'EOR' or $result eq 'EOF'));
		push @$answer_ref, $result;
		$column++;
		$x++;
	}
	$self->_set_reported_row( $row );
	$self->_set_reported_col( $column );
	
	###LogSD	$phone->talk( level => 'debug', message =>[
	###LogSD		'Final answer:', $answer_ref ] );
	return $answer_ref;
}

#########1 Phinish            3#########4#########5#########6#########7#########8#########9

no Moose;
__PACKAGE__->meta->make_immutable;
	
1;

#########1 Documentation      3#########4#########5#########6#########7#########8#########9
__END__

=head1 NAME

Spreadsheet::XLSX::Reader::LibXML::XMLReader::Worksheet - A LibXML::XMLReader worksheet base class

=head1 SYNOPSIS

	use Data::Dumper;
	use MooseX::ShortCut::BuildInstance qw( build_instance );
	use Types::Standard qw( Bool HasMethods );
	use Spreadsheet::XLSX::Reader::LibXML::Error;
	use Spreadsheet::XLSX::Reader::LibXML::XMLReader::Worksheet;
	my  $error_instance = Spreadsheet::XLSX::Reader::LibXML::Error->new( should_warn => 0 );
	my  $workbook_instance = build_instance(
	        package     => 'WorkbookInstance',
	        add_methods =>{
	            counting_from_zero         => sub{ return 0 },
	            boundary_flag_setting      => sub{},
	            change_boundary_flag       => sub{},
	            _has_shared_strings_file   => sub{ return 1 },
				get_shared_string_position => sub{},
				_has_styles_file           => sub{},
	            get_format_position        => sub{},
	            get_group_return_type      => sub{},
	            set_group_return_type      => sub{},
	            get_epoch_year             => sub{ return '1904' },
	            change_output_encoding     => sub{ $_[0] },
	            get_date_behavior          => sub{},
	            set_date_behavior          => sub{},
	            get_empty_return_type      => sub{ return 'undef_string' },
	            get_values_only            => sub{},
	            set_values_only            => sub{},
	        },
	        add_attributes =>{
	            error_inst =>{
	                isa => 	HasMethods[qw(
	                            error set_error clear_error set_warnings if_warn
	                        ) ],
	                clearer  => '_clear_error_inst',
	                reader   => 'get_error_inst',
	                required => 1,
				    handles =>[ qw(
						error set_error clear_error set_warnings if_warn
					) ],
	            },
	            empty_is_end =>{
	                isa     => Bool,
					writer  => 'set_empty_is_end',
	                reader  => 'is_empty_the_end',
	                default => 0,
	            },
	            from_the_edge =>{
	                isa     => Bool,
					reader  => '_starts_at_the_edge',
	                writer  => 'set_from_the_edge',
					default => 1,
	            },
			},
			error_inst => $error_instance,
	    );
	my $test_instance = Spreadsheet::XLSX::Reader::LibXML::XMLReader::Worksheet->new(
	        file              => 'xl/worksheets/sheet3.xml',
	        error_inst        => $error_instance,
	        sheet_name        => 'Sheet3',
	        workbook_instance => $workbook_instance,
	    );
	my $x = 0;
	my $result;
	while( $x < 20 and (!$result or $result ne 'EOF') ){
	    $result = $test_instance->_get_next_value_cell;
	    print "Collecting data from position: $x" . Dumper( $result );
		$x++;
	}

	###############################################
	# SYNOPSIS Screen Output
	# 01: Collecting data from position: 0
	# 02: $VAR1 = {
	# 03:           'r' => 'A2',
	# 04:           'row' => 2,
	# 05:           'col' => 1,
	# 06:           'v' => {
	# 07:                  'raw_text' => '0'
	# 08:                },
	# 09:           't' => 's'
	# 10:         };
	# 11: 
	# 12: Collecting data from position: 1
	# 13: $VAR1 = {
	# 14:           'r' => 'D2',
	# 15:           'row' => 2,
	# 16:           'col' => 4,
	# 17:           'v' => {
	# 18:                  'raw_text' => '2'
	# 19:                },
	# 20:           't' => 's'
	# 21:         };
	#		~~ Continuing ~~
	###############################################
	
    
=head1 DESCRIPTION

This documentation is written to explain ways to use this module when writing your own excel 
parser.  To use the general package for excel parsing out of the box please review the 
documentation for L<Workbooks|Spreadsheet::XLSX::Reader::LibXML>,
L<Worksheets|Spreadsheet::XLSX::Reader::LibXML::Worksheet>, and 
L<Cells|Spreadsheet::XLSX::Reader::LibXML::Cell>

This module provides the basic connection to individual worksheet files for parsing xlsx 
workbooks.  It does not provide a way to connect to L<chartsheets|Spreadsheet::XLSX::Reader::LibXML::Chartsheet>.  
It does not provide the final view of a given cell.  The final view of the cell is collated 
with the role L<Spreadsheet::XLSX::Reader::LibXML::GetCell>.  This reader extends the base 
reader class L<Spreadsheet::XLSX::Reader::LibXML::XMLReader>.  The functionality provided 
by those modules is not covered here.

Modification of this module probably means extending a different reader or using other roles 
for implementation of the class.  See lines 18 and on in the code here for the location to 
change and See line 54 in the code L<Spreadsheet::XLSX::Reader::LibXML> for the way to repoint 
the package at a new module.

=head2 Attributes

Data passed to new when creating an instance.  For access to the values in these 
attributes see the listed 'attribute methods'. For general information on attributes see 
L<Moose::Manual::Attributes>.  For ways to manage the instance when opened see the 
L<Public Methods|/Public Methods>.  The remaining undocumented attributes are used internally 
for tracking state.
	
=head3 sheet_type

=over

B<Definition:> This will always be 'worksheet' for this module.  It is provided as 
a simple introspection method for distinguishing between worksheets and chartsheets 
in case the circumstances are ambiguous.

B<Default:> 'worksheet'

B<Range:> 'worksheet'

B<attribute methods> Methods provided to adjust this attribute
		
=over

B<get_sheet_type>

=over

B<Definition:> return the attribute value

=back

=back

=back
	
=head3 sheet_rel_id

=over

B<Definition:> To coordinate information accross the various sub-files excel maintains 
a relId for sheets.  This is the value that excel assigned to this sheet.

B<Range:> a string

B<attribute methods> Methods provided to adjust this attribute
		
=over

B<rel_id>

=over

B<Definition:> return the attribute value

=back

=back

=back
	
=head3 sheet_id

=over

B<Definition:> When writing vbScript the sheet can be identified by a number instead 
of a name.  This is that number.

B<Range:> an integer

B<attribute methods> Methods provided to adjust this attribute
		
=over

B<sheet_id>

=over

B<Definition:> return the attribute value

=back

=back

=back
	
=head3 sheet_position

=over

B<Definition:> Even if there are chartsheets in the workbook you will only get a list 
of worksheets when using 'worksheet' methods.  However, this position is the visible 
position of the worksheet in the workbook including chartsheets.  This can be different 
than L<sheet_id|/sheet_id>

B<Range:> an integer

B<attribute methods> Methods provided to adjust this attribute
		
=over

B<position>

=over

B<Definition:> return the attribute value

=back

=back

=back
	
=head3 sheet_name

=over

B<Definition:> This is the visible string expressed on the tab of the 
worksheet in the workbook.

B<Range:> a String

B<attribute methods> Methods provided to adjust this attribute
		
=over

B<get_name>

=over

B<Definition:> return the attribute value

=back

=back

=back
	
=head3 _sheet_min_col

=over

B<Definition:> This is the minimum column in the sheet with data or formatting.  For this 
module it is pulled from the xml file at worksheet/dimension:ref = "upperleft:lowerright"

B<Range:> an integer

B<attribute methods> Methods provided to adjust this attribute
		
=over

B<_set_min_col>

=over

B<Definition:> sets the attribute value

=back

B<_min_col>

=over

B<Definition:> returns the attribute value

=back

B<has_min_col>

=over

B<Definition:> attribute predicate

=back

=back

=back
	
=head3 _sheet_min_row

=over

B<Definition:> This is the minimum row in the sheet with data or formatting.  For this 
module it is pulled from the xml file at worksheet/dimension:ref = "upperleft:lowerright"

B<Range:> an integer

B<attribute methods> Methods provided to adjust this attribute
		
=over

B<_set_min_row>

=over

B<Definition:> sets the attribute value

=back

B<_min_row>

=over

B<Definition:> returns the attribute value

=back

B<has_min_row>

=over

B<Definition:> attribute predicate

=back

=back

=back
	
=head3 _sheet_max_col

=over

B<Definition:> This is the maximum column in the sheet with data or formatting.  For this 
module it is pulled from the xml file at worksheet/dimension:ref = "upperleft:lowerright"

B<Range:> an integer

B<attribute methods> Methods provided to adjust this attribute
		
=over

B<_set_max_col>

=over

B<Definition:> sets the attribute value

=back

B<_max_col>

=over

B<Definition:> returns the attribute value

=back

B<has_max_col>

=over

B<Definition:> attribute predicate

=back

=back

=back
	
=head3 _sheet_max_row

=over

B<Definition:> This is the maximum row in the sheet with data or formatting.  For this 
module it is pulled from the xml file at worksheet/dimension:ref = "upperleft:lowerright"

B<Range:> an integer

B<attribute methods> Methods provided to adjust this attribute
		
=over

B<_set_max_row>

=over

B<Definition:> sets the attribute value

=back

B<_max_row>

=over

B<Definition:> returns the attribute value

=back

B<has_max_row>

=over

B<Definition:> attribute predicate

=back

=back

=back
	
=head3 _merge_map

=over

B<Definition:> This is an array ref of array refs where the first level represents rows 
and the second level of array represents cells.  If a cell is merged then the merge span 
is stored in the row sub array position.  This means the same span is stored in multiple 
positions.  The data is stored in the Excel convention of count from 1 so the first position 
in both levels of the array are essentially placeholders.  The data is extracted from the 
merge section of the worksheet at worksheet/mergeCells.  That array is read and converted 
into this format for reading by this module when it first opens the worksheet..

B<Range:> an array ref

B<attribute methods> Methods provided to adjust this attribute
		
=over

B<_set_merge_map>

=over

B<Definition:> sets the attribute value

=back

=back

B<delegated methods> This attribute uses the native trait 'Array'
		
=over

B<_get_row_merge_map( $int )> => 'get'

=over

B<Definition:> returns the sub array ref representing any merges for that 
row.  If no merges are available for that row it returns undef.

=back

=back

=back

=head2 Public Methods

These are the methods provided by this class for use by the end user.

=head3 min_row

=over

B<Definition:> This returns the minimum row with data or formatting in the worksheet.  It is 
separated from L<_min_row|/_sheet_min_row> so that the package can modify the output between the functions 
to match the attribute L<Spreadsheet::XLSX::Reader::LibXML/count_from_zero>

B<Accepts:> nothing

B<Returns:> the minimum row integer

=back

=head3 max_row

=over

B<Definition:> This returns the maximum row with data or formatting in the worksheet.  It is 
separated from L<_max_row|/_sheet_max_row> so that the package can modify the output between the functions 
to match the attribute L<Spreadsheet::XLSX::Reader::LibXML/count_from_zero>

B<Accepts:> nothing

B<Returns:> the maximum row integer

=back

=head3 min_col

=over

B<Definition:> This returns the minimum column with data or formatting in the worksheet.  It is 
separated from L<_min_col|/_sheet_min_col> so that the package can modify the output between the functions 
to match the attribute L<Spreadsheet::XLSX::Reader::LibXML/count_from_zero>

B<Accepts:> nothing

B<Returns:> the minimum column integer

=back

=head3 max_col

=over

B<Definition:> This returns the maximum column with data or formatting in the worksheet.  It is 
separated from L<_max_col|/_sheet_max_col> so that the package can modify the output between the functions 
to match the attribute L<Spreadsheet::XLSX::Reader::LibXML/count_from_zero>

B<Accepts:> nothing

B<Returns:> the maximum column integer

=back

=head3 row_range

=over

B<Definition:> This returns the first and last row with data or formatting in the worksheet.  It is 
separated from L<_min_row|/_sheet_min_row> and L<_max_row|/_sheet_max_row> so that the package can modify the 
output between the functions to match the attribute L<Spreadsheet::XLSX::Reader::LibXML/count_from_zero>

B<Accepts:> nothing

B<Returns:> ( $min_row, $max_row ) as a list

=back

=head3 col_range

=over

B<Definition:> This returns the first and last column  with data or formatting in the worksheet.  It is 
separated from L<_min_col|/_sheet_min_col> and L<_max_col|/_sheet_max_col> so that the package can modify the 
output between the functions to match the attribute L<Spreadsheet::XLSX::Reader::LibXML/count_from_zero>

B<Accepts:> nothing

B<Returns:> ( $min_col, $max_col ) as a list

=back

=head2 Private Methods

These are the methods provided by this class for use within the package but are not intended 
to be used by the end user.  Other private methods not listed here are used in the module but 
not used by the package.  If the private method is listed here then replacement of this module 
either requires replacing them or rewriting all the associated connecting roles and classes.

=head3 _load_unique_bits

=over

B<Definition:> This is called by L<Spreadsheet::XLSX::Reader::LibXML::XMLReader> when the file is 
loaded for the first time so that file specific data can be collected.  All the L<Attributes|/Attributes> 
with a leading _ in the documentation are filled in at this point.

B<Accepts:> nothing

B<Returns:> nothing

=back

=head3 _get_next_value_cell

=over

B<Definition:> This returns the worksheet file hash ref representation of the xml stored for the 
'next' value cell.  A cell is determined to have value based on the attribute 
L<Spreadsheet::XLSX::Reader::LibXML/values_only>.  Next is affected by the attribute 
L<Spreadsheet::XLSX::Reader::LibXML/empty_is_end>.  This method never returns an 'EOR' flag.  
It just wraps automatically.

B<Accepts:> nothing

B<Returns:> the cell (or value as requested)

=back

=head3 _get_next_cell

=over

B<Definition:> This returns on every cell position whether there is data or not.  For empty 
cells undef is returned.  Next is affected by the attribute L<Spreadsheet::XLSX::Reader::LibXML/empty_is_end>.  
This method never returns an 'EOR' flag.  It just wraps automatically.

B<Accepts:> nothing

B<Returns:> undef, a value, or the cell

=back

=head3 _get_col_row( $col, $row )

=over

B<Definition:> This is the way to return the information about a specific position in the worksheet.  
Since this is a private method it requires its inputs to be in the 'count from one' index.

B<Accepts:> ( $column, $row ) - both required in that order

B<Returns:> whatever is in that worksheet position as a hashref

=back

=head3 _get_row_all( $row )

=over

B<Definition:> This is returns an array ref of each of the values in the row placed in their 'count 
from one' position.  If the row is empty but it is not the end of the sheet then this will return an 
empty array ref.

B<Accepts:> ( $row ) - required

B<Returns:> an array ref

=back

=head1 SUPPORT

=over

L<github Spreadsheet::XLSX::Reader::LibXML/issues
|https://github.com/jandrew/Spreadsheet-XLSX-Reader-LibXML/issues>

=back

=head1 TODO

=over

B<1.> Nothing L<yet|/SUPPORT>

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

L<version>

L<perl 5.010|perl/5.10.0>

L<Moose>

L<MooseX::StrictConstructor>

L<MooseX::HasDefaults::RO>

L<Carp> - confess

L<Type::Tiny> - 1.000

L<Spreadsheet::XLSX::Reader::LibXML::XMLReader>

L<Spreadsheet::XLSX::Reader::LibXML::CellToColumnRow>

L<Spreadsheet::XLSX::Reader::LibXML::XMLReader::XMLToPerlData>

L<Spreadsheet::XLSX::Reader::LibXML::GetCell>

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

#########1 Documentation End  3#########4#########5#########6#########7#########8#########9
