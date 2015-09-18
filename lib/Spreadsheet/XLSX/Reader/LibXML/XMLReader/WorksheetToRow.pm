package Spreadsheet::XLSX::Reader::LibXML::XMLReader::WorksheetToRow;
use version; our $VERSION = qv('v0.38.14');
###LogSD	warn "You uncovered internal logging statements for Spreadsheet::XLSX::Reader::LibXML::XMLReader::WorksheetToRow-$VERSION";

use	5.010;
use	Moose;
use	MooseX::StrictConstructor;
use	MooseX::HasDefaults::RO;
use Clone 'clone';
use Carp qw( confess );
use Types::Standard qw(
		HasMethods		InstanceOf		ArrayRef
		Bool			Int				is_HashRef
    );
use MooseX::ShortCut::BuildInstance qw ( build_instance should_re_use_classes );
should_re_use_classes( 1 );
use lib	'../../../../../../lib';
###LogSD	use Log::Shiras::Telephone;
###LogSD	use Log::Shiras::UnhideDebug;
extends	'Spreadsheet::XLSX::Reader::LibXML::XMLReader';
use Spreadsheet::XLSX::Reader::LibXML::Row;

#########1 Dispatch Tables & Package Variables    5#########6#########7#########8#########9



#########1 Public Attributes  3#########4#########5#########6#########7#########8#########9

has is_hidden =>(
		isa		=> Bool,
		reader	=> 'is_sheet_hidden',
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
		'Spreadsheet::XLSX::Reader::LibXML::XMLToPerlData',
		;

#########1 Public Methods     3#########4#########5#########6#########7#########8#########9



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

has	_merge_map =>(
		isa		=> ArrayRef,
		traits	=> ['Array'],
		writer	=> '_set_merge_map',
		reader	=> '_get_merge_map',
		handles	=>{
			_get_row_merge_map => 'get',
		},
	);

has _column_formats =>(
		isa		=> ArrayRef,
		traits	=> ['Array'],
		writer	=> '_set_column_formats',
		reader	=> '_get_column_formats',
		default	=> sub{ [] },
		handles	=>{
			_get_custom_column_data => 'get',
		},
	);

has _old_row_inst =>(
		isa			=> InstanceOf[ 'Spreadsheet::XLSX::Reader::LibXML::Row' ],
		reader		=> '_get_old_row_inst',
		writer		=> '_set_old_row_inst',
		clearer		=> '_clear_old_row_inst',
		predicate	=> '_has_old_row_inst',
		handles	=>{
			_get_old_row_number 	=> 'get_row_number',
			_is_old_row_hidden		=> 'is_row_hidden',
			_get_old_row_formats	=> 'get_row_format', # pass the desired format key
			_get_old_column			=> 'get_the_column', # pass a column number (no next default) returns (cell|undef|EOR)
			#~ _get_old_value_column	=> 'get_the_next_value_column', # pass a column number (no next default) returns (cell|EOR)
			#~ _get_old_next_value		=> 'get_the_next_value_position', # pass nothing returns next (cell|EOR)
			_get_old_last_value_col	=> 'get_last_value_column',
			_get_old_row_list		=> 'get_row_all',
			_get_nold_row_end		=> 'get_row_end'
		},
	);

has _new_row_ref =>(
		isa			=> InstanceOf[ 'Spreadsheet::XLSX::Reader::LibXML::Row' ],
		reader		=> '_get_new_row_inst',
		writer		=> '_set_new_row_inst',
		clearer		=> '_clear_new_row_inst',
		predicate	=> '_has_new_row_inst',
		handles	=>{
			_get_new_row_number 	=> 'get_row_number',
			_is_new_row_hidden		=> 'is_row_hidden',
			_get_new_row_formats	=> 'get_row_format', # pass the desired format key
			_get_new_column			=> 'get_the_column', # pass a column number (no next default) returns (cell|undef|EOR)
			#~ _get_new_value_column	=> 'get_the_next_value_column', # pass a column number (no next default) returns (cell|EOR)
			_get_new_next_value		=> 'get_the_next_value_position', # pass nothing returns next (cell|EOR)
			_get_new_last_value_col	=> 'get_last_value_column',
			_get_new_row_list		=> 'get_row_all',
			_get_new_row_end		=> 'get_row_end'
		},
	);
	
has _row_hidden_states =>(
		isa		=> ArrayRef[ Bool ],
		traits	=>['Array'],
		default => sub{ [] },
		handles =>{
			_set_row_hidden => 'set',
			_get_row_hidden => 'get',
		},
	);

#########1 Private Methods    3#########4#########5#########6#########7#########8#########9

sub _get_col_row{
	my( $self, $target_col, $target_row ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD			$self->get_all_space . '::WorksheetToRow::_get_col_row', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"Reached _get_col_row",
	###LogSD			( $target_row ? "Requesting target row and column: [ $target_row, $target_col ]" : '' ),
	###LogSD			( $self->_has_old_row_inst ? ("With stored old row: " . $self->_get_old_row_number) : ''),
	###LogSD			( $self->_has_new_row_inst ? ("..and stored current row: " . $self->_get_new_row_number) : '') ] );

	# Attempt to pull the data from stored values or index the row forward
	my	$index_result = 'NoParse';
	my ( $cell_ref, $max_value_col );
	while( !defined $max_value_col ){
		if( !$self->_has_new_row_inst ){
			###LogSD	$phone->talk( level => 'debug', message => [
			###LogSD		"No stored data available - index to the next (first?) row" ] );
			$index_result = $self->_index_row;
		}elsif( $self->_get_new_row_number < $target_row ){
			###LogSD	$phone->talk( level => 'debug', message => [
			###LogSD		'The value is past the latest row pulled' ] );
			$index_result = $self->_index_row;
		}elsif( $self->_get_new_row_number == $target_row ){
			###LogSD	$phone->talk( level => 'debug', message => [
			###LogSD		'The value might be in the latest row pulled' ] );
			$max_value_col = $self->_get_new_last_value_col;
			if( $target_col > $max_value_col ){
				$cell_ref = ($self->is_empty_the_end or $self->_get_new_row_end < $target_col) ? 'EOR' : undef;
			}else{
				$cell_ref = $self->_get_new_column( $target_col );
			}
			if( $cell_ref and $cell_ref eq 'EOR' ){
				$index_result = $self->_index_row;
				$cell_ref = 'EOF' if $index_result eq 'EOF';
			}
		}elsif( $self->_has_old_row_inst ){
			if( $self->_get_old_row_number < $target_row ){
				###LogSD	$phone->talk( level => 'debug', message => [
				###LogSD		'The requested value falls between the last row and the current row' ] );
				$index_result = undef;
			}elsif( $self->_get_old_row_number == $target_row ){
				###LogSD	$phone->talk( level => 'debug', message => [
				###LogSD		'The value might be in the previous row pulled' ] );
				$max_value_col = $self->_get_old_last_value_col;
				if( $target_col > $max_value_col ){
					$cell_ref = ($self->is_empty_the_end or $self->_get_old_row_end < $target_col) ? 'EOR' : undef;
				}else{
					$cell_ref = $self->_get_old_column( $target_col );
				}
			}else{
				###LogSD	$phone->talk( level => 'debug', message => [
				###LogSD		'The value appears to exist prior to the older saved row - restarting the sheet' ] );
				$self->start_the_file_over;
				$self->_clear_old_row_inst;
				$self->_clear_new_row_inst;
				$index_result = $self->_index_row;
			}
		}else{
			###LogSD	$phone->talk( level => 'debug', message => [
			###LogSD		'The requested value falls between the beginning and the current row (and is empty)' ] );
			$index_result = undef;
		}
		if( !$index_result ){
			return 
				( 	$self->is_empty_the_end ? 'EOR' :
					$self->_max_col < $target_col ? 'EOR' : undef );
		}elsif( !$cell_ref and $index_result =~ /^EO(F|R)$/ ){
			return $index_result;
		}
	}
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		'The index result after parsing through the rows:', $index_result,
	###LogSD		"The cell ref after pulling column -$target_col-", $cell_ref, ] );
	my $updated_cell = 
		( $cell_ref ? 
			( $cell_ref =~ /^EO(F|R)$/ ? $cell_ref : $self->_complete_cell( $cell_ref ) ) : 
		($self->is_empty_the_end and $max_value_col < $target_col) ? 'EOR' : undef 		);
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		'returning ref:', $updated_cell,] );
	return $updated_cell;
}
	
sub _get_next_value_cell{
	my( $self, ) = @_; # to fast forward use _get_col_row
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD			$self->get_all_space . '::WorksheetToRow::_get_next_value_cell', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"Reached _get_next_value_cell",
	###LogSD			( $self->_has_new_row_inst ? ("With current stored new row: " . $self->_get_new_row_number) : '') ] );

	# Attempt to pull the data from stored values or index the row forward
	my	$index_result = 'NoParse';
	my	$cell_ref;
	while( !$cell_ref ){
		if( !$self->_has_new_row_inst ){
			###LogSD	$phone->talk( level => 'debug', message => [
			###LogSD		"No stored data available - index to the next (first?) row" ] );
			$index_result = $self->_index_row;
		}else{
			$cell_ref = $self->_get_new_next_value;
			###LogSD	$phone->talk( level => 'debug', message => [
			###LogSD		'Pulling the next value in the row:', $cell_ref ] );
			if( $cell_ref eq 'EOR' ){
				###LogSD	$phone->talk( level => 'debug', message => [
				###LogSD		'Reached the end of the row - starting over' ] );
				$index_result = $self->_index_row;
				$cell_ref = undef;
			}
		}
		if( !$cell_ref and $index_result =~ /^EOF/ ){
			return $index_result;
		}
	}
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		'The cell ref after parsing through the rows:', $cell_ref, ] );
	
	my $updated_cell = $self->_complete_cell( $cell_ref );
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		'returning ref:', $updated_cell,] );
	return $updated_cell;
}

sub _get_row_all{
	my( $self, $target_row ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD			$self->get_all_space . '::WorksheetToRow::_get_row_all', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"Reached _get_row_all",
	###LogSD			( $target_row ? "Requesting target row: $target_row" : '' ),
	###LogSD			( $self->_has_old_row_inst ? ("With stored old row: " . $self->_get_old_row_number) : ''),
	###LogSD			( $self->_has_new_row_inst ? ("..and stored current row: " . $self->_get_new_row_number) : '') ] );

	# Attempt to pull the data from stored values or index the row forward
	my	$index_result = 'NoParse';
	my $row_ref;
	while( $index_result and !defined $row_ref ){
		if( !$self->_has_new_row_inst ){
			###LogSD	$phone->talk( level => 'debug', message => [
			###LogSD		"No stored data available - index to the next (first?) row" ] );
			$index_result = $self->_index_row;
		}elsif( $self->_get_new_row_number < $target_row ){
			###LogSD	$phone->talk( level => 'debug', message => [
			###LogSD		'The requested row is past the latest row pulled' ] );
			$index_result = $self->_index_row;
		}elsif( $self->_get_new_row_number == $target_row ){
			###LogSD	$phone->talk( level => 'debug', message => [
			###LogSD		'The requested row is the latest row pulled' ] );
			$row_ref = $self->_get_new_row_list;
		}elsif( $self->_has_old_row_inst ){
			if( $self->_get_old_row_number < $target_row ){
				###LogSD	$phone->talk( level => 'debug', message => [
				###LogSD		'The requested row falls between the last row and the current row' ] );
				$index_result = undef;
			}elsif( $self->_get_old_row_number == $target_row ){
				###LogSD	$phone->talk( level => 'debug', message => [
				###LogSD		'The requested row is the previous row pulled' ] );
				$row_ref = $self->_get_new_old_list;
			}else{
				###LogSD	$phone->talk( level => 'debug', message => [
				###LogSD		'The requested row appears to exist prior to the older saved row - restarting the sheet' ] );
				$self->start_the_file_over;
				$self->_clear_old_row_inst;
				$self->_clear_new_row_inst;
				$index_result = $self->_index_row;
			}
		}else{
			###LogSD	$phone->talk( level => 'debug', message => [
			###LogSD		'The requested row falls between the beginning and the current row (and is empty)' ] );
			$index_result = undef;
		}
		if( !$index_result ){
			return [];
		}elsif( !$row_ref and $index_result eq 'EOF' ){
			return $index_result;
		}
	}
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		'The index result after parsing through the rows:', $index_result,
	###LogSD		"The row ref after pulling row -$target_row-", $row_ref, ] );
	my $updated_row;
	for my $cell_ref ( @$row_ref ){
		push @$updated_row, $cell_ref ? $self->_complete_cell( $cell_ref ) : $cell_ref ;
	}
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		'returning row ref:', $updated_row,] );
	return $updated_row;
}

sub _complete_cell{
	my( $self, $cell_ref ) = @_;#, $new_file, $old_file
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD			$self->get_all_space . '::WorksheetToRow::_complete_cell', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"adding worksheet data to the cell:", $cell_ref ] );
		
	#Add merge value
	my $merge_row = $self->_get_row_merge_map( $cell_ref->{cell_row} );
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"Row merge map:", $merge_row,	] );
	if( ref( $merge_row ) and $merge_row->[$cell_ref->{cell_col}] ){
		$cell_ref->{cell_merge} = $merge_row->[$cell_ref->{cell_col}];
	}
	
	# Check for hiddenness (This logic needs a deep rewrite when adding the skip_hidden attribute to the workbook)
	if( $self->is_sheet_hidden ){
		###LogSD	$phone->talk( level => 'trace', message => [
		###LogSD		'This cell is from a hidden sheet',] );
		$cell_ref->{cell_hidden} = 'sheet';
	}else{
		my $column_attributes = $self->_get_custom_column_data( $cell_ref->{cell_col} );
		#~ my $row_attributes		= $self->_get_custom_row_data( $sub_ref->{cell_row} );
		###LogSD	$phone->talk( level => 'trace', message => [
		###LogSD		"Column -$cell_ref->{cell_col}- has attributes:", $column_attributes, ] );
		if( $column_attributes and $column_attributes->{hidden} ){
			###LogSD	$phone->talk( level => 'trace', message => [
			###LogSD		'This cell is from a hidden column',] );
			$cell_ref->{cell_hidden} = 'column';
		}
	}
	###LogSD	$phone->talk( level => 'trace', message => [
	###LogSD		'Ref to this point:', $cell_ref,] );
	return $cell_ref;
}

sub _index_row{
	my( $self, ) = @_;#, $new_file, $old_file
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD			$self->get_all_space . '::WorksheetToRow::_index_row', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"Indexing the row forward by one value found position", ] );
	
	# Index the row as needed to get the next cell
	my $row_node_ref;
	while( !$row_node_ref ){
		
		# Advance to the next row node
		my ( $node_depth, $node_name, $node_type ) = $self->location_status;
		###LogSD		$phone->talk( level => 'debug', message => [
		###LogSD			"Attempting to build the next row node from node: $node_name", ] );
		if( $node_name eq 'row' ){
			###LogSD	$phone->talk( level => 'trace', message => [
			###LogSD		'The index is at a row node.' ] );
		}elsif( $self->advance_element_position( 'row' ) ){
			###LogSD	$phone->talk( level => 'trace', message => [
			###LogSD		'The index was moved to a row node.' ] );
		}else{
			$self->set_error( "No row node found where I was looking in the worksheet" );
			$self->_set_max_row( $self->_get_new_row_number );
			$self->start_the_file_over;
			$self->_clear_old_row_inst;
			$self->_clear_new_row_inst;
			return 'EOF';
		}
		
		# Turn the xml into basic perl data
		my $row_ref = $self->parse_element;
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		'Result of row read:', $row_ref ] );
				
		# Load text values for each cell where appropriate
		$row_ref->{list} =
			exists $row_ref->{list} ? $row_ref->{list} :
			exists $row_ref->{c} ? [ $row_ref->{c} ] : [];
		my ( $alt_ref, $column_to_cell_translations, $reported_column, $reported_position, $last_value_column );
		my $x = 0;
		for my $cell ( @{$row_ref->{list}} ){
			###LogSD	$phone->talk( level => 'info', message => [
			###LogSD		'Processing cell:', $cell	] );
			
			$cell->{cell_type} = 'Text';
			if( exists $cell->{t} and $cell->{t} =~ /^s/ ){# Test for all in one sheet here!(future)
				my $position = $self->get_shared_string_position( $cell->{v}->{raw_text} );
				###LogSD	$phone->talk( level => 'debug', message =>[
				###LogSD		"Shared strings returned:",  $position] );
				if( is_HashRef( $position ) ){
					@$cell{qw( cell_xml_value rich_text )} = ( $position->{raw_text}, $position->{rich_text} );
					delete $cell->{rich_text} if !$cell->{rich_text};
				}else{
					$cell->{cell_xml_value} = $position;
				}
				delete $cell->{t};
				delete $cell->{v};
				delete $cell->{cell_xml_value} if !$cell->{cell_xml_value};
			}elsif( exists $cell->{v} ){
				###LogSD	$phone->talk( level => 'debug', message =>[
				###LogSD		"Setting cell_xml_value from: $cell->{v}->{raw_text}", ] );
				$cell->{cell_xml_value} = $cell->{v}->{raw_text};
				$cell->{cell_type} = 'Numeric' if $cell->{cell_xml_value} and $cell->{cell_xml_value} ne '';
				delete $cell->{v};
			}
			if( $self->get_empty_return_type eq 'empty_string' ){
				$cell->{cell_xml_value} = '' if !exists $cell->{cell_xml_value};
			}elsif( !$cell->{cell_xml_value} or
					($cell->{cell_xml_value} and length( $cell->{cell_xml_value} ) == 0) ){
				delete $cell->{cell_xml_value};
			}
			###LogSD	$phone->talk( level => 'debug', message =>[
			###LogSD		"Updated cell:",  $cell] );
			# Clear empty cells if required
			if( $self->get_values_only and ( !$cell->{cell_xml_value} or length( $cell->{cell_xml_value} ) == 0 ) ){
					###LogSD	$phone->talk( level => 'info', message => [
					###LogSD		'Values only called - stripping this non-value cell'	] );
			}else{
				$cell->{cell_type} = 'Text' if !exists $cell->{cell_type};
				$cell->{cell_hidden} = 'row' if $row_ref->{hidden};
				@$cell{qw( cell_col cell_row )} = $self->_parse_column_row( $cell->{r} );
				$last_value_column = $cell->{cell_col};
				$cell->{cell_formula} = $cell->{f}->{raw_text} if $cell->{f};
				delete $cell->{f};
				$column_to_cell_translations->[$cell->{cell_col}] = $x++;
				$reported_column = $cell->{cell_col} if !defined $reported_column;
				$reported_position = 0;
				###LogSD	$phone->talk( level => 'info', message => [
				###LogSD		'Saving cell:', $cell	] );
				push @$alt_ref, $cell;
			}
		}
		$self->_set_row_hidden( $row_ref->{r} => ($row_ref->{hidden} ? 1 : 0) );
		
		if( $alt_ref ){
			my $new_ref;
			@$new_ref{qw( row_number row_span )} =
				( $row_ref->{r}, [split /:/, $row_ref->{spans}], );
			delete $row_ref->{r};
			delete $row_ref->{list};
			delete $row_ref->{spans};
			delete $row_ref->{hidden};
			###LogSD	$phone->talk( level => 'debug', message =>[
			###LogSD		"Updated row ref:", $new_ref, ] );
			$self->_set_old_row_inst( $self->_get_new_row_inst ) if $self->_has_new_row_inst;
			$row_node_ref =	build_instance( 
								package 		=> 'RowInstance',
								superclasses 	=> [ 'Spreadsheet::XLSX::Reader::LibXML::Row' ],
								%$new_ref,
								row_value_cells	=> $alt_ref,
								row_formats		=> $row_ref,
								row_last_value_column => $last_value_column,
								column_to_cell_translations	=> $column_to_cell_translations,
				###LogSD		log_space => $self->get_log_space
							);
			###LogSD	$phone->talk( level => 'debug', message =>[
			###LogSD		"New row instance:", $row_node_ref, ] );
			$self->_set_new_row_inst( $row_node_ref );
		}else{
			###LogSD	$phone->talk( level => 'debug', message =>[
			###LogSD		'Nothing to see here - move along', ] );
		}
	}
	return $row_node_ref ? 'GoodParse' : undef;
}

sub _load_unique_bits{
	my( $self, ) = @_;#, $new_file, $old_file
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD			$self->get_all_space . '::WorksheetToRow::_load_unique_bits', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"Setting the Worksheet unique bits", ] );
	
	# Read the sheet dimensions
	if( $self->node_name eq 'dimension' or $self->advance_element_position( 'dimension' ) ){
		my $dimension = $self->parse_element;
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"parsed dimension value:", $dimension ] );
		my	( $start, $end ) = split( /:/, $dimension->{ref} );
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"Start position: $start", 
		###LogSD		( $end ? "End position: $end" : '' ), ] );
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
		$self->_clear_old_row_inst;
		$self->_clear_new_row_inst;
	}else{
		confess "No sheet dimensions provided";# Shouldn't the error instance be loaded already?
	}
	
	#pull column stats
	my	$has_column_data = 1;
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"Loading the column configuration" ] );
	if( $self->node_name eq 'cols' or $self->advance_element_position( 'cols') ){
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"Already arrived at the column data" ] );
	}else{
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"Restart the sheet to find the column data" ] );
		$self->start_the_file_over;
		$has_column_data = $self->advance_element_position( 'cols' );
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"Column data search result: $has_column_data" ] );
	}
	if( $has_column_data ){
		my $column_data = $self->parse_element;
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"parsed column elements to:", $column_data ] );
		my $column_store = [];
		for my $definition ( @{$column_data->{list}} ){
			###LogSD	$phone->talk( level => 'debug', message => [
			###LogSD		"Processing:", $definition ] );
			my $row_ref;
			@$row_ref{qw( width customWidth bestFit hidden )} =
				( @$definition{qw( width customWidth bestFit hidden )} );
			for my $col ( $definition->{min} .. $definition->{max} ){
				$column_store->[$col] = $row_ref;
				###LogSD	$phone->talk( level => 'debug', message => [
				###LogSD		"Updated column store is:", $column_store ] );
			}
		}
		###LogSD	$phone->talk( level => 'trace', message => [
		###LogSD		"Final column store is:", $column_store ] );
		$self->_set_column_formats( $column_store );
	}
	
	#build a merge map
	my	$merge_ref = [];
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"Loading the mergeCell" ] );
	my $found_merges = 0;
	if( ($self->node_name and $self->node_name eq 'mergeCells') or $self->advance_element_position( 'mergeCells') ){
		$found_merges = 1;
	}else{
		$self->start_the_file_over;
		$found_merges = $self->advance_element_position( 'mergeCells');
	}
	if( $found_merges ){
		my $merge_range = $self->parse_element;
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"Processing all merge ranges:", $merge_range ] );
		my $final_ref;
		for my $merge_ref ( @{$merge_range->{list}} ){
			###LogSD	$phone->talk( level => 'debug', message => [
			###LogSD		"parsed merge element to:", $merge_ref ] );
			my ( $start, $end ) = split /:/, $merge_ref->{ref};
			my ( $start_col, $start_row ) = $self->_parse_column_row( $start );
			my ( $end_col, $end_row ) = $self->_parse_column_row( $end );
			###LogSD	$phone->talk( level => 'debug', message => [
			###LogSD		"Start column: $start_col", "Start row: $start_row",
			###LogSD		"End column: $end_col", "End row: $end_row" ] );
			my 	$min_col = $start_col;
			while ( $start_row <= $end_row ){
				$final_ref->[$start_row]->[$start_col] = $merge_ref->{ref};
				$start_col++;
				if( $start_col > $end_col ){
					$start_col = $min_col;
					$start_row++;
				}
			}
		}
		###LogSD	$phone->talk( level => 'trace', message => [
		###LogSD		"Final merge ref:", $final_ref ] );
		$self->_set_merge_map( $final_ref );
	}
	$self->start_the_file_over;
	return 1;
}

sub _is_column_hidden{
	my( $self, @column_requests ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD			$self->get_all_space . '::WorksheetToRow::is_column_hidden::subsub', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			'Pulling the hidden state for the columns:', @column_requests ] );
	
	my @tru_dat;
	for my $column ( @column_requests ){
		my $column_format = $self->_get_custom_column_data( $column );
		###LogSD	$phone->talk( level => 'trace', message =>[
		###LogSD		"Column formats for column -$column- are:", $column_format ] );
		push @tru_dat, (( $column_format and $column_format->{hidden} ) ? 1 : 0);
	}
	###LogSD	$phone->talk( level => 'info', message =>[
	###LogSD		"Final column hidden state is list:", @tru_dat] );
	return @tru_dat;
}

#########1 Phinish            3#########4#########5#########6#########7#########8#########9

no Moose;
__PACKAGE__->meta->make_immutable;
	
1;

#########1 Documentation      3#########4#########5#########6#########7#########8#########9
__END__

=head1 NAME

Spreadsheet::XLSX::Reader::LibXML::XMLReader::WorksheetToRow - Pull rows out of worksheet xml files

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

=head3 get_merged_areas

=over

B<Definition:> This method returns an array ref of cells that are merged.  This method does 
respond to the attribute L<Spreadsheet::XLSX::Reader::LibXML/count_from_zero>

B<Accepts:> nothing

B<Returns:> An arrayref of arrayrefs of merged areas or undef if no merged areas

	[ [ $start_row_1, $start_col_1, $end_row_1, $end_col_1], etc.. ]

=back

=head3 is_sheet_hidden

=over

B<Definition:> Method indicates if the excel program would hide the sheet or show it if the 
file were opened in the Microsoft Excel application

B<Accepts:> nothing

B<Returns:> a boolean value indicating if the sheet is hidden or not 1 = hidden

=back

=head3 is_column_hidden

=over

B<Definition:> Method indicates if the excel program would hide the identified column(s) or show 
it|them if the file were opened in the Microsoft Excel application.  If more than one column is 
passed then it returns true if any of the columns are hidden in scalar context and a list of 
1 and 0 values for each of the requested positions in array (list) context.  This method (input) 
does respond to the attribute L<Spreadsheet::XLSX::Reader::LibXML/count_from_zero>.  Unlike the 
method 'is_row_hidden' this method will always 'know' the correct answer since the information is 
stored outside of the dataa table in the xml file.

B<Accepts:> integer values or column letter values selecting the columns in question

B<Returns:> in scalar context it returns a boolean value indicating if any of the requested 
columns would be hidden by Excel.  In array/list context it returns a list of boolean values 
for each requested column indicating it's hidden state for Excel. (1 = hidden)

=back

=head3 is_row_hidden

=over

B<Definition:> Method indicates if the excel program would hide the identified row(s) or show 
it|them if the file were opened in the Microsoft Excel application.  If more than one row is 
passed then it returns true if any of the rows are hidden in scalar context and a list of 
1 and 0 values for each of the requested positions in array (list) context.  This method (input) 
does respond to the attribute L<Spreadsheet::XLSX::Reader::LibXML/count_from_zero>.  B<Warning: 
THIS METHOD WILL ONLY BE ACCURATE AFTER THE SHEET HAS READ AT LEAST ONE CELL FROM THE ROW 
NUMBER REQUESTED.  THIS ALLOWS THE SHEET TO AVOID READING ALL THE WAY THROUGH ONCE BEFORE STARTING 
THE CELL PARSING.>

B<Accepts:> integer values selecting the rows in question

B<Returns:> in scalar context it returns a boolean value indicating if any of the requested 
rows would be hidden by Excel.  In array/list context it returns a list of boolean values 
for each requested row indicating it's hidden state for Excel. (1 = hidden)

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
