package Spreadsheet::XLSX::Reader::LibXML::XMLReader::WorksheetToRow;
our $AUTHORITY = 'cpan:JANDREW';
use version; our $VERSION = qv('v0.38.18');
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
			_get_old_last_value_col	=> 'get_last_value_column',
			_get_old_row_list		=> 'get_row_all',
			_get_old_row_end		=> 'get_row_end'
		},
	);

has _new_row_inst =>(
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
		$row_ref->{list} =
			exists $row_ref->{list} ? $row_ref->{list} :
			exists $row_ref->{c} ? [ $row_ref->{c} ] : [];
		delete $row_ref->{c} if exists $row_ref->{c};# Delete the single column c placeholder as needed
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		'Result of row read:', $row_ref ] );
				
		# Load text values for each cell where appropriate
		my ( $alt_ref, $column_to_cell_translations, $reported_column, $reported_position, $last_value_column );
		my $x = 0;
		for my $cell ( @{$row_ref->{list}} ){
			###LogSD	$phone->talk( level => 'info', message => [
			###LogSD		'Processing cell:', $cell	] );
			
			$cell->{cell_type} = 'Text';
			if( exists $cell->{t} ){
				if( $cell->{t} eq 's' ){
					###LogSD	$phone->talk( level => 'debug', message =>[
					###LogSD		"Identified potentially required shared string for cell:",  $cell] );
					my $position = ( $self->_has_shared_strings_file ) ?
							$self->get_shared_string_position( $cell->{v}->{raw_text} ) : $cell->{v}->{raw_text};
					###LogSD	$phone->talk( level => 'debug', message =>[
					###LogSD		"Shared strings resolved to:",  $position] );
					if( is_HashRef( $position ) ){
						@$cell{qw( cell_xml_value rich_text )} = ( $position->{raw_text}, $position->{rich_text} );
						delete $cell->{rich_text} if !$cell->{rich_text};
					}else{
						$cell->{cell_xml_value} = $position;
					}
				}elsif( $cell->{t} eq 'str' ){
					###LogSD	$phone->talk( level => 'debug', message =>[
					###LogSD		"Identified a stored string in the worksheet file: " . ($cell->{v}//'')] );
					$cell->{cell_xml_value} = $cell->{v}->{raw_text};
				}else{
					confess "Unknow 't' attribute set for the cell: $cell->{t}";
				}
				delete $cell->{t};
				delete $cell->{v};
				delete $cell->{cell_xml_value} if !defined $cell->{cell_xml_value};
			}elsif( exists $cell->{v} ){
				###LogSD	$phone->talk( level => 'debug', message =>[
				###LogSD		"Setting cell_xml_value from: $cell->{v}->{raw_text}", ] );
				$cell->{cell_xml_value} = $cell->{v}->{raw_text};
				$cell->{cell_type} = 'Numeric' if $cell->{cell_xml_value} and $cell->{cell_xml_value} ne '';
				delete $cell->{v};
			}
			if( $self->get_empty_return_type eq 'empty_string' ){
				$cell->{cell_xml_value} = '' if !exists $cell->{cell_xml_value};
			}elsif( !defined $cell->{cell_xml_value} or
					($cell->{cell_xml_value} and length( $cell->{cell_xml_value} ) == 0) ){
				delete $cell->{cell_xml_value};
			}
			###LogSD	$phone->talk( level => 'debug', message =>[
			###LogSD		"Updated cell:",  $cell] );
			# Clear empty cells if required
			if( $self->get_values_only and ( !defined $cell->{cell_xml_value} or length( $cell->{cell_xml_value} ) == 0 ) ){
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
			###LogSD	$phone->talk( level => 'trace', message =>[
			###LogSD		"Row ref:", $row_ref, ] ) if !$row_ref->{spans};
			$new_ref->{row_number} = $row_ref->{r};
			$new_ref->{row_span} = $row_ref->{spans} ?
				[split /:/, $row_ref->{spans}] : [ $self->_min_col, $self->_max_col ];
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
	my ( $node_depth, $node_name, $node_type ) = $self->location_status;
	if( $node_name eq 'dimension' or $self->advance_element_position( 'dimension' ) ){
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
	( $node_depth, $node_name, $node_type ) = $self->location_status;
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"Loading the column configuration" ] );
	if( $node_name eq 'cols' or $self->advance_element_position( 'cols') ){
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
			next if !is_HashRef( $definition );
			###LogSD	$phone->talk( level => 'debug', message => [
			###LogSD		"Processing:", $definition ] );
			my $row_ref;
			map{ $row_ref->{$_} = $definition->{$_} if defined $definition->{$_} } qw( width customWidth bestFit hidden );
			###LogSD	$phone->talk( level => 'debug', message => [
			###LogSD		"Updated row ref:", $row_ref ] );
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
	( $node_depth, $node_name, $node_type ) = $self->location_status;
	my $found_merges = 0;
	if( ($node_name and $node_name eq 'mergeCells') or $self->advance_element_position( 'mergeCells') ){
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

See t\Spreadsheet\XLSX\Reader\LibXML02-worksheet_to_row.t
    
=head1 DESCRIPTION

This documentation is written to explain ways to use this module when writing your own excel 
parser.  To use the general package for excel parsing out of the box please review the 
documentation for L<Workbooks|Spreadsheet::XLSX::Reader::LibXML>,
L<Worksheets|Spreadsheet::XLSX::Reader::LibXML::Worksheet>, and 
L<Cells|Spreadsheet::XLSX::Reader::LibXML::Cell>

This module provides the basic connection to individual worksheet files (not chartsheets) for 
parsing xlsx workbooks and coalating shared strings data to cell data.  It does not provide 
a way to connect to L<chartsheets|Spreadsheet::XLSX::Reader::LibXML::Chartsheet>.  It does 
not provide the final view of a given cell.  The final view of the cell is collated with 
the role (Interface) L<Spreadsheet::XLSX::Reader::LibXML::Worksheet>.  This reader extends 
the base reader class L<Spreadsheet::XLSX::Reader::LibXML::XMLReader>.  The functionality 
provided by those modules is not explained here.

For now this module reads each full row (with values) into a L<Spreadsheet::XLSX::Reader::LibXML::Row> 
instance.  It stores only the currently read row and the previously read row.  Exceptions to 
this are the start of read and end of read.  For start of read only the current row is available 
with the assumption that all prior implied rows are empty.  When a position past the end of the sheet 
is called both current and prior rows are cleared and an 'EOF' or undef value is returned.  See
L<Spreadsheet::XLSX::Reader::LibXML/file_boundary_flags> for more details.  This allows for storage 
of row general formats by row and where a requested cell falls in a row without values that the empty 
state can be determined without rescanning the file.

I<All positions (row and column places and integers) at this level are stored and returned in count 
from one mode!>

Modification of this module probably means extending a different reader or using other roles 
for implementation of the class.  Search for

	extends	'Spreadsheet::XLSX::Reader::LibXML::XMLReader';
	
To replace the base reader. Search for the method 'worksheet' in L<Spreadsheet::XLSX::Reader::LibXML> 
and the variable '$parser_modules' to replace this whole thing.

=head2 Attributes

Data passed to new when creating an instance.  For access to the values in these 
attributes see the listed 'attribute methods'. For general information on attributes see 
L<Moose::Manual::Attributes>.  For ways to manage the instance when opened see the 
L<Public Methods|/Public Methods>.
	
=head3 is_hidden

=over

B<Definition:> This is set when the sheet is read from the sheet metadata level indicating 
if the sheet is hidden

B<Default:> none

B<Range:> (1|0)

B<attribute methods> Methods provided to adjust this attribute
		
=over

B<is_sheet_hidden>

=over

B<Definition:> return the attribute value

=back

=back

=back

=head3 workbook_instance

=over

B<Definition:> This attribute holds a reference back to the workbook instance so that 
the worksheet has access to the global settings managed there.  As a consequence many 
of the workbook methods are be exposed here.  This includes some setter methods for 
workbook attributes. I<Beware that setting or adjusting the workbook level attributes 
with methods here will be universal and affect other worksheets.  So don't forget to 
return the old value if you want the old behavour after you are done.>  If that 
doesn't make sense then don't use these methods.  (Nothing to see here! Move along.)

B<Default:> a Spreadsheet::XLSX::Reader::LibXML instance

B<attribute methods> Methods of the workbook exposed here by the L<delegation
|Moose::Manual::Attributes/Delegation> of the instance to this class through this 
attribute

=over

B<counting_from_zero>

=over

B<Definition:> returns the L<Spreadsheet::XLSX::Reader::LibXML/count_from_zero> 
instance state

=back

B<boundary_flag_setting>

=over

B<Definition:> returns the L<Spreadsheet::XLSX::Reader::LibXML/file_boundary_flags> 
instance state

=back

B<change_boundary_flag( $Bool )>

=over

B<Definition:> sets the L<Spreadsheet::XLSX::Reader::LibXML/file_boundary_flags> 
instance state (B<For the whole workbook!>)

=back

B<get_shared_string_position( $int )>

=over

B<Definition:> returns the shared string data stored in the sharedStrings 
file at position $int.  For more information review 
L<Spreadsheet::XLSX::Reader::LibXML::SharedStrings>.  I<This is a delegation 
of a delegation!>

=back

B<get_format_position( $int, [$header] )>

=over

B<Definition:> returns the format data stored in the styles 
file at position $int.  If the optional $header is passed only the data for that 
header is returned.  Otherwise all styles for that position are returned.  
For more information review 
L<Spreadsheet::XLSX::Reader::LibXML::Styles>.  I<This is a delegation 
of a delegation!>

=back

B<set_empty_is_end( $Bool )>

=over

B<Definition:> sets the L<Spreadsheet::XLSX::Reader::LibXML/empty_is_end> 
instance state (B<For the whole workbook!>)

=back

B<is_empty_the_end>

=over

B<Definition:> returns the L<Spreadsheet::XLSX::Reader::LibXML/empty_is_end> 
instance state.

=back

B<get_group_return_type>

=over

B<Definition:> returns the L<Spreadsheet::XLSX::Reader::LibXML/group_return_type> 
instance state.

=back

B<set_group_return_type( (instance|unformatted|value) )>

=over

B<Definition:> sets the L<Spreadsheet::XLSX::Reader::LibXML/group_return_type> 
instance state (B<For the whole workbook!>)

=back

B<get_epoch_year>

=over

B<Definition:> uses the L<Spreadsheet::XLSX::Reader::LibXML/get_epoch_year> method.

=back

B<get_date_behavior>

=over

B<Definition:> This is a L<delegated|Moose::Manual::Delegation> method from the 
L<styles|Spreadsheet::XLSX::Reader::LibXML::Styles> class (stored as a private 
instance in the workbook).  It is held (and documented) in the 
L<Spreadsheet::XLSX::Reader::LibXML::ParseExcelFormatStrings> role.  It will 
indicate how far unformatted L<transformation
|Spreadsheet::XLSX::Reader::LibXML::ParseExcelFormatStrings/datetime_dates> 
is carried for date coercions when returning formatted values. 

=back

B<set_date_behavior>

=over

B<Definition:> This is a L<delegated|Moose::Manual::Delegation> method from 
the L<styles|Spreadsheet::XLSX::Reader::LibXML::Styles> class (stored as a private 
instance in the workbook).  It is held (and documented) in the 
L<Spreadsheet::XLSX::Reader::LibXML::ParseExcelFormatStrings> role.  It will set how 
far unformatted L<transformation
|Spreadsheet::XLSX::Reader::LibXML::ParseExcelFormatStrings/datetime_dates> 
is carried for date coercions when returning formatted values. 

=back

B<get_values_only>

=over

B<Definition:> gets the L<Spreadsheet::XLSX::Reader::LibXML/values_only> 
instance state.

=back

B<set_values_only>

=over

B<Definition:> sets the L<Spreadsheet::XLSX::Reader::LibXML/values_only> 
instance state (B<For the whole workbook!>)

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
into this format for reading by this module when it first opens the worksheet.

B<Range:> an array ref

B<attribute methods> Methods provided to adjust this attribute
		
=over

B<_set_merge_map>

=over

B<Definition:> sets the attribute value

=back

=back

B<_get_merge_map>

=over

B<Definition:> returns the attribute array of arrays

=back

=back

B<delegated methods> This attribute uses the native trait 'Array'
		
=over

B<_get_row_merge_map( $int )> delgated from 'Array' 'get'

=over

B<Definition:> returns the sub array ref representing any merges for that 
row.  If no merges are available for that row it returns undef.

=back

=back
	
=head3 _column_formats

=over

B<Definition:> In order to (eventually) show all column formats that also affect individual 
cells the column based formats are read from the metada when the worksheet is opened.  They
are stored here for use although for now they are mostly used to determine the hidden state of 
the column.  The formats are stored in the array by count from 1 column position.

B<Range:> an array ref

B<attribute methods> Methods provided to adjust this attribute
		
=over

B<_set_set_column_formats>

=over

B<Definition:> sets the attribute value

=back

=back

B<_get_get_column_formats>

=over

B<Definition:> returns the attribute array

=back

=back

B<delegated methods> This attribute uses the native trait 'Array'
		
=over

B<_get_custom_column_data( $int )> delgated from 'Array' 'get'

=over

B<Definition:> returns the sub hash ref representing any formatting 
for that column.  If no custom formatting is available it returns undef.

=back

=back
	
=head3 _old_row_inst

=over

B<Definition:> This is the prior read row instance or undef for the beginning or 
end of the sheet read.

B<Range:> isa => InstanceOf[ L<Spreadsheet::XLSX::Reader::LibXML::Row> ]

B<attribute methods> Methods provided to adjust this attribute
		
=over

B<_set_old_row_inst>

=over

B<Definition:> sets the attribute value

=back

B<_get_old_row_inst>

=over

B<Definition:> returns the attribute

=back

B<_clear_old_row_inst>

=over

B<Definition:> clears the attribute

=back

B<_has_old_row_inst>

=over

B<Definition:> predicate for the attribute

=back

B<delegated methods> from L<Spreadsheet::XLSX::Reader::LibXML::Row>
		
=over

B<_get_old_row_number> = L<Spreadsheet::XLSX::Reader::LibXML::Row/get_row_number>

B<_is_old_row_hidden> = L<Spreadsheet::XLSX::Reader::LibXML::Row/is_row_hidden>

B<_get_old_row_formats> = L<Spreadsheet::XLSX::Reader::LibXML::Row/get_row_format>

=over

pass the desired format key

=back

B<_get_old_column> = L<Spreadsheet::XLSX::Reader::LibXML::Row/get_the_column( $column )>

=over

pass a column number (no next default) returns (cell|undef|EOR)

=back

B<_get_old_last_value_col> = L<Spreadsheet::XLSX::Reader::LibXML::Row/get_last_value_column>

B<_get_old_row_list> = L<Spreadsheet::XLSX::Reader::LibXML::Row/get_row_all>

B<_get_old_row_end> = L<Spreadsheet::XLSX::Reader::LibXML::Row/get_row_endl>

=back

=back

=back
	
=head3 _new_row_inst

=over

B<Definition:> This is the current read row instance or undef for the end of the sheet 
read.

B<Range:> isa => InstanceOf[ L<Spreadsheet::XLSX::Reader::LibXML::Row> ]

B<attribute methods> Methods provided to adjust this attribute
		
=over

B<_set_new_row_inst>

=over

B<Definition:> sets the attribute value

=back

B<_get_new_row_inst>

=over

B<Definition:> returns the attribute

=back

B<_clear_new_row_inst>

=over

B<Definition:> clears the attribute

=back

B<_has_new_row_inst>

=over

B<Definition:> predicate for the attribute

=back

B<delegated methods> from L<Spreadsheet::XLSX::Reader::LibXML::Row>
		
=over

B<_get_new_row_number> = L<Spreadsheet::XLSX::Reader::LibXML::Row/get_row_number>

B<_is_new_row_hidden> = L<Spreadsheet::XLSX::Reader::LibXML::Row/is_row_hidden>

B<_get_new_row_formats> = L<Spreadsheet::XLSX::Reader::LibXML::Row/get_row_format>

=over

pass the desired format key

=back

B<_get_new_column> = L<Spreadsheet::XLSX::Reader::LibXML::Row/get_the_column( $column )>

=over

pass a column number (no next default) returns (cell|undef|EOR)

=back

B<_get_new_next_value> = L<Spreadsheet::XLSX::Reader::LibXML::Row/get_the_next_value_position>

=over

pass nothing returns next (cell|EOR)

=back

B<_get_new_last_value_col> = L<Spreadsheet::XLSX::Reader::LibXML::Row/get_last_value_column>

B<_get_new_row_list> = L<Spreadsheet::XLSX::Reader::LibXML::Row/get_row_all>

B<_get_new_row_end> = L<Spreadsheet::XLSX::Reader::LibXML::Row/get_row_endl>

=back

=back

=back
	
=head3 _row_hidden_states

=over

B<Definition:> As the worksheet is parsed it will store the hidden state for 
the row in this attribute when each row is read.  This is the only worksheet 
level caching done.  B<It will not test whether the requested row hidden state 
has been read when accessing this data.>  If a method call a row past the 
current max parsed row it will return 0 (unhidden).

B<Range:> an array ref of Boolean values

B<delegated methods> This attribute uses the native trait 'Array'
		
=over

B<_set_row_hidden( $int )> delgated from 'Array' 'set'

=over

B<Definition:> sets the hidden state for that $int (row) counting from 1.

=back

B<_get_row_hidden( $int )> delgated from 'Array' 'get'

=over

B<Definition:> returns the known hidden state of the row.

=back

=back

=back

=head2 Methods

These are the methods provided by this class for use within the package but are not intended 
to be used by the end user.  Other private methods not listed here are used in the module but 
not used by the package.  If the private method is listed here then replacement of this module 
either requires replacing them or rewriting all the associated connecting roles and classes.

=head3 _load_unique_bits

=over

B<Definition:> This is called by L<Spreadsheet::XLSX::Reader::LibXML::XMLReader> when the file is 
loaded for the first time so that file specific metadata can be collected.

B<Accepts:> nothing

B<Returns:> nothing

=back

=head3 _get_next_value_cell

=over

B<Definition:> This returns the worksheet file hash ref representation of the xml stored for the 
'next' value cell.  A cell is determined to have value based on the attribute 
L<Spreadsheet::XLSX::Reader::LibXML/values_only>.  Next is affected by the attribute 
L<Spreadsheet::XLSX::Reader::LibXML/empty_is_end>.  This method never returns an 'EOR' flag.  
It just wraps automatically.  This does return values from the shared strings file integrated but 
not values from the Styles file integrated.

B<Accepts:> nothing

B<Returns:> a hashref of key value pairs

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

=head3 _is_column_hidden( @query_list )

=over

B<Definition:> This is returns a list of hidden states for each column integer in the @query_list 
it will generally return n array ref of each of the values in the row placed in their 'count 
from one' position.  If the row is empty but it is not the end of the sheet then this will return an 
empty array ref.

B<Accepts:> ( @query_list ) - integers in count from 1 representing requested columns

B<Returns (when wantarray):> a list of hidden states as follows; 1 => hidden, 0 => known to be unhidden, 
undef => unknown state (usually this represents columns before min_col or after max_col or at least past 
the last stored value in the column)

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

L<Clone> - clone

L<Carp> - confess

L<Type::Tiny> - 1.000

L<MooseX::ShortCut::BuildInstance> - build_instance should_re_use_classes

L<Spreadsheet::XLSX::Reader::LibXML>

L<Spreadsheet::XLSX::Reader::LibXML::XMLReader>

L<Spreadsheet::XLSX::Reader::LibXML::Row>

L<Spreadsheet::XLSX::Reader::LibXML::CellToColumnRow>

L<Spreadsheet::XLSX::Reader::LibXML::XMLToPerlData>

=back

=head1 SEE ALSO

=over

L<Log::Shiras|https://github.com/jandrew/Log-Shiras>

=over

All lines in this package that use Log::Shiras are commented out

=back

=back

=cut

#########1 Documentation End  3#########4#########5#########6#########7#########8#########9
