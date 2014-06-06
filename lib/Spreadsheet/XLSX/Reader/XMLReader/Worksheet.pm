package Spreadsheet::XLSX::Reader::XMLReader::Worksheet;
use version; our $VERSION = version->declare("v0.1_1");

use	5.010;
use	Moose;
use	MooseX::StrictConstructor;
use	MooseX::HasDefaults::RO;
use Types::Standard qw(
		Int
		Str
		ArrayRef
		HashRef
		Object
		InstanceOf
    );
		#~ ConsumerOf
use lib	'../../../../../lib';
###LogSD	use Log::Shiras::Telephone;
###LogSD	use Log::Shiras::UnhideDebug;
extends	'Spreadsheet::XLSX::Reader::XMLReader';
with	'Spreadsheet::XLSX::Reader::CellToColumnRow'; #here to load 'set_error' first
use		Spreadsheet::XLSX::Reader::XMLDOM::Cell;
use		Spreadsheet::XLSX::Reader::Types v0.1 qw( CustomFormat );

#########1 Dispatch Tables & Package Variables    5#########6#########7#########8#########9

#~ my	$r_name_space = 'xmlns';

#########1 Public Attributes  3#########4#########5#########6#########7#########8#########9

has styles_instance =>(
	isa		=> Object,
	writer	=>'_set_styles',
	handles	=> [ qw( get_number_format get_font_definition get_default_number_format ) ],
	clearer	=> '_clear_styles',
);

has shared_strings_instance =>(
	isa		=> Object,
	writer	=>'_set_sharedStrings',
	handles	=>{
		get_shared_string_position 	=> 'get_position',
	},
	clearer	=> '_clear_shared_strings',
);

has calc_chain_instance =>(
	isa		=> Object,
	writer	=>'_set_calcChain',
	handles	=>{
		get_calc_chain_position => 'get_position',
	},
	clearer	=> '_clear_calc_chain',
);

#~ has error_inst =>(
		#~ isa			=> InstanceOf[ 'Spreadsheet::XLSX::Reader::Error' ],
		#~ handles 	=>[ qw( error set_error clear_error set_warnings if_warn ) ],
		#~ clearer		=> '_clear_error_inst',
		#~ required	=> 1,
	#~ );

#~ has epoch_year =>(
		#~ isa		=> Int,
		#~ reader	=> 'get_epoch_year',
	#~ );

has sheet_min_col =>(
		isa			=> Int,
		writer		=> '_set_min_col',
		reader		=> 'min_col',
		predicate	=> 'has_min_col',
	);

has sheet_min_row =>(
		isa			=> Int,
		writer		=> '_set_min_row',
		reader		=> 'min_row',
		predicate	=> 'has_min_row',
	);

has sheet_max_col =>(
		isa			=> Int,
		writer		=> '_set_max_col',
		reader		=> 'max_col',
		predicate	=> 'has_max_col',
	);

has sheet_max_row =>(
		isa			=> Int,
		writer		=> '_set_max_row',
		reader		=> 'max_row',
		predicate	=> 'has_max_row',
	);

has sheet_rel_id =>(
		isa		=> Str,
		reader	=> 'rel_id',
	);

has sheet_id =>(
		isa		=> Int,
		reader	=> 'sheet_id',
	);

has sheet_position =>(# XML position
		isa		=> Int,
		reader	=> 'position',
	);

has sheet_name =>(
		isa		=> Str,
		writer	=> 'set_sheet_name',
		reader	=> 'name',
	);

has custom_formats =>(
		isa		=> HashRef[ CustomFormat ],
		handles	=>{
			has_custom_format => 'exists',
			get_custom_format => 'get',
		},
	);

#########1 Public Methods     3#########4#########5#########6#########7#########8#########9

sub row_range{
	my( $self ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD					name_space 	=> ($self->get_log_space .  '::row_range' ), );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"Returning row range( " . $self->min_row . ", " . $self->max_row . " )" ] );
	return( $self->min_row, $self->max_row );
}

sub col_range{
	my( $self ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD					name_space 	=> ($self->get_log_space .  '::col_range' ), );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"Returning col range( " . $self->min_col . ", " . $self->max_col . " )" ] );
	return( $self->min_col, $self->max_col );
}

sub get_cell{
    my ( $self, $row, $column ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD					name_space 	=> $self->get_log_space .  '::get_cell', );
	###LogSD		$phone->talk( level => 'info', message =>[
	###LogSD			"Arrived at get_cell with: ",
	###LogSD			"Row: " . ($row ? $row : ''),
	###LogSD			"Column: " . ($column ? $column : '' ) ] );
	
	# Handle implied next column
	my	$add_row = 0;
	if( defined $column and !defined $row ){
		$self->error( "Requested column -$column- provided but no requested row identified" );
		$self->_clear_column_position;
		$self->_clear_row_position;
		return undef;
	}elsif( !defined $column ){
		###LogSD	$phone->talk( level => 'info', message =>[
		###LogSD		"Handling an implied 'next column'" ] );
		$column = ( $self->_has_column_position ) ? ( $self->_get_column_position + 1 ): 1;
	}
	if( $column > $self->max_col ){
		###LogSD	$phone->talk( level => 'info', message =>[
		###LogSD		"Wrap the column to the next row ..." ] );
		$column = 1;
		$add_row = 1;
	}
	
	# Handle implied next row
	if( !defined $row ){
		###LogSD	$phone->talk( level => 'info', message =>[
		###LogSD		"No row passed - using the current row + " . (( $add_row ) ? 1 : 0) ] );
		$row = ( $self->____has_row_position ) ? $self->_get_row_position : 1;
		$row += $add_row;
	}elsif( $add_row ){
		$self->error( "Reached the end of the row" );
		$self->_clear_column_position;
		$self->_clear_row_position;
		return undef;
	}
	if( $row > $self->max_row ){
		$self->error( "That is past the bottom of the sheet" );
		$self->_clear_column_position;
		$self->_clear_row_position;
		return undef;
	}
	###LogSD	$phone->talk( level => 'info', message =>[
	###LogSD		"Next cell resolved to;", "Row: $row", "Column: $column" ] );
	$self->_set_row_position( $row );
	$self->_set_column_position( $column );
	
	# Find the cell data in the XML files
	my	$row_map_ref = $self->_get_row_map( $row );
	
	# Handle empty cells
	if( !defined $row_map_ref->[0] or !defined $row_map_ref->[$column] ){
		###LogSD	$phone->talk( level => 'info', message =>[
		###LogSD		"Found and empty cell at;", "Row: $row", "Column: $column" ] );
		return undef;
	}
	###LogSD	no warnings 'uninitialized';
	###LogSD	$phone->talk( level => 'info', message =>[
	###LogSD		"has_position: " . $self->has_position,
	###LogSD		"for row:", $row_map_ref,
	###LogSD		"where_am_i: " . $self->where_am_i, ] );
	###LogSD	use warnings 'uninitialized';
	# Get the right row element
	if(	!$self->has_position or $row_map_ref->[0] != $self->where_am_i ){
		###LogSD	$phone->talk( level => 'info', message =>[
		###LogSD		"Resetting the row element position to: " . $row_map_ref->[0] ] );
		$self->_set_row_chunk( $self->get_position( $row_map_ref->[0] ) );
	}
	
	my	$cell_node = ($self->get_all_cells)[$row_map_ref->[$column]];#$self->_find_nodes( '../row/c', $column_ref );#$xpath_expression
	$cell_node = $cell_node->cloneNode( 1 );
	###LogSD	$phone->talk( level => 'debug', message =>[
	###LogSD		"Working with cell: " . $cell_node->toString ] );
	my	$stored_type 	= $cell_node->getAttribute( 't' );
		$stored_type  //= 'number';
	my	$format_position= $cell_node->getAttribute( 's' );
	my	$value_node		= ($cell_node->getChildrenByTagName( 'v' ))[0];
	my	$text_content	= $value_node->textContent;
	###LogSD	$phone->talk( level => 'trace', message => [
	###LogSD		 "Cell data type: $stored_type",
	###LogSD		 "Value node: " . $value_node->toString,
	###LogSD		 "Text Content: $text_content"	] );
	###LogSD	$phone->talk( level => 'debug', #ask => 1,
	###LogSD		 message => [ "Cell is:", $cell_node->toString, ] );
	if( $stored_type eq 's' ){
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"Loading shared strings position: $text_content",
		###LogSD		"with data: " . $self->get_shared_string_position( $text_content )->textContent	] );
			$cell_node->replaceChild(
				$self->get_shared_string_position( $text_content ), $value_node,
			);
	}elsif( $stored_type eq 'number' ){
		###LogSD	$phone->talk( level => 'trace',
		###LogSD		 message => [ "No action needed for numbers ..." ] );
	}else{
		$self->_set_error( "The cell in row -$row- column -$column- has an unrecognized " .
							"value for attribute 't': $stored_type" );
	}
	my	$format_definition;
	if( defined $format_position ){
		$format_definition = $self->get_number_format( $format_position );
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"The defined format position is: $format_position",
		###LogSD		"With data:", $format_definition ] );
	}else{
		$format_definition = $self->get_default_number_format;
		###LogSD	$phone->talk( level => 'debug',
		###LogSD		 message => [ "Using the default number format", $format_position ] );
	}
	###LogSD	$phone->talk( level => 'trace', message => [ 
	###LogSD		 "Updated cell is:", $cell_node->toString, ] );
	my	$cell_instance = Spreadsheet::XLSX::Reader::XMLDOM::Cell->new(
							cell_element 	=> $cell_node,
							log_space		=> $self->get_log_space . '::Cell',
							value_encoding	=> $self->encoding,
							value_type		=> $stored_type,
							error_inst		=> $self->_get_error_inst,
							cell_column		=> $column,
							cell_row		=> $row,
							number_format	=> $format_definition, 
						);
	###LogSD	$phone->talk( level => 'trace', ask => "continue?", message => [
	###LogSD		"Built Cell:", $cell_instance ] );\
	return $cell_instance;
}

#########1 Private Attributes 3#########4#########5#########6#########7#########8#########9
	
has +_core_element =>(
		default => 'row',
	);

has _column_position_index =>(
		isa			=> Int,
		reader		=> '_get_column_position_index',
		writer		=> '_set_column_position_index',
		clearer		=> '_clear_column_position_index',
		predicate	=> '_has_column_position_index',
	);

has _row_position =>(
		isa			=> Int,
		reader		=> '_get_row_position',
		writer		=> '_set_row_position',
		clearer		=> '_clear_row_position',
		predicate	=> '_has_row_position',
	);

has _column_position =>(
		isa			=> Int,
		reader		=> '_get_column_position',
		writer		=> '_set_column_position',
		clearer		=> '_clear_column_position',
		predicate	=> '_has_column_position',
	);
	
has _current_row_chunk =>(
		isa			=> 'XML::LibXML::Element',#XPathContext',
		clearer		=> '_clear_row_chunk',
		predicate	=> '_has_row_chunk',
		writer		=> '_set_row_chunk',
		handles	=>{
			get_row_number	=>[ getAttribute => 'r' ],
			get_row_span	=>[ getAttribute => 'spans' ],
			get_all_cells	=>[ getChildrenByTagName => 'c' ],
			#~ _get_first_column	=> 'firstChild',
			#~ _get_next_column	=> 'nextSibling',
			#~ _get_columns	=> 'getChildrenByTagName'
		},
	);

has	_sheet_map =>(
		isa		=> ArrayRef,
		traits	=> ['Array'],
		writer	=> '_set_sheet_map',
		handles	=>{
			_get_row_map => 'get',
		},
	);

has _sheet_unique_count =>(
	isa			=> Int,
	writer		=> '_set_unique_count',
	clearer		=> '_clear_unique_count',
	predicate	=> '_has_unique_count',
	reader		=> '_get_unique_count',
);
	
#~ has _xmlns =>(
		#~ isa		=> Str,
		#~ writer	=> '_set_xmlns_file',
		#~ default	=> 'about:blank',
	#~ );
	
#~ has _xpath_context =>(
		#~ isa		=> 'XML::LibXML::XPathContext',
		#~ default	=> sub{ XML::LibXML::XPathContext->new },
		#~ handles	=>{
			#~ _set_r_name_space	=>[ registerNs => $r_name_space ],#
			#~ _set_top_name_space =>[ registerNs => 'xmlns' ]
			#~ _find_nodes => 'findnodes',
		#~ },
	#~ );


#########1 Private Methods    3#########4#########5#########6#########7#########8#########9

sub _load_unique_bits{
	my( $self, $reader, $encoding ) = @_;#, $new_file, $old_file
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD					name_space 	=> ($self->get_log_space . '::_load_unique_bits' ), );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"Setting the Worksheet unique bits with reader:", $reader ] );
	
	#~ # Load the xmlns value
	#~ if( ( $reader->name eq 'worksheet' ) or $reader->nextElement( 'worksheet' ) ){
		#~ my	$file = $reader->getAttribute( $r_name_space );
		#~ ###LogSD	$phone->talk( level => 'debug', message =>[ "-$r_name_space- value: $file" ] );
		#~ $self->_set_r_name_space( $file );
	#~ }else{
		#~ $self->_set_error( "No xmlns value provided - using default" );
	#~ }
	#~	###LogSD	$phone->talk( level => 'trace', message => [
	#~	###LogSD		"Self to this point: ", $self ] );
	
	# Read the dimension
	if( $reader->nextElement( 'dimension' ) ){
		my	$range = $reader->getAttribute( 'ref' );
		my	( $start, $end ) = split( /:/, $range );
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"Start position: $start", "End position: $end" ] );
		my ( $start_column, $start_row ) 	= $self->parse_column_row( $start );
		my ( $end_column, $end_row	)		= $self->parse_column_row( $end );
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"Start column: $start_column", "Start row: $start_row",
		###LogSD		"End column: $end_column", "End row: $end_row" ] );
		$self->_set_min_col( $start_column );
		$self->_set_min_row( $start_row );
		$self->_set_max_col( $end_column );
		$self->_set_max_row( $end_row );
	}else{
		$self->_set_error( "No sheet dimensions provided" );
	}
	###LogSD	$phone->talk( level => 'trace', message => [
	###LogSD		"Self to this point: ", $self ] );
		
	#Map the sheet
	my ( $sheet_map, $min_col, $min_row );
	my ( $max_col, $max_row ) = ( 0, 0 );
	my	$old_row = 1;
	my	$current_position = 0;
	my	$old_column = 1;
	my	$current_cell_position = 0;
	while( $reader->nextElement('c') ){
		my	$rc = $reader->getAttribute( 'r' );
		###LogSD	$phone->talk( level => 'trace', message => [
		###LogSD		"first cell:", $reader->copyCurrentNode( 1 )->toString,
		###LogSD		"Column Row attribute value: $rc",	] );
			#~ $stored_type //= 'number';
		#~ my	$cell_instance = Spreadsheet::XLSX::Reader::XMLDOM::Cell->new(
				#~ _cell_element 		=> $reader->copyCurrentNode( 1 ),
				#~ log_space			=> $self->get_log_space . '::Cell',
				#~ _string_encoding	=> $encoding,
				#~ _stored_type		=> $stored_type,
				#~ _workbook_link		=> $self->_get_workbook_link,
			#~ );
		#~ ###LogSD	$phone->talk( level => 'debug', message => [
		#~ ###LogSD		"Built Cell:\n", $cell_instance ] );
		my ( $col, $row ) = $self->parse_column_row( $rc );
		if( !defined $min_col or $min_col > $col ){
			$min_col = $col;
		}
		if( !defined $max_col or $max_col < $col ){
			$max_col = $col;
		}
		if( !defined $min_row or $min_row > $row ){
			$min_row = $row;
		}
		if( !defined $max_row or $max_row < $row ){
			$max_row = $row;
		}
		if( $row > ( $old_row + 1 ) ){
			# Fill in the empty rows
			for my $x ( ( $old_row + 1 ) .. ( $row - 1 ) ){
				$sheet_map->[$x]->[0] = undef;
				$sheet_map->[$x]->[($self->max_col // $max_col)] = undef;
			}
		}
		if( $row > $old_row ){
			# New row position
			$current_position++;
			$old_column = 1;
			$current_cell_position = 0;
			$old_row = $row;
		}
		$sheet_map->[$row]->[0] = $current_position;
		$sheet_map->[$row]->[$col] = $current_cell_position++;
		###LogSD	$phone->talk( level => 'trace', message => [
		###LogSD		"Current sheet map:\n", $sheet_map ] );
	}
	$self->_set_sheet_map( $sheet_map );
	$self->_set_unique_count( $current_position );
	if( !$self->has_min_col ){
		$self->set_min_col( $min_col );
	}
	if( !$self->has_max_col ){
		$self->set_max_col( $max_col );
	}
	if( !$self->has_min_row ){
		$self->set_min_row( $min_row );
	}
	if( !$self->has_max_row ){
		$self->set_max_row( $max_row );
	}
	return 1;# Reload the sheet
}

augment 'get_position' => sub{
	my ( $self, )	= shift;
	my	$position 	= $self->_get_requested_position;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD					$self->get_log_space . '::get_position::augmented', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"Reached augment::_get_position for: $position" ] );
	
	#checking if the reqested position is too far
	if( $position > $self->_get_unique_count ){
		###LogSD	$phone->talk( level => 'warn', message => [
		###LogSD		"Asking for position -$position- (from 0) but the worksheet " .
		###LogSD		"max row position is: " . ($self->_get_unique_count) ] );
		return 1;#  fail
	}else{
		###LogSD	$phone->talk( level => 'warn', message =>[ "No end in sight" ] );
		return undef;#No failure
	}
	#~ my( $self, ) = @_;
	#~	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	#~	###LogSD			name_space 	=> $self->get_log_space . '::get_next_position::augmented', );
	#~ 	###LogSD		$phone->talk( level => 'debug', message => [
	#~	###LogSD			"Reached get_next_position to see if we have reached the end" ] );
	#~ if( $self->has_position and $self->_has_unique_count and
		#~ ( 1 + $self->where_am_i )  > ( $self->unique_count - 1 ) ){
		#~	###LogSD	$phone->talk( level => 'debug', message => [
		#~	###LogSD		"Reached the end of the file" ] );
		#~ return 1;# Reached the end
	#~ }else{
		#~ return undef;# No end in sight
	#~ }
};

sub DEMOLISH{
	my ( $self ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD					name_space 	=> $self->get_log_space .  '::DEMOLISH', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD				'Last called by:' . (caller(2))[3] . ' - ' . (caller(2))[2],
	###LogSD				"Clearing the open row chunk" ] );
	$self->_clear_row_chunk;
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"Clearing the XML::LibXML::Reader file link for: " . $self->name,
	###LogSD		"File name: " . $self->get_file_name,									] );
	$self->_clear_xml_parser;
	###LogSD	$phone->talk( level => 'debug', message =>[
	###LogSD		"Clearing all shared files" 			] );
	#~ $self->_clear_error_inst,
	$self->_clear_calc_chain;
	$self->_clear_shared_strings;
	$self->_clear_styles;
}

#########1 Phinish            3#########4#########5#########6#########7#########8#########9

no Moose;
__PACKAGE__->meta->make_immutable(
	inline_constructor => 0,
);
	
1;

#########1 Documentation      3#########4#########5#########6#########7#########8#########9
__END__

=head1 NAME

Spreadsheet::XLSX::Reader::Worksheet - Spreadsheet::XLSX::Reader::Worksheet reader
    
=head1 DESCRIPTION


#########1 Documentation End  3#########4#########5#########6#########7#########8#########9