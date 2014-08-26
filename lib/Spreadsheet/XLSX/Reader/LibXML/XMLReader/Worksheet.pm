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
		HasMethods
    );
		#~ ConsumerOf
use lib	'../../../../../lib';
extends	'Spreadsheet::XLSX::Reader::XMLReader';
###LogSD	use Log::Shiras::Telephone;
###LogSD	use Log::Shiras::UnhideDebug;
with	'Spreadsheet::XLSX::Reader::CellToColumnRow';
use	Spreadsheet::XLSX::Reader::Cell;
use	Spreadsheet::XLSX::Reader::Types v0.1 qw(
		CellType		CellID
	);

#########1 Dispatch Tables & Package Variables    5#########6#########7#########8#########9

#~ my	$r_name_space = 'xmlns';

#########1 Public Attributes  3#########4#########5#########6#########7#########8#########9

has styles_instance =>(
	isa		=> Object,
	writer	=>'_set_styles',
	handles	=> [ qw(
		get_format_position
		get_default_format_position
		process_element_to_perl_data
	) ],
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

has error_inst =>(
		isa			=> InstanceOf[ 'Spreadsheet::XLSX::Reader::Error' ],
		handles 	=>[ qw( error set_error clear_error set_warnings if_warn ) ],
		clearer		=> '_clear_error_inst',
		reader		=> '_get_error_instance',
		required	=> 1,
	);

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

#~ has sheet_rel_id =>(
		#~ isa		=> Str,
		#~ reader	=> 'rel_id',
	#~ );

has sheet_id =>(
		isa		=> Int,
		reader	=> 'sheet_id',
	);

#~ has sheet_position =>(# XML position
		#~ isa		=> Int,
		#~ reader	=> 'position',
	#~ );

has sheet_name =>(
		isa		=> Str,
		writer	=> 'set_sheet_name',
		reader	=> 'name',
	);

has custom_formats =>(
		isa		=> HashRef[ HasMethods[ 'coerce', 'display_name' ] ],
		traits	=> ['Hash'],
		handles	=>{
			has_custom_format => 'exists',
			get_custom_format => 'get',
			set_custom_format => 'set',
		},
		writer	=> 'set_custom_formats',
		default => sub{ {} },
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
    my ( $self, $request_row, $request_column ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD					name_space 	=> $self->get_log_space .  '::get_cell', );
	###LogSD	no	warnings 'uninitialized';
	###LogSD		$phone->talk( level => 'info', message =>[
	###LogSD			"Arrived at get_cell with: ",
	###LogSD			"Requested row: " . (defined( $request_row ) ? $request_row : ''),
	###LogSD			"File row: " . (defined( $request_row ) ? $self->get_excel_position( $request_row ) : ''),
	###LogSD			"Last Row: " . (($self->_has_file_row_now) ? $self->_get_file_row_now : ''),
	###LogSD			"Requested column: " . (defined( $request_column ) ? $request_column : '' ),
	###LogSD			"File column: " . (defined( $request_column ) ? $self->get_excel_position( $request_column ) : '' ),
	###LogSD			"Last Column: " . (($self->_has_file_col_now) ? $self->_get_file_col_now : '' ),] );
	###LogSD	use	warnings 'uninitialized';
	
	# Handle an implied next column
	my	$add_row = 0;
	my	$row_column = 0;
	if( defined $request_column ){
		$request_column = $self->get_excel_position( $request_column );
		if( $request_column < 1 ){
			$self->error( 'Requested column -' . $self->get_used_position( $request_column ) . 
				'- is not an allowed position (negative or zero)' );
			$self->_clear_last_report_row;
			$self->_clear_last_report_col;
			return undef;
		}
		if( defined $request_row ){
			$row_column = 1;
		}else{
			$self->error( "Requested column -$request_column- provided but no requested row identified" );
			$self->_clear_last_report_row;
			$self->_clear_last_report_col;
			return undef;
		}
	}else{
		###LogSD	$phone->talk( level => 'info', message =>[
		###LogSD		"Handling an implied 'next column'" ] );
		$request_column = ( $self->_has_last_report_col ) ?
					( $self->_get_last_report_col + 1 ): $self->_get_file_col_now;
		if( $self->_is_past_the_last_column( $request_column ) ){
			$request_column = ( $self->has_min_col ) ? $self->min_col : 1 ;
			$add_row = 1;
		}
	}
	
	
	# Handle an implied next row
	if( defined $request_row ){
		$request_row = $self->get_excel_position( $request_row );
		if( $request_row < 1 ){
			$self->error( 'Requested row -' . $self->get_used_position( $request_row ) . 
				'- is not an allowed position (negative or zero)' );
			$self->_clear_last_report_row;
			$self->_clear_last_report_col;
			return undef;
		}
	}else{
		###LogSD	$phone->talk( level => 'info', message =>[
		###LogSD		"No row passed - using the current row" ] );
		$request_row = ( $self->_has_last_report_row ) ?
					( $self->_get_last_report_row ): $self->_get_file_row_now;
		$request_row += $add_row;
	}
	
	# Handle past upper bounds
	if( $self->_is_past_the_last_row( $request_row ) ){
		$self->_clear_last_report_row;
		$self->_clear_last_report_col;
		return 'EOF';
	}elsif( $self->_is_past_the_last_column( $request_column ) ){
		$self->_clear_last_report_row;
		$self->_clear_last_report_col;
		return 'EOR';
	}
	
	#reset the sheet as needed
	my $should_reset = 0;
	if( !$self->_has_file_row_now ){
		###LogSD	$phone->talk( level => 'debug', message =>[
		###LogSD		"Worksheet parsing not started - pulling the first cell",
		###LogSD		"Current file position:" . $self->byte_consumed ] );
		$self->_get_next_row;
		###LogSD	$phone->talk( level => 'trace', message =>[
		###LogSD		"New file position:" . $self->byte_consumed ] );
	}
	if( $self->_has_file_row_before and $request_row <= $self->_get_file_row_before ){
		###LogSD	$phone->talk( level => 'debug', message =>[
		###LogSD		"Need a previous row - resetting the sheet" ] );
		$should_reset = 1;
	}elsif( $self->_get_file_row_now == $request_row ){
		if( $self->_has_file_col_before and $request_column <= $self->_get_file_col_before ){
			###LogSD	$phone->talk( level => 'debug', message =>[
			###LogSD		"Need a previous column in this row - (still) resetting the sheet" ] );
			$should_reset = 1;
		}
	}
	if( $should_reset ){
		$self->start_the_file_over;
		$self->_get_next_row;
		###LogSD	$phone->talk( level => 'debug', message =>[
		###LogSD		"Reset sheet to read from row: " . $self->_get_file_row_now,
		###LogSD		"... and column: " . $self->_get_file_col_now ] );
	}
		
	
	# Advance to the proper row
	while( $self->_get_file_row_now < $request_row ){
		my $result = $self->_get_next_row;
		###LogSD	$phone->talk( level => 'trace', message =>[
		###LogSD		"Result of next_row: " . $result, ] );
		return $result if $result =~ /^EO/;
		###LogSD	$phone->talk( level => 'debug', message =>[
		###LogSD		"Advanced sheet to read from row: " . $self->_get_file_row_now,
		###LogSD		"... and column: " . $self->_get_file_col_now ] );
	}
	
	# Advance to the proper column
	if( $self->_get_file_row_now == $request_row ){
		while( $self->_get_file_col_now < $request_column ){
			my $result = $self->_get_next_column;
			###LogSD	$phone->talk( level => 'trace', message =>[
			###LogSD		"Result of next_column: " . $result, ] );
			return $result if $result =~ /^EO/;
			###LogSD	$phone->talk( level => 'debug', message =>[
			###LogSD		"Advanced sheet to read from column: " . $self->_get_file_col_now, ] );
		}
	}
	
	# Prepare to return an answer (finally!)
	$self->_set_last_report_row( $request_row );
	$self->_set_last_report_col( $request_column );
	
	# Handle empty rows
	if( $request_row < $self->_get_file_row_now ){
		###LogSD	$phone->talk( level => 'info', message =>[ "This row is empty!" ] );
		return undef;
	}

	# Handle empty columns
	if( $request_column < $self->_get_file_col_now ){
		###LogSD	$phone->talk( level => 'info', message =>[ "This column is empty!" ] );
		return undef;
	}

	# Return a Spreadsheet::XLSX::Reader::Cell instance with data
	my	$args ={
			log_space		=> $self->get_log_space . '::Cell',
			value_encoding	=> $self->encoding,
			value_type		=> $self->_get_column_t,
			cell_column		=> $self->get_used_position( $self->_get_file_col_now ),
			cell_row		=> $self->get_used_position( $self->_get_file_row_now ),
		};
	if( !$self->_has_xml_parser ){
		die "Should have an XML parser at this point";
	}
	my	$cell_node = $self->copy_current_node( 1 );
	###LogSD	$phone->talk( level => 'trace', message =>[ "The cell node:", $cell_node->toString, "Byte position: " . $self->byte_consumed ] );
	my ( $unformatted, $rich_text_content );
	if( !$cell_node->hasChildNodes() ){
		###LogSD	$phone->talk( level => 'trace',
		###LogSD		 message => [ "Cell active but no child nodes - the unformatted value will be the empty string" ] );
		$unformatted = '';
	}else{
		my	$formula_nodes	= $cell_node->getChildrenByTagName( 'f' );
		if( $formula_nodes->size() ){
			###LogSD	$phone->talk( level => 'trace',
			###LogSD		 message => [ "Found a formula node" ] );
			$args->{cell_formula} = $formula_nodes->get_node(0)->textContent;
		}
		my	$worksheet_content	= $cell_node->getChildrenByTagName( 'v' )->get_node(0)->textContent;
		###LogSD	$phone->talk( level => 'debug', message =>[
		###LogSD		"Generated the initial text content of:", $worksheet_content ] );
		if( $self->_get_column_t eq 's' ){
			my	$string_node = $self->get_shared_string_position( $worksheet_content );
			###LogSD	$phone->talk( level => 'info', message => [
			###LogSD		"Loading shared strings position: $worksheet_content",
			###LogSD		"with data: " . $string_node->textContent, ] );
			for my $node ( $string_node->getChildrenByTagName( 'r' ) ){#$rich_text_nodes->get_nodelist
				###LogSD	$phone->talk( level => 'info', message => [
				###LogSD		"Working on r subnode: " . $node->toString, ] );
				my	$position = length( $unformatted );
					$position //= 0;
				###LogSD	$phone->talk( level => 'info', message => [
				###LogSD		"Working on r subnode: " . $node->toString,
				###LogSD		"At position: $position"] );
				for my $rt_node ( $node->getChildrenByTagName( 'rPr' ) ){
					###LogSD	$phone->talk( level => 'info', message => [
					###LogSD		"Working on rPr subnode: " . $rt_node->toString, ] );
					my	$rich_text_element = $self->process_element_to_perl_data( $rt_node );
					#~ for my $child ( $rt_node->childNodes ){
						#~ ###LogSD	$phone->talk( level => 'info', message => [
						#~ ###LogSD		"Working on child node: " . $child->toString, ] );
						#~ my	$node_name = $child->nodeName;
						#~ if( defined $child->nodeValue ){
							#~ $rich_text_element->{$node_name} = $child->nodeValue;
						#~ }else{
							#~ my	$ref;
							#~ for my $attribute ( $child->attributes ){
								#~ ###LogSD	$phone->talk( level => 'info', message => [
								#~ ###LogSD		"parsing attribute ref: " . ref $attribute,
								#~ ###LogSD		"for attribute name: " . $attribute->get_name,
								#~ ###LogSD		"and attribute value: " . $attribute->value,
								#~ ###LogSD		"parsing attribute: " . $attribute->toString, ] );
								#~ if( $attribute->name eq 'val' ){
									#~ $rich_text_element->{$node_name} = $attribute->value;
								#~ }else{
									#~ $rich_text_element->{$node_name}->{$attribute->name} = $attribute->value;
								#~ }
							#~ }
						#~ }
						#~ $rich_text_element->{$node_name} //= 1;
						###LogSD	$phone->talk( level => 'info', message => [
						###LogSD		"current rich text element : ", $rich_text_element, ] );
					#~ }
					push @$rich_text_content, [ $position, $rich_text_element ];
					###LogSD	$phone->talk( level => 'info', ask => '',
					###LogSD		 message =>[ "Rich text to this point:" , $rich_text_content	] );
				}
				$unformatted .= $node->textContent;
				
			}
			$unformatted //= $string_node->textContent;
			###LogSD	$phone->talk( level => 'info', message => [
			###LogSD		"Unformatted value: $unformatted"	] );
		}elsif( $self->_get_column_t  eq 'number' ){
			###LogSD	$phone->talk( level => 'trace',
			###LogSD		 message => [ "No action needed for numbers ..." ] );
			$unformatted = $worksheet_content;
		}else{
			$self->_set_error(
				'The cell in row -' . $self->get_used_position( $self->_get_file_row_now ) . 
				'- column -' . $self->get_used_position( $self->_get_file_col_now ) . 
				"- has an unrecognized value for attribute 't': " . $self->_get_column_t );
		}
	}
	$args->{raw_value} = $unformatted;
	$args->{rich_text} = $rich_text_content if $rich_text_content;
	###LogSD	$phone->talk( level => 'trace', ask => "continue?",
	###LogSD		 message => [ "unformatted to this point: $unformatted", "Byte position: " . $self->byte_consumed ] );
	my $merge_row = $self->_get_row_merge_map( $self->_get_file_row_now  );
	if( $merge_row and $merge_row->[$self->_get_file_col_now]){
		$args->{merge_range} = $merge_row->[$self->_get_file_col_now];
	}
	my	$format_definition;
	if( $self->_has_column_s ){
		$format_definition = $self->get_format_position( $self->_get_column_s );
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"The defined format position is: " . $self->_get_column_s,
		###LogSD		"With data:", $format_definition ] );
	}else{
		$format_definition = $self->get_default_format_position(  );#'NumberFormat'
		###LogSD	$phone->talk( level => 'debug',
		###LogSD		 message => [ "Using the default number format", $format_definition ] );
	}
	map{ $args->{$_} = $format_definition->{$_} }( keys %$format_definition );
	my @format_list = ($self->_get_column_id =~ /^(([A-Za-z]+)(\d+))$/);
	for my $callout ( @format_list ){
		if( $self->has_custom_format( $callout ) ){
			$args->{NumberFormat} = $self->get_custom_format( $callout );
			last;
		}
	}
	#~ $args->{value_coercion} = $format_definition;
	$args->{error_inst} = $self->_get_error_instance;
	$args->{count_from_zero} = $self->counting_from_zero;
	###LogSD	$phone->talk( level => 'trace',#ask => "continue?" 
	###LogSD		 , message => [ "Updated cell data is:", $args, ] );
	my	$cell_instance = Spreadsheet::XLSX::Reader::Cell->new( %$args );
	###LogSD	$phone->talk( level => 'fatal', message => [#, ask => "continue?"
	###LogSD		"Built Cell:", $cell_instance ] );
	return $cell_instance;
}


#########1 Private Attributes 3#########4#########5#########6#########7#########8#########9
	
has +_core_element =>(
		default => 'row',
	);

has _last_report_row =>(
		isa			=> Int,
		reader		=> '_get_last_report_row',
		writer		=> '_set_last_report_row',
		clearer		=> '_clear_last_report_row',
		predicate	=> '_has_last_report_row',
	);

has _file_row_now =>(
		isa			=> Int,
		reader		=> '_get_file_row_now',
		writer		=> '_set_file_row_now',
		clearer		=> '_clear_file_row_now',
		predicate	=> '_has_file_row_now',
	);

has _row_span_start =>(
		isa			=> Int,
		reader		=> '_get_row_span_start',
		writer		=> '_set_row_span_start',
		clearer		=> '_clear_row_span_start',
		predicate	=> '_has_row_span_start',
	);

has _row_span_end =>(
		isa			=> Int,
		reader		=> '_get_row_span_end',
		writer		=> '_set_row_span_end',
		clearer		=> '_clear_row_span_end',
		predicate	=> '_has_row_span_end',
	);

has _file_row_before =>(
		isa			=> Int,
		reader		=> '_get_file_row_before',
		writer		=> '_set_file_row_before',
		clearer		=> '_clear_file_row_before',
		predicate	=> '_has_file_row_before',
	);

has _last_report_col =>(
		isa			=> Int,
		reader		=> '_get_last_report_col',
		writer		=> '_set_last_report_col',
		clearer		=> '_clear_last_report_col',
		predicate	=> '_has_last_report_col',
	);

has _file_col_now =>(
		isa			=> Int,
		reader		=> '_get_file_col_now',
		writer		=> '_set_file_col_now',
		clearer		=> '_clear_file_col_now',
		predicate	=> '_has_file_col_now',
	);

has _column_id =>(
		isa			=> CellID,
		reader		=> '_get_column_id',
		writer		=> '_set_column_id',
		clearer		=> '_clear_column_id',
		predicate	=> '_has_column_id',
	);

has _column_t =>(
		isa			=> CellType,
		reader		=> '_get_column_t',
		writer		=> '_set_column_t',
		clearer		=> '_clear_column_t',
		predicate	=> '_has_column_t',
	);

has _column_s =>(
		isa			=> Int,
		reader		=> '_get_column_s',
		writer		=> '_set_column_s',
		clearer		=> '_clear_column_s',
		predicate	=> '_has_column_s',
	);

has _file_col_before =>(
		isa			=> Int,
		reader		=> '_get_file_col_before',
		writer		=> '_set_file_col_before',
		clearer		=> '_clear_file_col_before',
		predicate	=> '_has_file_col_before',
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
			#~ get_row_number	=>[ getAttribute => 'r' ],
			#~ get_row_span	=>[ getAttribute => 'spans' ],
			_get_all_cells	=>[ getChildrenByTagName => 'c' ],
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

has	_merge_map =>(
		isa		=> ArrayRef,
		traits	=> ['Array'],
		writer	=> '_set_merge_map',
		handles	=>{
			_get_row_merge_map => 'get',
		},
	);

has _sheet_unique_count =>(
	isa			=> Int,
	writer		=> '_set_unique_count',
	clearer		=> '_clear_unique_count',
	predicate	=> '_has_unique_count',
	reader		=> '_get_unique_count',
);


#########1 Private Methods    3#########4#########5#########6#########7#########8#########9

sub _load_unique_bits{
	my( $self, ) = @_;#, $new_file, $old_file
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD					name_space 	=> ($self->get_log_space . '::_load_unique_bits' ), );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"Setting the Worksheet unique bits", "Byte position: " . $self->byte_consumed ] );
	
	# Read the sheet dimensions
	if( $self->_next_element( 'dimension' ) ){
		my	$range = $self->_get_attribute( 'ref' );
		my	( $start, $end ) = split( /:/, $range );
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"Start position: $start", "End position: $end", "Byte position: " . $self->byte_consumed ] );
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
	
	#build a merge map
	my	$merge_ref = [];
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"Loading the mergeCell" ] );
	while( $self->_next_element('mergeCell') ){
		my	$merge_range = $self->_get_attribute( 'ref' );
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"Loading the merge range for: $merge_range", "Byte position: " . $self->byte_consumed ] );
		my( $start, $end ) = split /:/, $merge_range;
		my( $start_col, $start_row ) = $self->parse_column_row( $start );
		my( $end_col, $end_row ) = $self->parse_column_row( $end );
		my( $min_col, $max_col ) = ($start_col,$end_col);
		while ( $start_row <= $end_row ){
			$merge_ref->[$start_row]->[$start_col] = $merge_range;
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

sub _get_next_row{
	my ( $self, )	= @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD					$self->get_log_space . '::_get_next_row', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"Getting the next row element" ] );
	if( $self->_has_file_row_now ){
		$self->_set_file_row_before( $self->_get_file_row_now );
	}
	my $reader = $self->_get_xml_parser;
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"Reader:", $reader, "Byte position: " . $self->byte_consumed ] );
	my $found_it = $reader->nextElement( 'row' );
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"Reader:", $reader, "Byte position: " . $self->byte_consumed ] );
	if( $found_it > 0 ){
		my $row_num = $reader->getAttribute( 'r' );
		$self->_set_file_row_now( $row_num );
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"Found row: $row_num" ] ),
		my $span = $reader->getAttribute( 'spans' );
		if( $span ){
			my ( $start_col, $end_col ) = ( $span ) ? (split /:/, $span ) : ( undef, undef );
			$self->_set_row_span_start( $start_col ) if $start_col;
			$self->_set_row_span_end( $end_col ) if $end_col;
			###LogSD	$phone->talk( level => 'debug', message => [
			###LogSD		"It starts at column: $start_col", "And ends at column: $end_col" ] );
		}else{
			$self->_clear_row_span_start;
			$self->_clear_row_span_end;
			$self->_clear_file_row_now;
		}
		$self->_clear_file_col_before;
		$self->_clear_file_col_now;
		$self->_clear_column_id;
		$self->_clear_column_s;
		$self->_clear_column_t;
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"Reader", $reader->copyCurrentNode(1)->toString ] );
		my $found_it = $reader->nextElement( 'c' );
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"Found it: $found_it",
		###LogSD		"Cell", $reader->copyCurrentNode(1)->toString, "Byte position: " . $reader->byteConsumed ] );
		return ( $found_it > 0 ) ? $self->_read_cell_attributes( $reader ) : 'EOR';
	}else{
		###LogSD	$phone->talk( level => 'warn', message => [
		###LogSD		"Reached the EOF", ] );
		return 'EOF';
	}
}

sub _get_next_column{
	my ( $self, )	= @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD					$self->get_log_space . '::_get_next_column', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"Getting the next column element" ] );
	if( $self->_has_file_col_now ){
		$self->_set_file_col_before( $self->_get_file_col_now );
	}
	my $reader = $self->_get_xml_parser;
	my $found_it = $reader->nextSibling;
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"Found it: $found_it",
	###LogSD		"Cell", $reader->copyCurrentNode(1)->toString, "Byte position: " . $reader->byteConsumed ] );
	if( $found_it < 1 ){
		###LogSD	$phone->talk( level => 'info', message => [
		###LogSD		"Reached the EOR", ] );
		$self->_clear_file_col_now;
		$self->_clear_column_id;
		$self->_clear_column_s;
		$self->_clear_column_t;
		return 'EOR';
	}else{
		return $self->_read_cell_attributes( $reader );
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"Found it: $found_it",
		###LogSD		"Cell", $reader->copyCurrentNode(1)->toString, "Byte position: " . $reader->byteConsumed ] );
	}
}

sub _read_cell_attributes{
	my ( $self, $reader )	= @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD					$self->get_log_space . '::_get_next_column::_read_cell_attributes', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"Getting the cell attributes with reader", $reader->copyCurrentNode(1)->toString ] );
	my $cell_id = $reader->getAttribute( 'r' );
	###LogSD	$phone->talk( level => 'debug', message =>[ "Cell ID: $cell_id" ] );
	$self->_set_column_id( $cell_id );
	my ( $cell_col, $cell_row ) = $self->parse_column_row( $cell_id, 1 );
	###LogSD	$phone->talk( level => 'debug', message =>[ 
	###LogSD		"Column: $cell_col", "Row: $cell_row", "File row now:" . $self->_get_file_row_now ] );
	if( $cell_row != $self->_get_file_row_now ){
		###LogSD	$phone->talk( level => 'debug', message =>[ 
		###LogSD		"Malformed xml worksheet with mismatched cell row and row row - " . $reader->readOuterXml() ] );
		die "Malformed xml worksheet with mismatched cell row and row row - " . $reader->readOuterXml();
	}
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"processing cell ID: $cell_id", "As Excel column: $cell_col", "And Excel row: $cell_row" ] );
	$self->_set_file_col_now( $cell_col );
	my $cell_type = $reader->getAttribute( 't' );
	$cell_type = 'number' if !$cell_type;
	$self->_set_column_t( $cell_type );
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"Cell type: $cell_type" ] );
	my $cell_format = $reader->getAttribute( 's' );
	if( $cell_format ){
		$self->_set_column_s( $cell_format );
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"Cell format: $cell_format" ] );
	}else{
		$self->_clear_column_s;
	}
	$self->_set_xml_parser( $reader );
	return 1;
}

sub _is_past_the_last_row{
	my ( $self, $request_row )	= @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD					$self->get_log_space . '::_is_past_the_last_row', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"Checking if -$request_row- is past the last row" ] );
	if( $self->has_max_row ){
		return 1 if $self->get_used_position( $request_row ) > $self->max_row;
	}
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"The row is within range" ] );
	return 0;
}

sub _is_past_the_last_column{
	my ( $self, $request_col )	= @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD					$self->get_log_space . '::_is_past_the_last_column', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"Checking if -$request_col- is past the last column" ] );
	if( $self->has_max_col ){
		return 1 if $self->get_used_position( $request_col ) > $self->max_col;
	}elsif( $self->_has_row_span_end ){
		return 1 if $request_col > $self->_get_row_span_end;
	}
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"The column is within range" ] );
	return 0;
}
		
		
		
		
#~ augment 'get_position' => sub{
	#~ my ( $self, )	= shift;
	#~ my	$position 	= $self->_get_requested_position;
	#~ ###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	#~ ###LogSD					$self->get_log_space . '::get_position::augmented', );
	#~ ###LogSD		$phone->talk( level => 'debug', message => [
	#~ ###LogSD			"Reached augment::_get_position for: $position" ] );
	
	#~ #checking if the reqested position is too far
	#~ if( $position > $self->_get_unique_count ){
		#~ ###LogSD	$phone->talk( level => 'warn', message => [
		#~ ###LogSD		"Asking for position -$position- (from 0) but the worksheet " .
		#~ ###LogSD		"max row position is: " . ($self->_get_unique_count) ] );
		#~ return 1;#  fail
	#~ }else{
		#~ ###LogSD	$phone->talk( level => 'info', message =>[ "No end in sight" ] );
		#~ return undef;#No failure
	#~ }
#~ };

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

Spreadsheet::XLSX::Reader::XMLReader::Worksheet - A class for exploring XLSX worksheets

=head1 SYNOPSIS

See the SYNOPSIS in L<Spreadsheet::XLSX::Reader>
    
=head1 DESCRIPTION

This is the class used to interrogate Excel xlsx worksheets for information.  Because 
the Excel xlsx storage of information can be (but not always) spread across multiple xml 
files this class is the way to retreive information about the sheet independant of which 
file it is in.  This is the L<XMLReader|Spreadsheet::XLSX::Reader::XMLReader> version of 
this class.  Where possible this class makes the decision to read files by line rather 
than parse the file using a DOM tree.  The up side to this is that large files will 
(hopefully) not crash when opened and the data will available using the same methods.  
The down side is that the file opens slower since the whole sheet is read twice before 
information is available.  Additionally data is best accessed sequentially left to right 
and top to bottom since going back will involve a system file close and re-open action.

=head2 Attributes

Attributes of this cell are not included in the documentation because 'new' should be 
called by L<other|Spreadsheet::XLSX::Reader> classes in this package.

=head2 Methods

These are ways to access the data and formats in the cell.  They also provide a 
way to modifiy the output of the format.

=head3 get_name

=over

B<Definition:> Returns the sheet name

B<Accepts:>Nothing

B<Returns:> $sheet_name

=back

=head3 row_range

=over

B<Definition:> Returns the minimum row number and the maximum row number based on the 
settings of attribute 'count_from_zero' set when first creating the file parser with 
L<Spreadsheet::XLSX::Reader>.

B<Accepts:>Nothing

B<Returns:> a list of ( $minimum_row, $max_row )

=back

=head3 col_range

=over

B<Definition:> Returns the minimum column number and the maximum column number based on 
the settings of attribute 'count_from_zero' set when first creating the file parser with 
L<Spreadsheet::XLSX::Reader>.

B<Accepts:>Nothing

B<Returns:> a list of ( $minimum_column, $max_column )

=back

=head3 get_cell( $row, $column )

=over

B<Definition:> Returns a L<Spreadsheet::XLSX::Reader::Cell> instance corresponding to the 
identified $row and $column.  The actual position returned are affected by the attribute 
'count_from_zero' set when first creating the file parser with L<Spreadsheet::XLSX::Reader>.  
If there is no data stored in that cell it returns undef.  If the $column selected is past 
the L<max_col|/max_col> value then it returns the string 'EOR'.  If the $row selected is 
past the L<max_row|/max_row> value then it returns the string 'EOF'.

If both $row and $column are left blank this is effectivly a 'get_(next)_cell' command 
moving left to right and top to bottom starting from either the last position identified 
or the sheet minimum row and column.  When the end of the file is reached it returns the 
string 'EOF';

If only the $row is specified it will return the next cell in that $row starting from the 
last column specified I<even if it was a different row> or starting from the minimum column.  
When the row is finished it will return the string 'EOR' and reset the next column to be 
the minimum column.  (The sheet starts with the minimum column as the next column on 
opening.)  This implementation is pre-deprecated and will be removed when the 
'fetchrow_arrayref' function is implemented.  That that point this method will require 
either both $row and $column or neither.

If only the $column is specified this will return undef and set the 
L<error|Spreadsheet::XLSX::Reader::Error> message returning undef.  

B<Accepts:> ( $row, $column ) - as indicated in the Definition

B<Returns:> (undef|a blessed L<Spreadsheet::XLSX::Reader::Cell> instance|'EOR'|'EOF')

=back
		
=head3 set_warnings( $bool )

=over

B<Definition:> Turn clucked warnings on or off from L<Spreadsheet::XLSX::Reader::Error>

B<Accepts:> Boolean values

B<Returns:> nothing

=back
		
=head3 if_warn

=over

B<Definition:> Check the state of the boolean affected by L<set_warnings
|/set_warnings( $bool )> attribute value from L<Spreadsheet::XLSX::Reader::Error>

B<Accepts:> Nothing

B<Returns:> $bool

=back
		
=head3 error

=over

B<Definition:> Returns the currently stored error string from 
L<Spreadsheet::XLSX::Reader::Error>

B<Accepts:> Nothing

B<Returns:> $error_string

=back

=head3 clear_error

=over

B<Definition:> method to clear the current error string from 
L<Spreadsheet::XLSX::Reader::Error>

B<Accepts:> Nothing

B<Returns:> Nothing (string is cleared)

=back

=head3 min_col

=over

B<Definition:> method to read the minimum column for the sheet

B<Accepts:> Nothing

B<Returns:> $minimum_column (Integer)

=back

=head3 has_min_col

=over

B<Definition:> indicates if a minimum column has been determined

B<Accepts:> Nothing

B<Returns:> $bool TRUE = exists

=back

=head3 min_row

=over

B<Definition:> method to read the minimum row for the sheet

B<Accepts:> Nothing

B<Returns:> $minimum_row (Integer)

=back

=head3 has_min_row

=over

B<Definition:> indicates if a minimum row has been determined

B<Accepts:> Nothing

B<Returns:> $bool TRUE = exists

=back

=head3 max_col

=over

B<Definition:> method to read the maximum column for the sheet

B<Accepts:> Nothing

B<Returns:> $maximum_column (Integer)

=back

=head3 has_max_col

=over

B<Definition:> indicates if a maximum column has been determaxed

B<Accepts:> Nothing

B<Returns:> $bool TRUE = exists

=back

=head3 max_row

=over

B<Definition:> method to read the maximum row for the sheet

B<Accepts:> Nothing

B<Returns:> $maximum_row (Integer)

=back

=head3 has_max_row

=over

B<Definition:> indicates if a maximum row has been determaxed

B<Accepts:> Nothing

B<Returns:> $bool TRUE = exists

=back

=head3 set_custom_formats( $hashref )

=over

B<Definition:> It is not inconceivable that the module user would need/want the data 
manipulated in some way that was not provided natively by excel.  This package uses 
the excellent L<Type::Tiny> to implement the default data manipulations identified 
by the spreadsheet.  However, it is possible for the user to supply a hashref of 
custom data manipulations.  The hashref is read where the key is a row-column 
indicator and the value is a data manipulation coderef/object that has (at least) 
the following two methods.  The first method is 'coerce' and the second method is 
'display_name'.  For each cell instance generated the L<get_cell
|/get_cell( $row, $column)> method will check the cell_id (ex. B34) for matches in this 
hashref and then if none are found it will apply any format(data manipulation) defined 
in the spreadsheet. For a match on any given cell checks will be done in this order; 
full cell_id (ex. B34), column_id (ex. B), row_id (ex.34)

B<Accepts:> a $hashref (ex. { B34 => MyTypeTinyType->plus_coercions( MyCoercion ) } )

B<Returns:> Nothing

=back

=head3 set_custom_format( $key => $value_ref )

=over

B<Definition:> The difference with this method from L<set_custom_formats
|/set_custom_formats( $hashref )> is this will only set specific key value pairs.

B<Accepts:> a $key => $value_ref list

B<Returns:> Nothing

=back

=head3 get_custom_format( $key )

=over

B<Definition:> This returns the custom format associated with that key

B<Accepts:> a $key

B<Returns:> The $value_ref (data manipulation ref) associated with $key

=back

=head3 has_custom_format( $key )

=over

B<Definition:> This checks if a custom format is registered against the $key

B<Accepts:> a $key

B<Returns:> $boolean representing existance

=back
				
=head1 SUPPORT

=over

L<github Spreadsheet::XLSX::Reader/issues|https://github.com/jandrew/Spreadsheet-XLSX-Reader/issues>

=back

=head1 TODO

=over

B<1.> Add an attribute and supporting alternate read path that avoids pre-mapping the 
cells so that less is in memory at any one time - for extremly large files - 
expected to slow performance

B<2.> Add 'fetchrow_arrayref( $row )' (as a Role?)

B<3.> Add 'set_header_row( $row )' and 'fetchrow_hashref( $row )' (as a Role?)

B<4.> Add L<Data::Walk::Graft> capabilities to 'set_custom_formats'

B<5.> Move 'get_cell( $row, $column )' into a role?

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

This software is copyrighted (c) 2014 by Jed Lund

=head1 DEPENDENCIES

=over

B<5.010> - (L<perl>)

L<version>

L<Moose>

L<MooseX::StrictConstructor>

L<MooseX::HasDefaults::RO>

L<Types::Standard>

L<Spreadsheet::XLSX::Reader::XMLReader>

L<Spreadsheet::XLSX::Reader::CellToColumnRow>

L<Spreadsheet::XLSX::Reader::Cell>

=back

=head1 SEE ALSO

=over

L<Spreadsheet::XLSX>

L<Spreadsheet::ParseExcel::Worksheet>

L<Log::Shiras|https://github.com/jandrew/Log-Shiras> - to activate the debug logging

=back

=cut

#########1 Documentation End  3#########4#########5#########6#########7#########8#########9