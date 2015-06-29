package Spreadsheet::XLSX::Reader::LibXML::GetCell;
use version; our $VERSION = qv('v0.38.2');

use Carp 'confess';
use	Moose::Role;
requires qw(
	min_row						max_row						min_col
	max_col						row_range					col_range
	_get_col_row				_get_next_value_cell		_get_row_all
);
###LogSD	requires 'get_log_space';
use Types::Standard qw(
	Bool 						HasMethods					Enum
	Int							is_Int						ArrayRef
	is_ArrayRef					HashRef						is_HashRef
	is_Object
);# Int
use lib	'../../../../../lib',;
###LogSD	use Log::Shiras::Telephone;
###LogSD	use Log::Shiras::UnhideDebug;
use	Spreadsheet::XLSX::Reader::LibXML::Cell;

#########1 Dispatch Tables    3#########4#########5#########6#########7#########8#########9

my $format_headers ={
		fonts			=> 'cell_font',
		borders			=> 'cell_border',
		cellStyleXfs	=> 'cell_style',
		fills			=> 'cell_fill',
		numFmts			=> 'cell_coercion',
	};

#########1 Public Attributes  3#########4#########5#########6#########7#########8#########9

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
#################################################
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

#########1 Public Methods     3#########4#########5#########6#########7#########8#########9

sub get_cell{
    my ( $self, $requested_row, $requested_column ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD					name_space 	=> $self->get_log_space .  '::get_cell', );
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
	my $result = $self->get_col_row( $requested_column, $requested_row );
	
	# Handle EOF EOR flags
	my $return = undef;
	if( $result and is_HashRef( $result ) ){
		$return = $self->_build_out_the_cell( $result );
	}elsif( $result and $self->boundary_flag_setting ){
		$return = $result;
	}
	
	return $return;
}

sub get_next_value{
    my ( $self, ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD					name_space 	=> $self->get_log_space .  '::get_next_value', );
	###LogSD		$phone->talk( level => 'info', message =>[ 'Arrived at get_next_value', ] );
	
	my $result = $self->_get_next_value_cell;
	###LogSD	$phone->talk( level => 'info', message =>[ 'Next value;', $result ] );
	
	# Handle EOF EOR flags
	my $return = undef;
	if( $result and is_HashRef( $result ) ){
		$return = $self->_build_out_the_cell( $result );
	}elsif( $result and $self->boundary_flag_setting ){
		$return = $result;
	}
	
	return $return;
}

sub fetchrow_arrayref{
    my ( $self, $row ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD					name_space 	=> $self->get_log_space .  '::fetchrow_arrayref', );
	###LogSD		$phone->talk( level => 'info', message =>[
	###LogSD			"Arrived at fetchrow_arrayref for row: " . ((defined $row) ? $row : ''), ] );
	
	my $result = $self->get_row_all( $row );
	###LogSD	$phone->talk( level => 'debug', message =>[ 'Returned row ref;', $result ] );
	my $return = undef;
	if( $result and is_ArrayRef( $result ) ){
		for my $cell ( @$result ){
			if( is_HashRef( $cell ) ){
				###LogSD	$phone->talk( level => 'debug', message =>[
				###LogSD		'Building out the cell:', $cell ] );
				push @$return, $self->_build_out_the_cell( $cell, );
			}else{
				push @$return, $cell;
			}
		}
	}elsif( $result and $self->boundary_flag_setting ){# Handle EOF EOR flags
		$return = $result;
	}
	###LogSD	$phone->talk( level => 'trace', message =>[ 'Final return:', $return ] );
	
	return $return;
}

sub fetchrow_array{
    my ( $self, $row ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD					name_space 	=> $self->get_log_space .  '::fetchrow_array', );
	###LogSD		$phone->talk( level => 'info', message =>[
	###LogSD			"Arrived at fetchrow_array for row: " . ((defined $row) ? $row : ''), ] );
	my $array_ref = $self->fetchrow_arrayref( $row );
	
	return is_ArrayRef( $array_ref ) ? @$array_ref : $array_ref;
}

sub set_headers{
    my ( $self, @header_row_list ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD					name_space 	=> $self->get_log_space .  '::set_headers', );
	###LogSD		$phone->talk( level => 'info', message =>[
	###LogSD			"Arrived at set_headers for row: ", @header_row_list, ] );
	my $header_ref;
	$self->_clear_last_header_row;
	$self->_clear_header_ref;
	my $old_output = $self->get_group_return_type;
	if( $old_output eq 'instance' ){
		$self->set_group_return_type( 'value' );
	}else{
		$old_output = undef;
	}
	if( scalar( @header_row_list ) == 0 ){
		$self->set_error( "No row numbers passed to use as headers" );
		return undef;
	}
	my $last_header_row = 0;
	for my $row ( @header_row_list ){
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
	$self->_set_last_header_row( $last_header_row );
	$self->_set_header_ref( $header_ref );
	$self->set_group_return_type( $old_output ) if $old_output;
	return $header_ref;
}

sub fetchrow_hashref{
    my ( $self, $row ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD					name_space 	=> $self->get_log_space .  '::fetchrow_hashref', );
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
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD					name_space 	=> $self->get_log_space .  '::set_custom_formats', );
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

#########1 Private Methods    3#########4#########5#########6#########7#########8#########9

around [ qw( build_cell_label get_col_row get_row_all ) ] => sub{
	my ( $method, $self, @input_list ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD					($self->get_log_space .  '::scrubbing_input' ), );
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
	###LogSD					($self->get_log_space .  '::scrubbing_output' ), );
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
	###LogSD					($self->get_log_space .  '::_get_excel_position' ), );
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
	###LogSD					($self->get_log_space .  '::_get_used_position' ), );
	###LogSD		$phone->talk( level => 'debug', message =>[
	###LogSD			"Converting Excel  -$excel_int- to the used number" ] );
	my	$used_position = $excel_int;
	$used_position -= 1 if $self->counting_from_zero;
	###LogSD		$phone->talk( level => 'debug', message =>[
	###LogSD			"The used position is: $used_position" ] );
	return $used_position;
}

sub _build_out_the_cell{
	my ( $self, $result, ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD					($self->get_log_space .  '::_build_out_the_cell' ), );
	###LogSD		$phone->talk( level => 'debug', message =>[  
	###LogSD			 "Building out the cell ref:", $result, "..with results as: ". $self->get_group_return_type ] );
	my $return;
	if( is_HashRef( $result ) ){
		###LogSD	$phone->talk( level => 'debug', message =>[
		###LogSD		"processing cell object from cell ref:", $result ] );
		$return->{cell_type} = 'Text';
		$return->{r} = $result->{r};
		$return->{cell_merge} = $result->{cell_merge} if exists $result->{cell_merge};
		if( exists $result->{t} and $result->{t} eq 's' ){# Test for all in one sheet here!(future)
			my $position = $self->get_shared_string_position( $result->{v}->{raw_text} );
			###LogSD	$phone->talk( level => 'debug', message =>[
			###LogSD		"Shared strings returned:",  $position] );
			@$return{qw( cell_unformatted rich_text )} = ( $position->{raw_text}, $position->{rich_text} );
			delete $return->{rich_text} if !$return->{rich_text};
			###LogSD	$phone->talk( level => 'debug', message =>[
			###LogSD		"Updated return:",  $return] );
		}else{
			###LogSD	$phone->talk( level => 'debug', message =>[
			###LogSD		"Setting unformatted from:", $result ] );
			$return->{cell_unformatted} = $result->{v}->{raw_text};
			$return->{cell_type} = 'Numeric' if $return->{cell_unformatted} and $return->{cell_unformatted} ne '';
		}
		if( !$return->{cell_unformatted} and $self->get_empty_return_type eq 'empty_string' ){
			###LogSD	$phone->talk( level => 'debug', message =>[ "(Re)setting undef to ''"] );
			$return->{cell_unformatted} = '';
		}
		###LogSD	$phone->talk( level => 'debug', message =>[
		###LogSD		"Cell raw text is:", $return->{cell_unformatted}] );
		$return->{cell_unformatted} = $self->change_output_encoding( $return->{cell_unformatted} );
		###LogSD	$phone->talk( level => 'debug', message =>[
		###LogSD		"With output encoding changed: " . $return->{cell_unformatted} ] ) if $return->{cell_unformatted};
		if( $self->get_group_return_type eq 'unformatted' ){
			###LogSD	$phone->talk( level => 'debug', message =>[
			###LogSD		"Sending back just the unformatted value: " . ($return->{cell_unformatted}//'') ] ) ;
			return $return->{cell_unformatted}
		}
		# Get any relevant custom format
		my	$custom_format;
		#~ ###LogSD	$phone->talk( level => 'trace', message =>[
		#~ ###LogSD		"custom formats", $self->get_custom_formats ] );
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
		# First check for return of value only
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
		if( exists $result->{f} ){
			$return->{cell_formula} = $result->{f}->{raw_text};
		}
		###LogSD	$phone->talk( level => 'debug', message =>[
		###LogSD		"Cell args to this point:", $return] );
		
		if( exists $result->{s} ){
			my $header = ($self->get_group_return_type eq 'value') ? 'numFmts' : undef;
			my $exclude_header = ($custom_format) ? 'numFmts' : undef;
			my $format;
			if( $self->_has_styles_file ){
				$format = $self->get_format_position( $result->{s}, $header, $exclude_header );
				
				###LogSD	$phone->talk( level => 'trace', message =>[
				###LogSD		"format position is:", $format ] );
			}
			# Second check for value only
			if( $self->get_group_return_type eq 'value' ){
				###LogSD	$phone->talk( level => 'debug', message =>[
				###LogSD		'Applying (a possible) regular format to: ' .  $return->{cell_unformatted} ] );
				return	Spreadsheet::XLSX::Reader::LibXML::Cell->_return_value_only(
							$return->{cell_unformatted}, 
							$format->{numFmts},
							$self->_get_error_inst,
				###LogSD	$self->get_log_space,
						);
			}
			if( $self->_has_styles_file ){
				for my $header ( keys %$format_headers ){
					if( exists $format->{$header} ){
						###LogSD	$phone->talk( level => 'trace', message =>[
						###LogSD		"Transferring styles header -$header- to cell attribute: $format_headers->{$header}", ] );
						if( $header eq 'numFmts' ){
							$return->{cell_coercion} = $format->{$header};
							if(	$return->{cell_type} eq 'Numeric' and
								$format->{$header}->name =~ /date/i ){
								###LogSD	$phone->talk( level => 'trace', message =>[
								###LogSD		"Found a -Date- cell", ] );
								$return->{cell_type} = 'Date';
							}
						}else{
							$return->{$format_headers->{$header}} = $format->{$header};
						}
					}
				}
			}
		}
		###LogSD	$phone->talk( level => 'trace', message =>[
		###LogSD		"Checking return type: " . $self->get_group_return_type,  ] );
		# Final check for value only
		if( $self->get_group_return_type eq 'value' ){
			###LogSD	$phone->talk( level => 'debug', message =>[
			###LogSD		'Applying (a possible) regular format to: ' .  $return->{cell_unformatted} ] );
			return	Spreadsheet::XLSX::Reader::LibXML::Cell->_return_value_only(
						$return->{cell_unformatted}, 
						$return->{cell_coercion},
						$self->_get_error_inst,
			###LogSD	$self->get_log_space,
					);
		}
		$return->{cell_row} = $self->_get_used_position( $result->{row} );
		$return->{cell_col} = $self->_get_used_position( $result->{col} );
		$return->{error_inst} = $self->_get_error_inst;
		###LogSD	$result->{log_space} = $self->get_log_space . '::Cell';
		#~ $return->{unformatted_converter} = sub{ 
			#~ my	$string = $_[0];
			#~ ###LogSD	$phone->talk( level => 'debug', message =>[
			#~ ###LogSD		"Sending stuff to converter:", @_ ] );
			#~ $self->change_output_encoding( $string ) 
		#~ };
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

# THE NEXT TWO METHODS ARE USED FOR ABSTRACTION AND DO NOT DO WHAT YOU THINK THEY DO
sub get_col_row{ my $self = shift; $self->_get_col_row( @_ ); };

sub get_row_all{ my $self = shift; $self->_get_row_all( @_ ); };

#########1 Phinish            3#########4#########5#########6#########7#########8#########9

no Moose::Role;
1;

#########1 Documentation      3#########4#########5#########6#########7#########8#########9
__END__

=head1 NAME

Spreadsheet::XLSX::Reader::LibXML::GetCell - Top level xlsx Worksheet interface

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

=head3 set_headers( @header_row_list )

=over

B<Definition:> This function is used to set headers used in the function 
L<fetchrow_hashref|/fetchrow_hashref( $row )>.  It accepts a list of row numbers that 
will be collated into a set of headers used to build the hashref for each row.
The header rows are coallated in sequence with the first number taking precedence.  
The list is also used to set the lowest row of the headers in the table.  All rows 
at that level and higher will be considered out of the table and will return undef 
while setting the error instance.  If some of the columns do not have values then 
the instance will auto generate unique headers for each empty header column to fill 
out the header ref.

B<Accepts:> a list of row numbers (modified as needed by the attribute state of 
L<Spreadsheet::XLSX::Reader::LibXML/count_from_zero>)

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