package Spreadsheet::XLSX::Reader::LibXML::GetCell;
use version; our $VERSION = qv('v0.24.2');

use	Moose::Role;
requires qw(
	get_log_space				set_error					min_row
	max_row						min_col						max_col
	row_range					col_range
);
use Types::Standard qw(
	Bool 						HasMethods					Enum
	Int							is_Int						ArrayRef
	is_ArrayRef					HashRef						is_HashRef
);# Int
###LogSD	use Log::Shiras::Telephone;
use lib	'../../../../../lib',;
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

has custom_formats =>(
		isa		=> HashRef[ HasMethods[ 'assert_coerce', 'display_name' ] ],
		traits	=> ['Hash'],
		handles	=>{
			has_custom_format => 'exists',
			get_custom_format => 'get',
			set_custom_format => 'set',
		},
		writer	=> 'set_custom_formats',
		default => sub{ {} },
		clearer	=> '_clear_custom_formats',
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
						get_empty_return_type
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
						get_empty_return_type
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
	if( !defined $row ){
		my $last_col = $self->_get_reported_col;
		###LogSD	$phone->talk( level => 'info', message =>[
		###LogSD		"No row requested - Determining what it should be with the last column: $last_col", ] );
		$row = $self->_get_reported_row + !!$last_col;
		###LogSD	$phone->talk( level => 'info', message =>[
		###LogSD		"Now requesting Excel row: $row", ] );
		$row = $self->_get_used_position( $row );
		###LogSD	$phone->talk( level => 'info', message =>[
		###LogSD		"Which is used row: $row", ] );
	}
	
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

#########1 Private Attributes 3#########4#########5#########6#########7#########8#########9

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
		###LogSD		"Initial result is:", $result ] );
		$result->{cell_type} = 'Text';
		if( exists $result->{t} and $result->{t} eq 's' ){
			@$result{qw( cell_unformatted rich_text )} =
				@{$self->get_shared_string_position( $result->{v}->{raw_text} )}{qw( raw_text rich_text )};
			delete $result->{t};
			delete $result->{v};
			delete $result->{rich_text} if !$result->{rich_text};
		}else{
			$result->{cell_unformatted} = $result->{v}->{raw_text};
			$result->{cell_type} = 'Numeric' if $result->{cell_unformatted} and $result->{cell_unformatted} ne '';
			delete $result->{v};
		}
		if( !$result->{cell_unformatted} and $self->get_empty_return_type eq 'empty_string' ){
			###LogSD	$phone->talk( level => 'debug', message =>[ "(Re)setting undef to ''"] );
			$result->{cell_unformatted} = '';
		}
		###LogSD	$phone->talk( level => 'debug', message =>[
		###LogSD		"Cell raw text is:", $result->{cell_unformatted}] );
		if( $self->get_group_return_type eq 'unformatted' ){
			my $reformatted = $self->change_output_encoding( $result->{cell_unformatted} );
			return  $reformatted;
		}
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
		if( $custom_format ){
			###LogSD	$phone->talk( level => 'debug', message =>[
			###LogSD		'Cell custom_format is:', $custom_format ] );
			if( $self->get_group_return_type eq 'value' ){
				my $reformatted = $self->change_output_encoding( $result->{cell_unformatted} );
				return $custom_format->assert_coerce( $reformatted );
			}
		}
		# handle the formula
		if( exists $result->{f} ){
			$result->{cell_formula} = $result->{f}->{raw_text};
			delete $result->{f};
		}
		
		if( exists $result->{s} ){
			my $header = ($self->get_group_return_type eq 'value') ? 'numFmts' : undef;
			my $format;
			if( $self->_has_styles_file ){
				$format = $self->get_format_position( $result->{s}, $header );#############  Change here when allow a list rather than a string filter
				###LogSD	$phone->talk( level => 'trace', message =>[
				###LogSD		"format position is:", $format ] );
			}
			if( $self->get_group_return_type eq 'value' ){
				###LogSD	$phone->talk( level => 'trace', message =>[
				###LogSD		"Attempting to just return the cell value with:",
				###LogSD		((exists $format->{numFmts}) ? $format->{numFmts} : undef), ] );
				if( exists $format->{numFmts} ){
					return $format->{numFmts}->assert_coerce( $result->{cell_unformatted} );
				}else{
					return $result->{cell_unformatted};
				}
			}
			if( $self->_has_styles_file ){
				for my $header ( keys %$format_headers ){
					if( exists $format->{$header} ){
						###LogSD	$phone->talk( level => 'trace', message =>[
						###LogSD		"Transferring styles header -$header- to cell attribute: $format_headers->{$header}", ] );
						if( $header eq 'numFmts' ){
							$result->{cell_coercion} = $format->{$header};
							if(	$result->{cell_type} eq 'Numeric' and
								$format->{$header}->name =~ /date/i ){
								###LogSD	$phone->talk( level => 'trace', message =>[
								###LogSD		"Found a -Date- cell", ] );
								$result->{cell_type} = 'Date';
							}
						}else{
							$result->{$format_headers->{$header}} = $format->{$header};
						}
					}
				}
			}
			if( $custom_format ){
				###LogSD	$phone->talk( level => 'trace', message =>[
				###LogSD		"Using custom number formats", ] );
				$result->{cell_coercion} = $custom_format;
				$result->{cell_type} = 'Custom';
			}
			delete $result->{s},
		}
		###LogSD	$phone->talk( level => 'trace', message =>[
		###LogSD		"Checking return type: " . $self->get_group_return_type ] );
		if( $self->get_group_return_type eq 'value' ){
			###LogSD	$phone->talk( level => 'trace', message =>[
			###LogSD		"The caller only wants the coerced value (Not the whole cell)", ] );
			if( !$result->{cell_coercion} ){
				$result = $result->{cell_unformatted};
			}elsif( !defined $result->{cell_unformatted} ){
				$result = undef;
				$self->set_error( "The cell does not have a value" );
				;
			}else{
				eval '$result = $result->{cell_coercion}->assert_coerce( $result->{cell_unformatted} )';
				if( $@ ){
					$self->set_error( $@ );
				}
			}
			$result =~ s/\\//g if $result;
			###LogSD	$phone->talk( level => 'trace', message =>[
			###LogSD		"Returning: $result", ] );
			return $result;
		}
		$result->{cell_row} = $self->_get_used_position( $result->{row} );
		delete $result->{row};
		$result->{cell_col} = $self->_get_used_position( $result->{col} );
		delete $result->{col};
		$result->{error_inst} = $self->_get_error_inst;
		$result->{log_space} = $self->get_log_space . '::Cell';
		$result->{unformatted_converter} = sub{ 
			my	$string = $_[0];
			###LogSD	$phone->talk( level => 'debug', message =>[
			###LogSD		"Sending stuff to converter:", @_ ] );
			$self->change_output_encoding( $string ) 
		};
		###LogSD	$phone->talk( level => 'debug', message =>[
		###LogSD		"Final ref is:", $result ] );
		
		# build cell
		$return = Spreadsheet::XLSX::Reader::LibXML::Cell->new( %$result );
	}elsif( $result ){
		$return = ( $self->boundary_flag_setting ) ? $result : undef;
	}
	###LogSD		$phone->talk( level => 'debug', message =>[
	###LogSD			"Cell is:", $return ] );
	return $return;
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
    
=head1 DESCRIPTION

B<This documentation is written to explain ways to extend this package.  To use the data 
extraction of Excel workbooks, worksheets, and cells please review the documentation for  
L<Spreadsheet::XLSX::Reader::LibXML>,
L<Spreadsheet::XLSX::Reader::LibXML::Worksheet>, and 
L<Spreadsheet::XLSX::Reader::LibXML::Cell>>

This is the extracted L<Role|Moose::Manual::Roles> to be used as a top level worksheet 
interface.  This is the place where all the various details in each sub XML sheet are 
coallated into a set of data representing all the necessary information for a requested 
cell.  Since this is the center of data coallation all elements that may be customized 
should reside outside of this role.  This includes any specific elements that would be 
different between each of the sheet parser types and any element of Excel data presentation 
that may lend itself to customization.  For instance all the XML parser 
methods, (Reader, DOM, and possibly SAX) can plug in outside of this role.

This role is also where a layer of abstraction is maintained to manage user defined 
count-from-one or count-from-zero mode.  The layer of abstraction is use with the Moose 
L<around|Moose::Manual::MethodModifiers/AROUND modifiers> modifier.  The behaviour is 
managed with the workbook attribute L<counting_from_zero
|Spreadsheet::XLSX::Reader::LibXML/count_from_zero>.

=head2 requires

These are method(s) used by this Role but not provided by the role.  Any class consuming this 
role will not build without first providing these methods prior to loading this role.  
I<Since this is the center of data coallation the list is long>.

=head3 get_log_space

=over

B<Definition:> Used to return the log space used by the code protected by ###LogSD.  See
L<Log::Shiras|https://github.com/jandrew/Log-Shiras> for more information.

=back

=head3 set_error( $error_string )

=over

B<Definition:> Used to set errors that occur in code from this role.  See
L<Spreadsheet::XLSX::Reader::LibXML::Error> for the default implementation of this functionality.

=back

=head3 min_row

=over

B<Definition:> Used to get the minimum data row set for the worksheet.

=back

=head3 max_row

=over

B<Definition:> Used to get the maximum data row set for the worksheet.

=back

=head3 min_col

=over

B<Definition:> Used to get the minimum data column set for the worksheet.

=back

=head3 max_col

=over

B<Definition:> Used to get the maximum data column set for the worksheet.

=back

=head3 row_range

=over

B<Definition:> Used to return a list of the $minimum_row and $maximum_row

=back

=head3 col_range

=over

B<Definition:> Used to return a list of the $minimum_column and $maximum_column

=back

=head3 change_output_encoding

=over

B<Definition:> This should be using the L<localization
|Spreadsheet::XLSX::Reader::LibXML::FmtDefault/change_output_encoding> 
role.  The role should be reachable by delegation through a workbook 
parser attribute (trait) in the class consuming this role.

=back

=head3 get_group_return_type

=over

B<Definition:> This is where the L<group_return_type
|Spreadsheet::XLSX::Reader::LibXML/group_return_type> attribute is used 
as a consequence this role needs to understand the current setting.  It 
should be reachable by delegation through a workbook parser attribute 
(trait) in the class consuming this role.

=back

=head3 set_group_return_type

=over

B<Definition:> In order to speed up header reading and L<setting
|/set_headers( @header_row_list )> the full cell is not needed, just the 
value.  As a consequence this role changes the group_return_type attribute 
as needed for header retrieval.

=back

=head3 _get_next_value_cell

=over

B<Definition:> This should return the next cell data from the worksheet file 
that contains unique formatting or information.  The data is expected in a 
perl hash ref.  This method should collect data left to right and top to 
bottom.  I<The styles.xml, sharedStrings.xml, and calcChain.xml etc. sheet data 
are coallated into the cell later in this role>.  An 'EOF' string should be 
returned when the file has reached the end and then the method should wrap 
back to the beginning.

B<Example>

	{
		'r' => 'A6',				# The cell ID
		'cell_merge' => 'A6:B6',	# The merge range
		'row' => 6,					# Already converted to count by 1
		'col' => 1,					# Already converted to count by 1
		't' => 's'					# Cell data type (string)
		'v' => {					# Cell data (since this cell is string 
			'raw_text' => '15'		# 	data this actually points to position 
		}							# 	15 in the sharedStrings.xml file )
		's' => '11',				# Styles type (position 11 in the styles sheet)
	}

=back

=head3 _get_next_cell

=over

B<Definition:> Like L<_get_next_value_cell|/_get_next_value_cell> this method should 
return the next cell.  The difference is it should return undef for empty cells rather 
than skipping them.  This method should collect data left to right and top to bottom.  
I<The styles.xml, sharedStrings.xml, and calcChain.xml etc. sheet data are coallated 
into the cell later in this role>.  An 'EOF' string should be returned when the file 
has reached the end and then the method should wrap back to the beginning.

=back

=head3 _get_col_row

=over

B<Definition:> This method should provide a targeted way to return the worksheet file 
information on a cell.  It should only accept count-from-one column and row numbers 
and the column should be required before the row.  If the request is made for an out 
of row bounds position the method should provide an 'EOR' string.  An 'EOF' string 
should be returned when the file has reached the end and then the method should 
wrap back to the beginning.

=back

=head3 _get_error_inst

=over

B<Definition:> This method is used to access the shared 
L<Spreadsheet::XLSX::Reader::LibXML/error_inst>

=back

=head3 counting_from_zero

=over

B<Definition:> This role implements some of the L<count_from_zero
|Spreadsheet::XLSX::Reader::LibXML/count_from_zero> attribute behaviors so it needs to 
be able to read the current state.

=back

=head3 boundary_flag_setting

=over

B<Definition:> This role implements some of the L<file_boundary_flags
|Spreadsheet::XLSX::Reader::LibXML/file_boundary_flags> attribute behaviors so it needs to 
be able to read the current state.

=back

=head3 _has_shared_strings_file

=over

B<Definition:> This should indicate if the sharedStrings.xml file is available

=back

=head3 get_shared_string_position( $position )

=over

B<Definition:> This should return a hashref of data for the indicated $position

B<Example:>  from the example shared strings position 15

	{
		'rich_text' => [				# Rich text definition
			2,							# Position from 0 to start this element
			{							# Element definition
				'color' => {
					'rgb' => 'FFFF0000'
				},
				'sz' => '11',
				'b' => 1,
				'scheme' => 'minor',
				'rFont' => 'Calibri',
				'family' => '2'
			},
			6,							# Position from 0 to start this element
			{							# Element definition
				'color' => {
					'rgb' => 'FF0070C0'
				},
				'sz' => '20',
				'b' => 1,
				'scheme' => 'minor',
				'rFont' => 'Calibri',
				'family' => '2'
			}
			],
		'raw_text' => 'Hello World',	# The raw text the format is applied to
	}

=back

=head3 _has_styles_file

=over

B<Definition:> This should indicate if the styles.xml file is available

=back

=head3 get_format_position( $position )

=over

B<Definition:> This should return a hashref of data for the indicated $position 
in the styles.xml file.  This will include any general cell formatting as well as 
any references to the subroutines for number conversions either defined by Excel or
any custom (user defined) conversion subroutines

B<Example:>  from the example styles position 11

	{
		'fontId' => '0',
		'fonts' => {
			'color' => {
				'theme' => '1'
			},
			'sz' => '11',
			'name' => 'Calibri',
			'scheme' => 'minor',
			'family' => '2'
		},
		'numFmtId' => '0',
		'fillId' => '0',
		'xfId' => '0',
		'applyAlignment' => '1',
		'borders' => {
			'left' => 1,
			'right' => 1,
			'top' => 1,
			'diagonal' => 1,
			'bottom' => 1
		},
		'borderId' => '0',
		'alignment' => {
			'horizontal' => 'left'
		},
		'cellStyleXfs' => {
			'fillId' => '0',
			'fontId' => '0',
			'borderId' => '0',
			'numFmtId' => '0'
		},
		'fills' => {
			'patternFill' => {
				'patternType' => 'none'
			}
		},
		'numFmts' => bless( {						# This is the package build Type::Tiny object
			'name' => 'Excel_number_0',				#  to be use for coercion
			'coercion' => bless( { 
				~~ Type::Coercion instance here ~~
			 }, 'Type::Coercion' ),
		'display_name' => 'Excel_number_0',
		'uniq' => 94
		}, 'Type::Tiny' )
	}

=back

=head2 Primary Methods

These are the various methods provided by this role.

=head3 get_cell( $row, $column )

=over

B<Definition:> Used to return the cell or information from the cell at the 
specified $row and $column.  Both $row and $column are required.

B<Accepts:> the list ( $row, $column ) both required

B<Returns:> see L<group_return_type|Spreadsheet::XLSX::Reader::LibXML/group_return_type> 
for details on what is returned

=back

=head3 get_next_value

=over

B<Definition:> Reading left to right and top to bottom this will return the next cell with 
a value.  This actually includes cells with no value but some unique formatting such as 
cells that have been merged with other cells.

B<Accepts:> nothing

B<Returns:> see L<group_return_type|Spreadsheet::XLSX::Reader::LibXML/group_return_type> 
for details on what is returned

=back

=head3 fetchrow_arrayref( $row )

=over

B<Definition:> In an homage to L<DBI> I included this function to return an array ref of 
the cells or values in the requested $row.  If no row is requested this returns the 'next' 
row.  In the array ref any empty and non unique cell will show as 'undef'.

B<Accepts:> undef = next|$row = a row integer indicating the desired row

B<Returns:> an array ref of all possible column positions in that row with data filled in 
as appropriate.

=back

=head3 fetchrow_array( $row )

=over

B<Definition:> This function is just like L<fetchrow_arrayref|/fetchrow_arrayref( $row )> 
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

B<Accepts:> a list of row numbers

B<Returns:> an array ref of the built headers for review

=back

=head3 fetchrow_hashref( $row )

=over

B<Definition:> This function is used to return a hashref representing the data in the 
specified row.  If no $row value is passed it will return the 'next' row of data.  A call 
to this function without L<setting|/( @header_row_list )> the headers first 
will return undef and set the error instance.

B<Accepts:> a target $row number for return values or undef meaning 'next'

B<Returns:> a hash ref of the values for that row

=back

=head2 Attributes

Data passed to new when creating the L<Styles|Spreadsheet::XLSX::Reader::LibXML::Worksheet> 
instance.   For modification of these attributes see the listed 'attribute methods'.
For more information on attributes see L<Moose::Manual::Attributes>.  Most of these are 
not exposed to the top level of the workbook L<parser|Spreadsheet::XLSX::Reader::LibXML>.  
As a consequence these attribute methods which are available at the worksheet instance 
level are the best way to manipulate the attribute settings.

=head3 last_header_row

=over

B<Definition:> This is generally set by the method L<set_headers
|/set_headers( @header_row_list )> and is the largest row number of the @header_row_list 
even if the list is out of sequence.

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

B<Definition:> When the method L<fetchrow_array(ref)|/fetchrow_arrayref( $row )> 
or L<fetchrow_hashref|/fetchrow_hashref( $row )> are called it is possible to 
only collect a set of information between two defined columns.  This is the 
attribute that defines the start point.

B<Default:> undef

B<attribute methods> Methods provided to adjust this attribute
		
=over

B<set_min_header_col>

=over

B<Definition:> returns the value of the attribute

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

B<Definition:> sets min_header_col to undef

=back

=back

=back

=back

=head3 max_header_col

=over

B<Definition:> When the method L<fetchrow_array(ref)|/fetchrow_arrayref( $row )> 
or L<fetchrow_hashref|/fetchrow_hashref( $row )> are called it is possible to 
only collect a set of information between two defined columns.  This is the 
attribute that defines the end point.

B<Default:> undef

B<attribute methods> Methods provided to adjust this attribute
		
=over

B<set_max_header_col>

=over

B<Definition:> returns the value of the attribute

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

B<Definition:> sets min_header_col to undef

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
must have two methods 'assert_coerce' and 'display_name'.  The collater first 
checks the cellID as a key, then it checks for just the column letter(s) as a key, 
and finally it checks the row number as a key.

B<Default:> undef

B<attribute methods> Methods provided to adjust this attribute
		
=over

B<set_custom_formats( { $key => $conversion } )>

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

B<set_custom_format( $key => $conversion )>

=over

B<Definition:> set the custom format $conversion for the identified $key

=back

=back

=back

=back

=head1 SUPPORT

=over

L<github Spreadsheet::XLSX::Reader::LibXML/issues
|https://github.com/jandrew/Spreadsheet-XLSX-Reader-LibXML/issues>

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

This software is copyrighted (c) 2014 by Jed Lund

=head1 DEPENDENCIES

=over

L<Spreadsheet::XLSX::Reader::LibXML>

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