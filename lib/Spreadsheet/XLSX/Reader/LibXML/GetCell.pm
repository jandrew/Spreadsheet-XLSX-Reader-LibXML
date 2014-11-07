package Spreadsheet::XLSX::Reader::LibXML::GetCell;
use version; our $VERSION = qv('v0.5_1');

use	Moose::Role;
requires qw(
	min_row						max_row						min_col
	max_col						row_range					col_range
	_get_next_value_cell		_get_next_cell				_get_col_row
	_get_row_all				_get_error_inst				counting_from_zero
	boundary_flag_setting		change_boundary_flag		_has_shared_strings_file
	get_shared_string_position	_has_styles_file			get_format_position
	get_group_return_type		set_group_return_type
);
use Types::Standard qw(
	Bool 						HasMethods					Enum
	Int							is_Int						ArrayRef
	is_ArrayRef					HashRef						is_HashRef
);# Int
###LogSD	use Log::Shiras::Telephone;
###LogSD	use Log::Shiras::UnhideDebug;
use lib	'../../../../../lib',;
use	Spreadsheet::XLSX::Reader::LibXML::Cell v0.5;

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
		clearer		=> '_clear_min_header_col',
		predicate	=> 'has_min_header_col'
	);

has max_header_col =>(
		isa			=> Int,
		reader		=> 'get_max_header_col',
		writer		=> 'set_max_header_col',
		clearer		=> '_clear_max_header_col',
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
				push @$return, $self->_build_out_the_cell( $cell, $self->get_group_return_type );
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
			$result->{cell_type} = 'Numeric' if $result->{cell_unformatted};
			delete $result->{v};
		}
		###LogSD	$phone->talk( level => 'debug', message =>[
		###LogSD		"Cell raw text is:", $result->{cell_unformatted}] );
		return $result->{cell_unformatted} if $self->get_group_return_type eq 'unformatted';
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
				return $custom_format->assert_coerce( $result->{cell_unformatted} );
			}
		}
		if( exists $result->{f} ){
			$result->{cell_formula} = $result->{f}->{raw_text};
			delete $result->{f};
		}
		
		if( exists $result->{s} ){
			my $header = ($self->get_group_return_type eq 'value') ? 'numFmts' : undef;
			my $format = $self->get_format_position( $result->{s}, $header );#############  Change here when allow a list rather than a string filter
			###LogSD	$phone->talk( level => 'trace', message =>[
			###LogSD		"format position is:", $format ] );
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
		###LogSD		$phone->talk( level => 'debug', message =>[
		###LogSD			"Final ref is:", $result ] );
		
		# build cell
		$return = Spreadsheet::XLSX::Reader::LibXML::Cell->new( %$result );
	}elsif( $result ){
		$return = ( $self->boundary_flag_setting ) ? $result : undef;
	}
	###LogSD		$phone->talk( level => 'debug', message =>[
	###LogSD			"Cell is:", $return ] );
	return $return;
}

sub get_col_row{ my $self = shift; $self->_get_col_row( @_ ); };

sub get_row_all{ my $self = shift; $self->_get_row_all( @_ ); };

#########1 Phinish            3#########4#########5#########6#########7#########8#########9

no Moose::Role;
1;

#########1 Documentation      3#########4#########5#########6#########7#########8#########9
__END__

=head1 NAME

Spreadsheet::XLSX::Reader::LibXML::CellToRowColumn - Translate Excel cell IDs to row column
    
=head1 DESCRIPTION

This is a fairly simple implementation of a regex and math to find the column and row
position in excel from an 'A1' style Excel cell ID.  It is important to note that column 
letters do not equal digits in a modern 26 position numeral system since the excel 
implementation is effectivly zeroless.

The default of this module, however, is to count from zero.  Meaning that cell A1 is 
equal to (0, 0).  See the L<Attributes|/Attributes> and L<Methods|/Methods> section for 
ways to change this behaviour.

=head1 SYNOPSIS
	
	#!perl
	package MyPackage;
	use Moose;
	with 'Spreadsheet::XLSX::Reader::LibXML::CellToColumnRow';

	sub set_error{}
	sub get_log_space{}
		
	sub my_method{
		my ( $self, $cell ) = @_;
		my ($column, $row ) = $self->parse_column_row( $cell );
		print $self->error if( !defined $column or !defined $row );
		return ($column, $row );
	}

	package main;

	my $parser = MyPackage->new( count_from_zero => 0 );
	print '(' . join( ', ', $parser->my_method( 'B2' ) ) . ")'\n";
	
	###########################
	# SYNOPSIS Screen Output
	# 01: (2, 2)
	###########################
	
=head2 Attributes

Attiributes are ways to change the instances behaviour and can be set as arguments 
to -E<gt>new

=head3 count_from_zero

=over

B<Definition:> A boolean attribute that determines if the numerical output of 
L<parse_column_row|/parse_column_row( $excel_row_id )> provides a response counting from 
Zero or One. True = count from Zero.

B<Accepts:> $bool = (1|0)

=back
	
=head2 Methods

Methods are object methods (not functional methods)

=head3 parse_column_row( $excel_row_id, $count_from_one )

=over

B<Definition:> This is the way to turn an alpha numeric Excel cell ID into row and column 
integers.  If count_from_zero = 1 but you want (column, row) pairs returned counting from 
1 then set $count_from_one = 1.  Or leave it blank to have the pair returned in the format 
defined by L<count_from_zero|/count_from_zero>

B<Accepts:> $excel_row_id, $count_from_one

B<Returns:> ( $column_number, $row_number ) - integers

=back

=head3 build_cell_label( $column, $row, $count_from_one )

=over

B<Definition:> This is the way to turn a (column, row) pair into an excel ID.  If 
$count_from_one is set then the ($column, $row pair will be treated at counting from one 
independant of how L<count_from_zero|/count_from_zero> is set.
integers

B<Accepts:> $column, $row, $count_from_one (in that order and position)

B<Returns:> ( $excel_cell_id ) - integers

=back

=head3 counting_from_zero( $bool )

=over

B<Definition:> This turns on (or off) counting from zero where the alternative is to 
count from 1.

B<Accepts:> $bool = (1|0)

B<Returns:> nothing

=back

=head3 get_excel_position( $int )

=over

B<Definition:> If you wish to use this sheet agnostically of the L<count_from_zero|/count_from_zero> 
setting then you can use this method to translate integers to a count-from-one number.  No action is 
taken if the attribute is set to 0.

B<Accepts:> a $count_from_one or a $count_from_zero int

B<Returns:> a $count_from_one int

=back

=head3 get_used_position( $int )

=over

B<Definition:> If you wish to use this sheet agnostically of the L<count_from_zero|/count_from_zero> 
setting then you can use this method to translate integers from a count-from-one number to whatever 
scheme is in force from the attribute.  No action is taken if the attribute is set to 0.

B<Accepts:> a $count_from_one int

B<Returns:> a $count_from_one or a $count_from_zero int

=back

=head1 SUPPORT

=over

L<github Spreadsheet-XLSX-Reader-LibXML/issues|https://github.com/jandrew/Spreadsheet-XLSX-Reader-LibXML/issues>

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

This software is copyrighted (c) 2014 by Jed Lund

=head1 DEPENDENCIES

=over

L<version>

L<Moose::Role>

L<Types::Standard>

requires

	name_space
	set_error

=back

=head1 SEE ALSO

=over

L<Spreadsheet::XLSX>

L<Log::Shiras|https://github.com/jandrew/Log-Shiras>

=back

=cut

#########1#########2 main pod documentation end  5#########6#########7#########8#########9