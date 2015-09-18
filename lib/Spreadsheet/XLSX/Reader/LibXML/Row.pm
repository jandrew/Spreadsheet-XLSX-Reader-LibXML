package Spreadsheet::XLSX::Reader::LibXML::Row;
use version; our $VERSION = qv('v0.38.16');
###LogSD	warn "You uncovered internal logging statements for Spreadsheet::XLSX::Reader::LibXML::Row-$VERSION";

$| = 1;
use 5.010;
use Moose;
use MooseX::StrictConstructor;
use MooseX::HasDefaults::RO;
use Carp qw( confess );
use Clone qw( clone );
use Types::Standard qw(
		ArrayRef			Int						Bool
		HashRef
    );
use lib	'../../../../../lib';
###LogSD	use Log::Shiras::Telephone;
###LogSD	use Log::Shiras::UnhideDebug;
###LogSD	with 'Log::Shiras::LogSpace';

#########1 Public Attributes  3#########4#########5#########6#########7#########8#########9

#~ has	error_inst =>(
		#~ isa			=> InstanceOf[ 'Spreadsheet::XLSX::Reader::LibXML::Error' ],
		#~ clearer		=> '_clear_error_inst',
		#~ reader		=> '_get_error_inst',
		#~ required	=> 1,
		#~ handles =>[ qw(
			#~ error set_error clear_error set_warnings if_warn
		#~ ) ],
	#~ );

has row_number =>(
		isa			=> Int,
		reader		=> 'get_row_number',
		required	=> 1,
	);

has row_span =>(
		isa			=> ArrayRef[ Int ],
		traits		=> ['Array'],
		writer		=> 'set_row_span',
		predicate	=> 'has_row_span',
		required 	=> 1,
		handles 	=>{
			get_row_start => [ 'get' => 0 ],
			get_row_end   => [ 'get' => 1 ],
		},
	);

has row_last_value_column =>(
		isa		=> Int,
		reader	=> 'get_last_value_column',
	);

has row_formats =>(# Add to the cell values?
		isa		=> HashRef,
		traits	=> ['Hash'],
		writer	=> 'set_row_formts',
		handles =>{
			get_row_format => 'get',
		},
	);

has column_to_cell_translations =>(
		isa			=> ArrayRef,
		traits		=>[ 'Array' ],
		required	=> 1,
		handles	=>{
			get_position_for_column => 'get',
		},
	);

has row_value_cells =>(
		isa			=> ArrayRef,
		traits		=>[ 'Array' ],
		reader		=> 'get_row_value_cells',
		required	=> 1,
		handles	=>{
			get_cell_position => 'get',
			total_cell_positions => 'count',
		},
	);

#########1 Public Methods     3#########4#########5#########6#########7#########8#########9

sub get_the_column{
	my ( $self, $desired_column ) = @_;
	confess "Desired column required" if !defined $desired_column;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space => 
	###LogSD					($self->get_all_space . '::get_the_column' ), );
	###LogSD		$phone->talk( level => 'debug', message =>[  
	###LogSD			 "Getting the cell value at column: $desired_column", ] );
	my	$max_column = $self->get_row_end;
	if( $desired_column > $max_column ){
		###LogSD	$phone->talk( level => 'debug', message =>[  
		###LogSD			"Requested column -$desired_column- is past the end of the row", ] );
		return 'EOR';
	}
	my	$value_position = $self->get_position_for_column( $desired_column );
	if( !defined $value_position ){
		###LogSD	$phone->talk( level => 'debug', message =>[  
		###LogSD			"No cell value stored for column: $desired_column", ] );
		return undef;
	}
	my $return_cell = $self->get_cell_position( $value_position );
	###LogSD	$phone->talk( level => 'debug', message =>[  
	###LogSD		"Returning the cell:", $return_cell, ] );
	#~ $self->_set_reported_column( $desired_column );
	$self->_set_reported_position( $value_position );
	return clone( $return_cell );
}

sub get_the_next_value_position{
	my ( $self, ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space => 
	###LogSD					($self->get_all_space . '::get_the_next_value_column' ), );
	###LogSD		$phone->talk( level => 'debug', message =>[  
	###LogSD			 "Returning the next value position in this row as available", ] );
	my $next_position = defined $self->_get_reported_position ? ($self->_get_reported_position + 1) : 0;
	if( $next_position == $self->total_cell_positions ){# Counting from zero vs counting from 1
		###LogSD	$phone->talk( level => 'debug', message =>[  
		###LogSD		"Already reported the last value position" ] );
		return 'EOR';
	}
	my $return_cell = $self->get_cell_position( $next_position );
	#~ $self->_set_reported_column( $return_cell->{cell_col} );
	$self->_set_reported_position( $next_position );
	###LogSD	$phone->talk( level => 'debug', message =>[  
	###LogSD		"Returning the cell:", $return_cell, ] );
	return clone( $return_cell );
}

sub get_row_all{
	my ( $self, ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space => 
	###LogSD					($self->get_all_space . '::get_row_all' ), );
	###LogSD		$phone->talk( level => 'debug', message =>[  
	###LogSD			 "Getting an array ref of all the cells in the row by column position", ] );
	
	my $return_ref;
	for my $cell ( @{$self->get_row_value_cells} ){
		$return_ref->[$cell->{cell_col} - 1] = clone $cell;
	}
	###LogSD	$phone->talk( level => 'debug', message =>[  
	###LogSD		"Returning the row ref:", $return_ref, ] );
	return $return_ref;
}

#########1 Private Attributes 3#########4#########5#########6#########7#########8#########9

has _reported_position =>(
		isa			=> Int,
		reader		=> '_get_reported_position',
		writer		=> '_set_reported_position',
	);

#########1 Private Methods    3#########4#########5#########6#########7#########8#########9

###LogSD	sub BUILD {
###LogSD	    my $self = shift;
###LogSD			$self->set_class_space( 'Row' );
###LogSD	}


#########1 Phinish            3#########4#########5#########6#########7#########8#########9

no Moose;
__PACKAGE__->meta->make_immutable;
	
1;

#########1 Documentation      3#########4#########5#########6#########7#########8#########9
__END__

=head1 NAME

Spreadsheet::XLSX::Reader::LibXML::Row - XLSX Row data class

=head1 SYNOPSIS

This is really an internal class that is not intended to be used in a stand-alone fashion;
    
=head1 DESCRIPTION

Documentation not written yet!

=cut

#########1#########2 main pod documentation end  5#########6#########7#########8#########9