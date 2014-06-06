package Spreadsheet::XLSX::Reader::XMLReader;
use version; our $VERSION = version->declare("v0.1_1");

use 5.010;
use Moose;
use MooseX::StrictConstructor;
use MooseX::HasDefaults::RO;
use Types::Standard qw(
		Int
		Str
		InstanceOf
    );
use XML::LibXML;
my	$chunk_parser = XML::LibXML->new;
use XML::LibXML::Reader;
use lib	'../../../../lib',;
with 'Spreadsheet::XLSX::Reader::LogSpace';
###LogSD	use Log::Shiras::Telephone;
use Spreadsheet::XLSX::Reader::Types v0.1 qw(
		XMLFile
	);

#########1 Public Attributes  3#########4#########5#########6#########7#########8#########9

has file_name =>(
		isa			=> XMLFile,
		reader		=> 'get_file_name',
		trigger		=> \&_set_file_name,
		required	=> 1,
	);
	
has	epoch_year =>(
		isa		=> Int,
		reader	=> 'get_epoch_year',
		writer	=> 'set_epoch_year',
	);

has	error_inst =>(
		isa			=> InstanceOf[ 'Spreadsheet::XLSX::Reader::Error' ],
		clearer		=> '_clear_error_inst',
		reader		=> '_get_error_inst',
		required	=> 1,
		handles =>[ qw(
			error set_error clear_error set_warnings if_warn
		) ],
	);

#########1 Public Methods     3#########4#########5#########6#########7#########8#########9


sub get_position{
	my( $self, $position ) = @_;
	$position //= ( $self->has_position ) ? ( 1 + $self->where_am_i ) : 0;
	$self->_set_requested_position( $position );
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD					name_space => $self->get_log_space . '::get_position', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"Does have current position: " . $self->has_position,
	###LogSD			"Reached get_position for: $position",
	###LogSD			"Parsing file: " . $self->get_file_name ] );
	
	#checking if the reqested position is too far
	my	$result = inner();
	$self->_clear_requested_position;
	return undef if $result;
	
	# force the requested position to be <= the current position
	my	$reader;
	if( $self->has_position and $self->where_am_i <= $position ){
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"No reset needed - the requested postion -$position" .
		###LogSD		"- is on or after the current position: " . $self->where_am_i, ] );
	}elsif( !$self->has_position or $self->where_am_i > $position ){
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"Clearing the old reader and resetting the file: " . $self->get_file_name ] );
		$reader = $self->_set_file_name( $self->get_file_name, undef, 1 );
		#Set the reader to position 0
		$reader->nextElement( $self->get_core_element );
		$self->_i_am_here( 0 );
	}
	return undef if !$self->_has_xml_parser;
	
	#Advance to the requested position
		$reader //= $self->_get_xml_parser;
	if( $self->where_am_i == $position ){
		###LogSD	$phone->talk( level => 'debug', message => [ 
		###LogSD		"Already at position: $position" ] );
	}else{
		for my $x ( (1 + $self->where_am_i) .. $position ){
			###LogSD	$phone->talk( level => 'debug', message => [ 
			###LogSD		"Advancing to position: $x" ] );
			$reader->nextElement( $self->get_core_element );
			$self->_i_am_here( 1 + $self->where_am_i );
		}
	}
	my	$node = $reader->copyCurrentNode( 1 );
	###LogSD	$phone->talk( level => 'debug', message => [ 
	###LogSD		"Returning value: ", $node->toString ] );
	return $node;
}

#########1 Private Attributes 3#########4#########5#########6#########7#########8#########9

has _xml_reader =>(
	isa			=> 'XML::LibXML::Reader',
	reader		=> '_get_xml_parser',
	writer		=> '_set_xml_parser',
	predicate	=> '_has_xml_parser',
	clearer		=> '_clear_xml_parser',
);

has _file_encoding =>(
		isa		=> Str,
		reader	=> 'encoding',
		writer	=> '_set_encoding',
	);

has _position_index =>(
		isa			=> Int,
		reader		=> 'where_am_i',
		writer		=> '_i_am_here',
		clearer		=> '_clear_location',
		predicate	=> 'has_position',
	);
	
has _core_element =>(
		isa			=> Str,
		required	=> 1,
		reader		=> 'get_core_element',
	);

has _requested_position =>(
		isa			=> Int,
		reader		=> '_get_requested_position',
		writer		=> '_set_requested_position',
		clearer		=> '_clear_requested_position',
	);

#########1 Private Methods    3#########4#########5#########6#########7#########8#########9

sub _set_file_name{
	my( $self, $new_file, $old_file, $mapped ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD					name_space 	=> $self->get_log_space .  '::_set_file_name', );
	###LogSD	no warnings 'uninitialized';
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"(Re)setting the file to: $new_file",
	###LogSD			"From the file: $old_file",
	###LogSD			"With mapped setting: $mapped" ] );
	###LogSD	use warnings 'uninitialized';
	
	if( $self->_has_xml_parser ){
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"Pre-existing reader in place - clearing it" ] );
		my	$parser = $self->_get_xml_parser;
			$parser->finish();
			$parser->close();
		$self->_clear_xml_parser;
		$self->_clear_location;
	}
	
	# Set the reader file
	my	$reader		= XML::LibXML::Reader->new( location => $self->get_file_name );
	return undef if !$reader;
	###LogSD	$phone->talk( level => 'debug', message => [ "Good reader built" ], );
	
	
	#Set the file unique bits
	my	$reload = 0;
	if( !$mapped ){
		# Get file encoding
			$reader->read();
		my	$encoding	= $reader->encoding();
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"Encoding of file is: $encoding" ], );
		$self->_set_encoding( $encoding );
		if( $self->can( '_load_unique_bits' ) ){
			###LogSD	$phone->talk( level => 'debug', message => [ "Loading unique bits" ], );
			$reload = $self->_load_unique_bits( $reader, $encoding );
		}
	}
	if( !$reload ){
		###LogSD	$phone->talk( level => 'debug', message => ["All ready" ], );
	}elsif( $reload == 1 ){
		###LogSD	$phone->talk( level => 'debug', message => ["Reloading the reader ..." ], );
		$reader		= XML::LibXML::Reader->new( location => $self->get_file_name );
	}else{
		return undef;
	}
	
	# Set attributes
	$self->_set_xml_parser( $reader );
	###LogSD	$phone->talk( level => 'debug', message => [ "Parser ready for use ..." ] );
	
	return $reader;
}

sub DEMOLISH{
	my ( $self ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD					name_space 	=> $self->get_log_space .  '::XMLReader::DEMOLISH', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"clearing the reader for file_name:" . $self->get_file_name,
	###LogSD			"clearing the error instance",									] );
	$self->_clear_xml_parser,
	$self->_clear_error_inst,
}

#########1 Phinish            3#########4#########5#########6#########7#########8#########9

no Moose;
#~ __PACKAGE__->meta->make_immutable;
	
1;

#########1 Documentation      3#########4#########5#########6#########7#########8#########9
__END__

=head1 NAME

Spreadsheet::XLSX::Reader::XMLReader - Spreadsheet::XLSX LibXML Reader base class

=head1 SYNOPSIS

	package MyPackage;
	use MooseX::StrictConstructor;
	use MooseX::HasDefaults::RO;
	extends	'Spreadsheet::XLSX::Reader::XMLReader';
	
	has +_core_element =>(
			default => 'c',#<--Set the iterated element here
		);
    
=head1 DESCRIPTION

L<XML::LibXML> supports TIMTOWDI by providing multiple ways to parse a file.  This package 
is built to support general pull parsing using L<XML::LibXML::Reader>.  All specific 
readers are built on this.  If you want to use this to write your own reader is to provide 
the default for 'core_element' so that when you call L<get_position> you will know what 
xml element is being pulled.

=head2 AUGMENT

If the section on L<inner() and augment|Moose::Manual::MethodModifiers> in the method modifiers 
of the Moose manual dont make any sense then skip this section.  Otherwise this is a list of 
inner calls in L</METHOD>s where subclassing can leverage this feature.

=head3 get_position

=over

B<Definition:> The default method for get_position is to iterate through each element until the 
position is reached.  If it finds an empty position at or before the requested position it will 
return undef.  However, if a position is requested substantially farther down the stack reaching 
an end of file could take measureable time.  The inner call in this function will allow 
the subclass to tell if the requested position (even in a get next situation) is past the eof 
prior to the full iteration.  For some reason augment doesn't always read the passed value so you 
have to call _get_requested_position to reliably know what the currently requested position is.  
If the inner() call returns positive 'get_position' will assume that the position is past the end 
of the file and return undef for get_position.  An example of the 
L<augment|https://metacpan.org/pod/Moose::Manual::MethodModifiers#INNER-AND-AUGMENT> implementation 
is:

	augment 'get_position' => sub{
		my( $self, ) = @_;
	my	$position = $self->_get_requested_position;
		#checking if the reqested position is too far
		if( $position > $self->total_elements - 1 ){
			return 1;# EOF reached
		}else{
			return undef;# Not at the end yet
		}
	};
		
=back

=head3 _load_unique_bits

=over

B<Definition:> _load_unique_bits isn't a function with and inner() call in it.  I could never 
get Moose to call inner() in a trigger function and I'm sure there is good theory behind not doinng it.  
However, when the file is loaded to the parser there may be elements of the file that are not 
found in the _core_element(s) and should be loaded in the subclass for reference.  When this class
loads the file and checks if the subclass has a _load_unique_bits function.  The subclass can then 
load any relevent data to the instance that is unique to the subclass.  The subclass function 
_load_unique_bits is expected to either return undef|0 or '1'.  Returning '1' will force the file to 
reload.  This is useful when the _load_unique_bits function in the sublclass needs to iterate 
through the _core_elements section to find non-core data.  For the reader to then not be lost it 
needs to reset at the beginning of the file.  The relevant section in this class looks like this:

	my	$reload = 0;
	if( !$mapped and $self->can( '_load_unique_bits' ) ){
		###LogSD Loading unique bits
		$reload = $self->_load_unique_bits( $reader, $encoding );
	}
	if( !$reload ){
		###LogSD All ready
	}elsif( $reload == 1 ){
		$reader		= XML::LibXML::Reader->new( location => $self->get_file_name );
	}else{
		return undef;
	}
		
=back

=head2 ATTRIBUTES

Data passed to new when creating an instance (parser).  For modification of 
these attributes see the listed L</METHODS> of the instance.

=head3 file_name

=over

B<Definition:> This attribute stores the string used to access the file (the file name).  
It will be L<tested|Spreadsheet::XLSX::Reader::Types> as an XMLFile type.

B<Default> none

B<Range> any readable xml file (null not allowed)
		
=back

=head2 METHODS

These include methods to adjust attributes as well as providing methods to 
implement the functionality of the module.

=head3 get_position( $int )

=over

B<Definition:> This calls a routine that searches for the identified position of the 
defined _core_element in the xml file.  If no value is passed it will test for the 
current recorded position and pull the next one.

B<Accepts:> an integer representing the position in the array of elements indicated 
by the attribute _core_element.  It will calculate the next position if no value is 
passed.

B<Returns:> An L<XML::LibXML::Element> instance with the element data from the file 
contained in it.  If the requested position (or the next position) is passed the end 
of the file then this returns undef.

=back

=head3 error( $error_string )

=over

B<Definition:> This method is handled from the workbook link generally built by 
L<Spreadsheet::XLSX::Reader>.  See L<Spreadsheet::XLSX::Reader::Error> for details of 
the error_string attribute associated with this method.

B<Accepts:> a message string

B<Returns:> nothing

=back

=head3 get_epoch_year

=over

B<Definition:> This method is handled from the workbook link generally built by 
L<Spreadsheet::XLSX::Reader>.  See L<DateTimeX::Format::Excel> for why excel sheets 
can have different epoch years for dates.  This method assumes that the workbook 
already knows this information.

B<Accepts:> nothing

B<Returns:> a four digit integer representing the epoch year of the excel datas.  
Usually 1900 or 1904.

=back

=head3 encoding

=over

B<Definition:> This is the encoding of the file as recorded in the xml attribute.  
In general L<XML::LibXML> should be converting the data into unicode for perl..

B<Accepts:> nothing

B<Returns:> the value of the encoding attribute in the xml file

=back

=head3 where_am_i

=over

B<Definition:> The module tracks the last recorded _core_element position returned.  
This is the way to read that value.

B<Accepts:> nothing

B<Returns:> An integer counting from 0 of the last _core_element returned

=back

=head3 has_position

=over

B<Definition:> Either before the first position is returned or after the end of 
the _core_element list is reached the last recorded position will be undef.  
This is a way to test for that state.

B<Accepts:> nothing

B<Returns:> a boolean value indicating if there is a current _core_element 
position.

=back

=head3 get_core_element

=over

B<Definition:> If you want to know what the currently identified _core_element 
attribute is.

B<Accepts:> nothing

B<Returns:> the string stored in the _core_element attribute.

=back

=head1 SUPPORT

=over

L<github Spreadsheet::XLSX::Reader/issues|https://github.com/jandrew/Spreadsheet-XLSX-Reader/issues>

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

B<5.010> - (L<perl>)

L<version>

L<Moose>

L<MooseX::StrictConstructor>

L<MooseX::HasDefaults::RO>

L<Types::Standard>

L<XML::LibXML>

L<XML::LibXML::Reader>

L<Spreadsheet::XLSX::Reader::LogSpace>

L<Spreadsheet::XLSX::Reader::Types>

=back

=head1 SEE ALSO

=over

L<Spreadsheet::XLSX>

L<Spreadsheet::XLSX::Reader::TempFilter>

L<Log::Shiras|https://github.com/jandrew/Log-Shiras>

=back

=cut

#########1#########2 main pod documentation end   5#########6#########7#########8#########9