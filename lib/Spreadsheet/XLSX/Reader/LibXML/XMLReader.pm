package Spreadsheet::XLSX::Reader::LibXML::XMLReader;
use version; our $VERSION = qv('v0.18.2');

use 5.010;
use Moose;
use MooseX::StrictConstructor;
use MooseX::HasDefaults::RO;
use Types::Standard qw(
		Int				Str				HasMethods
		FileHandle
    );
use XML::LibXML::Reader;
use lib	'../../../../../lib',;
with 'Spreadsheet::XLSX::Reader::LibXML::LogSpace';
###LogSD	use Log::Shiras::Telephone;
###LogSD	use Log::Shiras::UnhideDebug;
use Spreadsheet::XLSX::Reader::LibXML::Types qw(
		XMLFile
	);

#########1 Public Attributes  3#########4#########5#########6#########7#########8#########9

has file_name =>(
		isa			=> XMLFile,
		reader		=> 'get_file_name',
		trigger		=> \&_set_file_name,
		required	=> 1,
	);

has	error_inst =>(
		isa			=> 	HasMethods[qw(
							error set_error clear_error set_warnings if_warn
						) ],
		clearer		=> '_clear_error_inst',
		reader		=> '_get_error_inst',
		required	=> 1,
		handles =>[ qw(
			error set_error clear_error set_warnings if_warn
		) ],
	);

#########1 Public Methods     3#########4#########5#########6#########7#########8#########9


sub start_the_file_over{
	my( $self, ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD					name_space => $self->get_log_space . '::start_the_file_over', );
	###LogSD		$phone->talk( level => 'debug', message =>[ "Resetting the XML file" ] );
	$self->_go_to_the_end;
	$self->_close_the_sheet;
	$self->_clear_xml_parser;
	$self->_clear_location;
	my $fh = $self->_get_file_handle;
	seek( $fh, 0, 0 );
	$self->_set_xml_parser( XML::LibXML::Reader->new( IO => $fh ) );
}

#########1 Private Attributes 3#########4#########5#########6#########7#########8#########9

has _xml_reader =>(
	isa			=> 'XML::LibXML::Reader',
	reader		=> '_get_xml_parser',
	writer		=> '_set_xml_parser',
	predicate	=> '_has_xml_parser',
	clearer		=> '_clear_xml_parser',
	handles	=>{
		copy_current_node	=> 'copyCurrentNode',
		byte_consumed		=> 'byteConsumed',
		start_reading		=> 'read',
		next_element		=> 'nextElement',
		next_sibling		=> 'nextSibling',
		get_attribute		=> 'getAttribute',
		read_state			=> 'readState',
		node_name			=> 'name',
		node_value			=> 'value',
		has_value			=> 'hasValue',
		inner_xml			=> 'readInnerXml',
		node_depth			=> 'depth',
		is_empty			=> 'isEmptyElement',
		inner_xml			=> 'readInnerXml',
		has_attributes		=> 'hasAttributes',
		get_attribute_count	=> 'attributeCount',
		read_attribute		=> 'readAttributeValue',
		constant_value		=> 'ConstValue',
		move_to_first_att	=> 'moveToFirstAttribute',
		move_to_next_att	=> 'moveToNextAttribute',
		_encoding			=> 'encoding',
		_go_to_the_end		=> 'finish',
		_close_the_sheet	=> 'close',
		next_sibling_element	=> 'nextSiblingElement',
	}
);

has _file_handle =>(
		isa			=> FileHandle,
		reader		=> '_get_file_handle',
		writer		=> '_set_file_handle',
		predicate	=> '_has_file_handle',
		clearer		=> '_clear_file_handle',
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

#########1 Private Methods    3#########4#########5#########6#########7#########8#########9

sub _set_file_name{
	my( $self, $new_file, $old_file, $mapped ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD					name_space 	=> $self->get_log_space .  '::_set_file_name', );
	###LogSD	no warnings 'uninitialized';
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"(Re)setting the file to: $new_file",
	###LogSD			"From the file: $old_file",
	###LogSD			"With mapped setting: $mapped", ] );
	###LogSD	use warnings 'uninitialized';
	
	if( $self->_has_xml_parser ){
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"Pre-existing reader in place - clearing it" ] );
		$self->_go_to_the_end;
		$self->_close_the_sheet;
		$self->_clear_xml_parser;
		$self->_clear_location;
		$self->_close_file_handle if $self->_has_file_handle;
	}
	
	# Set the reader file
	open my $fh, '<', $new_file;
	binmode( $fh );
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"File handle open;", $fh ] );
	my	$reader		= XML::LibXML::Reader->new( IO => $fh );#'XMLFILElocation => $new_file )', );#recover => 2, 
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"XML parser open;", $reader,
	###LogSD		'Read state: ' . $reader->readState ] );
	if( !$reader ){
		$self->_clear_xml_parser;
		return undef;
	}else{
		###LogSD	$phone->talk( level => 'debug', message =>[ 'Success - Loading file handle: ' . $fh ], );
		$self->_set_file_handle( $fh );
		###LogSD	$phone->talk( level => 'debug', message =>[ 'Loading XML reader: ' . $reader ], );
		$self->_set_xml_parser( $reader );
		if( $self->byte_consumed == 0 ){
			###LogSD	$phone->talk( level => 'debug', message =>[ 'Starting the read' ], );
			$self->start_reading;
		}
		###LogSD	$phone->talk( level => 'debug', message =>[
		###LogSD		"Good reader built", "Byte position: " . $self->byte_consumed ], );#$reader->byteConsumed
		return 1 if $mapped;
	}
	
	# Get file encoding
	my	$encoding	= $self->_encoding;
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"Encoding of file is: $encoding" ], );
	$self->_set_encoding( $encoding );
	
	# Set the file unique bits
	if( $self->can( '_load_unique_bits' ) ){
		###LogSD	$phone->talk( level => 'debug', message => [ "Loading unique bits" ], );
		$self->_load_unique_bits;
		###LogSD	$phone->talk( level => 'debug', message => [ "Finished loading unique bits" ], );
	}
	return 1;#$reader;
}

sub DEMOLISH{
	my ( $self ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD					name_space 	=> $self->get_log_space .  '::XMLReader::DEMOLISH', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"clearing the reader for file_name:" . $self->get_file_name, ] );
	$self->_clear_xml_parser,
	$self->_clear_file_handle,
}

#########1 Phinish            3#########4#########5#########6#########7#########8#########9

no Moose;
	
1;

#########1 Documentation      3#########4#########5#########6#########7#########8#########9
__END__

=head1 NAME

Spreadsheet::XLSX::Reader::LibXML::XMLReader - LibXML::Reader base class for xlsx sheets

=head1 SYNOPSIS

	package MyPackage;
	use MooseX::StrictConstructor;
	use MooseX::HasDefaults::RO;
	extends	'Spreadsheet::XLSX::Reader::LibXML::XMLReader';
    
=head1 DESCRIPTION

POD not written yet!

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

#########1#########2 main pod documentation end   5#########6#########7#########8#########9