package Spreadsheet::XLSX::Reader::LibXML::FmtDefault;
use version; our $VERSION = qv('v0.10.2');

use	5.010;
use	Moose::Role;
requires qw(
	get_log_space
);

use Types::Standard qw( InstanceOf ArrayRef Str );
use lib	'../../../../../lib',;
###LogSD	use Log::Shiras::Telephone;

#########1 Dispatch Tables    3#########4#########5#########6#########7#########8#########9



#########1 Public Attributes  3#########4#########5#########6#########7#########8#########9

has excel_region =>(
		isa		=> Str,
		default	=> 'en',
		reader	=> 'get_excel_region',
	);
	
has	target_encoding =>(
		isa			=> Str,
		reader		=> 'get_target_encoding',
		writer		=> 'set_target_encoding',
		default		=> 'UTF-8',
		required	=> 1,
	);

#########1 Public Methods     3#########4#########5#########6#########7#########8#########9

sub	change_output_encoding{
	my ( $self, $string, ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD				name_space 	=> $self->get_log_space . '::change_output_encoding', );
	###LogSD		$phone->talk( level => 'info', message => [
	###LogSD				"Changing the encoding of: $string",
	###LogSD				'..to encoding type: ' . $self->get_target_encoding ] );
	return $string;
}

#########1 Private Attributes 3#########4#########5#########6#########7#########8#########9

has _defined_excel_translations =>(
		isa		=> ArrayRef,
		traits	=> ['Array'],
		default	=> sub{ [
						'General',
						'0',
						'0.00',
						'#,##0',
						'#,##0.00',
						'$#,##0_);($#,##0)',
						'$#,##0_);[Red]($#,##0)',
						'$#,##0.00_);($#,##0.00)',
						'$#,##0.00_);[Red]($#,##0.00)',
						'0%',
						'0.00%',
						'0.00E+00',
						'# ?/?',
						'# ??/??',
						'yyyy-m-d',      # Was 'm-d-yy', which is bad as system default
						'd-mmm-yy',
						'd-mmm',
						'mmm-yy',
						'h:mm AM/PM',
						'h:mm:ss AM/PM',
						'h:mm',
						'h:mm:ss',
						'm-d-yy h:mm',
						undef, undef, undef, undef, undef, undef, undef, undef,
						'#,##0_);(#,##0)',
						'#,##0_);[Red](#,##0)',
						'#,##0.00_);(#,##0.00)',
						'#,##0.00_);[Red](#,##0.00)',
						'_(*#,##0_);_(*(#,##0);_(*"-"_);_(@_)',
						'_($*#,##0_);_($*(#,##0);_($*"-"_);_(@_)',
						'_(*#,##0.00_);_(*(#,##0.00);_(*"-"??_);_(@_)',
						'_($*#,##0.00_);_($*(#,##0.00);_($*"-"??_);_(@_)',
						'mm:ss',
						'[h]:mm:ss',
						'mm:ss.0',
						'##0.0E+0',
						'@'
					]
		},
		reader => 'get_defined_excel_format_list',
		writer => 'set_defined_excel_format_list',
		handles =>{
			get_defined_excel_format => 'get',
			total_defined_excel_formats => 'count',
		},
	);

#########1 Private Methods    3#########4#########5#########6#########7#########8#########9



#########1 Phinish            3#########4#########5#########6#########7#########8#########9

no Moose::Role;
1;

#########1 Documentation      3#########4#########5#########6#########7#########8#########9
__END__

=head1 NAME

Spreadsheet::XLSX::Reader::LibXML::FmtDefault - Default xlsx number formats and localization
    
=head1 DESCRIPTION

POD not written yet!

=cut

#########1#########2 main pod documentation end  5#########6#########7#########8#########9