package ExampleTypes;
		
use strict;
use warnings;
use Type::Utils -all;
use Types::Standard qw(
		Num						Str
		InstanceOf				Maybe
	);
use Type::Library 0.046
	-base,
	-declare => qw(
		PositiveNum				DateTimeType
		DateTimeStringOneType	DateTimeStringTwoType
		
	);
use DateTimeX::Format::Excel;
use DateTime::Format::Flexible;

#########1 Settings 2#########3#########4#########5#########6#########7#########8#########9

my	$number_converter	= DateTimeX::Format::Excel->new( system_type => 'apple_excel' );

#########1 Initial filter type          4#########5#########6#########7#########8#########9
	
declare PositiveNum,
	as Num, where{ $_ > 0 };

#########1 Intermediate type  3#########4#########5#########6#########7#########8#########9
	
declare DateTimeType,
	as InstanceOf[ 'DateTime' ];

#########1 Initial coercions  3#########4#########5#########6#########7#########8#########9
	
coerce DateTimeType,
	from PositiveNum, via{
		my	$num = $_[0];
		return $number_converter->parse_datetime( $num );
	},
	from Str, via{ 
		my	$str = $_[0];
		return DateTime::Format::Flexible->parse_datetime( $str );
	};

#########1 Final types        3#########4#########5#########6#########7#########8#########9

declare DateTimeStringOneType,
	as Maybe[Str], where{ !$_[0] or $_[0] =~ /\d{4}\-\d{2}\-\d{2}/ };

my $match = qr/^[A-Za-z]+, \d+ of [A-Za-z]+, \d{4}$/;
declare DateTimeStringTwoType,
	as Maybe[Str], where{ !$_[0] or $_[0] =~ $match },
	message{ "|$_[0]| could not match: $match" };#

#########1 Final coercions    3#########4#########5#########6#########7#########8#########9

coerce DateTimeStringOneType,
	from DateTimeType->coercibles, via{
		my $tmp = to_DateTimeType( $_ );
		$tmp->format_cldr( 'yyyy-MM-dd' );
	};

coerce DateTimeStringTwoType,
	from DateTimeType->coercibles, via{
		my $tmp = to_DateTimeType( $_ );
		$tmp->format_cldr( 'EEEE, d of MMMM, yyyy' );
	};