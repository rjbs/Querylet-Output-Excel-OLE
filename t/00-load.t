use Test::More tests => 2;

BEGIN {
	use_ok('Querylet::Query');
  use_ok('Querylet::Output::Excel::OLE');
}

diag( "Testing Querylet::Output::Excel::OLE $Querylet::Output::Excel::OLE::VERSION" );
