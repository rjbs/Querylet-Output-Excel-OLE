use strict;
use warnings;
use ExtUtils::MakeMaker;

WriteMakefile(
    NAME                => 'Querylet::Output::Excel::OLE',
    AUTHOR              => 'Ricardo Signes <rjbs@cpan.org>',
    VERSION_FROM        => 'lib/Querylet/Output/Excel/OLE.pm',
    ABSTRACT_FROM       => 'lib/Querylet/Output/Excel/OLE.pm',
    PL_FILES            => {},
    PREREQ_PM => {
        'Querylet'   => 0.30,
        'Test::More' => 0,
        'Win32::OLE' => 0.17,
    },
    dist                => { COMPRESS => 'gzip -9f', SUFFIX => 'gz', },
    clean               => { FILES => 'Querylet-Output-Excel-OLE-*' },
);
