#!/usr/bin/perl

use Spreadsheet::WriteExcel;

my ($filename,$total,$email)=@ARGV;


#print @ARGV;

if (not defined $filename){
	die ("Specify xls filename, please");
	}


my $workbook = Spreadsheet::WriteExcel->new($filename); 
$workbook->set_properties(
        title    => 'This is an example spreadsheet',
        author   => 'Alex Martynov',
        comments => 'Coments',
    );

my $worksheet = $workbook->add_worksheet();



$head = $workbook->add_format();
$head->set_bold();
$head->set_color('white');
$head->set_bg_color('blue');
$head->set_align('center');
$head->set_border(1);

$data = $workbook->add_format();
$data->set_color('black');
$data->set_align('left');
$data->set_border(1);


$c=0;

foreach $line ( <STDIN> ) {
    chomp( $line );
	@row=split(/\t/,  $line);
	for ($i=0; $i<=$#row; $i++)
		{
		if ( $c==0 )
			{
				$worksheet->write($c, $i, $row[$i],$head);
			}
		else
			{
				$worksheet->write($c, $i, $row[$i],$data);
			}
		}
	$c++
}

if (defined $total){
	my @columns= split /,/, $total;
#	print @columns;

	foreach $i  (@columns) {
		print "=SUM(".$i."2:".$i.($c).")\n";
	 my $total_row=$i.($c+1);
	print $total_row
	 $worksheet->write($total_row, "=SUM(".$i."2:".$i.($c).")".$c.")",$head);
	
	} 
}


#$worksheet->write($c, $i-1, '=SUM(B2:B'.$c.")",$head);
print $#row."columns\n";
$workbook->close();


