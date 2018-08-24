<?php
/**
 * Created by PhpStorm.
 * User: emeka
 * Date: 8/14/18
 * Time: 11:23 AM
 */

require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

$DB_Server = "xxxxxx"; //MySQL Server
$DB_Username = "root"; //MySQL Username
$DB_Password = "root";             //MySQL Password

//create MySQL connection
$conn = mysqli_connect($DB_Server, $DB_Username, $DB_Password);

if (!$conn) {
    die("Connection failed: " . mysqli_connect_error());
}

$sql = <<<EOF
select data_table.clid as 'Caller ID', users_table.name as 'Caller Name',
count(data_table.clid) as 'Number of Calls', DATE(data_table.calldate) as 'Call Date' 
from asteriskcdrdb.cdr as data_table
join asterisk.users as users_table
on data_table.clid = users_table.extension
where DATE(data_table.calldate) = DATE(NOW()) and data_table.disposition = 'ANSWERED' and
data_table.clid IN(6009, 6008, 6059, 6045, 6062, 6006, 6028, 6015, 6049, 6061, 6003, 6020, 
6021, 6001, 6011, 6002, 6027, 6005, 6013, 6025, 6026)
LIMIT 10
EOF;

$result = mysqli_query($conn, $sql);

$spreadsheetObj = new Spreadsheet();

// writer already created the first sheet for us, let's get it
$excelSheet = $spreadsheetObj->getActiveSheet();

$downloadersInfo = [];

while ($row = $result->fetch_assoc()) {
    $downloadersInfo[] = $row;
}

//set default font-style
$excelSheet->getDefaultStyle()->getFont()->setName('Calibri');

$excelSheet->getDefaultStyle()->getFont()->setSize(12);

// number format, with thousands separator and two decimal points.
$numberFormat = '#,#0.##;[Red]-#,#0.##';

// rename the sheet
$excelSheet->setTitle('Call Log - '. date('Y-m-d'));

// let's bold and size the header font and write the header
// as you can see, we can specify a range of cells, like here: cells from A1 to C1
$excelSheet->getStyle('A1:E1')->getFont()->setBold(true)->setSize(12);

// write header
$excelSheet->getCell('A1')->setValue('Caller ID');
$excelSheet->getCell('B1')->setValue('Caller Name');
$excelSheet->getCell('C1')->setValue('Number of Calls');
$excelSheet->getCell('D1')->setValue('Call Date');

foreach ($downloadersInfo as $key => $data) {
    $excelSheet->getCell('A' . ($key + 2))->setValue($data['Caller ID']);
    $excelSheet->getCell('B' . ($key + 2))->setValue($data['Caller Name']);
    $excelSheet->getCell('C' . ($key + 2))->setValue($data['Number of Calls']);
    $excelSheet->getCell('D' . ($key + 2))->setValue($data['Call Date']);

    // autosize the columns
    $excelSheet->getColumnDimension('A')->setAutoSize(true);
    $excelSheet->getColumnDimension('B')->setAutoSize(true);
    $excelSheet->getColumnDimension('C')->setAutoSize(true);
    $excelSheet->getColumnDimension('D')->setAutoSize(true);
}

$writer = new Xlsx($spreadsheetObj);
$writer->save('hello world.xlsx');