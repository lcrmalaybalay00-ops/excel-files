<?php
/**
 * Birth Statistics Export (PhpSpreadsheet)
 * PHP port that writes Excel files using PhpSpreadsheet.
 *
 * Requirements:
 *  - Composer + phpoffice/phpspreadsheet
 *    Install: composer require phpoffice/phpspreadsheet
 *
 * Notes:
 *  - Keeps the HTML frontend design you provided unchanged (only server-side file).
 *  - Assumes classes/SecurityHelper.php and classes/MySQL_DatabaseManager.php exist.
 */

declare(strict_types=1);

require_once 'config/config.php';
require_once 'classes/SecurityHelper.php';
require_once 'classes/MySQL_DatabaseManager.php';
require_once __DIR__ . '/vendor/autoload.php'; // Composer autoload (PhpSpreadsheet)

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\NumberFormat;

SecurityHelper::requireLogin();

/* ---------- Helpers ---------- */

function loadMunicipalityRef(): array {
    $paths = [
        'C:/PhilCRIS/Resources/References/RMunicipality.ref',
        __DIR__ . '/Resources/References/RMunicipality.ref',
        __DIR__ . '/references/RMunicipality.ref',
    ];
    $dict = [];
    foreach ($paths as $p) {
        if (!file_exists($p)) continue;
        foreach (file($p, FILE_IGNORE_NEW_LINES | FILE_SKIP_EMPTY_LINES) as $line) {
            $parts = explode('|', $line);
            if (count($parts) >= 4) {
                $code = trim($parts[3]);
                if ($code !== '' && !isset($dict[$code])) {
                    $dict[$code] = [
                        'municipality' => trim($parts[0]),
                        'province'     => trim($parts[1]),
                        'country'      => trim($parts[2]),
                    ];
                }
            }
        }
        break;
    }
    return $dict;
}

function weekFields(string $dateStr): array {
    $dt    = new DateTime($dateStr);
    $y     = (int)$dt->format('Y');
    $m     = (int)$dt->format('n');
    $d     = (int)$dt->format('j');
    $first = new DateTime("$y-$m-01");
    $last  = new DateTime($dt->format('Y-m-t'));
    $wn    = (int)(($d - 1) / 7) + 1;
    $ws    = clone $first; $ws->modify('+' . (($wn - 1) * 7) . ' days');
    $we    = clone $ws;    $we->modify('+6 days');
    if ($we > $last) $we = clone $last;
    return ['wn'=>$wn, 'ws'=>$ws->format('Y-m-d'), 'we'=>$we->format('Y-m-d'),
            'year'=>$y, 'month'=>$m, 'day'=>$d];
}

function safeInt($v, $default = null) {
    if ($v === null || $v === '') return $default;
    return is_numeric($v) ? (int)$v : $default;
}

function colLetter(int $index): string {
    $letters = '';
    while ($index > 0) {
        $mod = ($index - 1) % 26;
        $letters = chr(65 + $mod) . $letters;
        $index = (int)(($index - 1) / 26);
    }
    return $letters;
}

/* ---------- Export handler ---------- */
$exportError   = null;
$exportSuccess = false;
$savedFile     = '';

if ($_SERVER['REQUEST_METHOD'] === 'POST' && ($_POST['action'] ?? '') === 'export_birth') {
    $year       = intval($_POST['year'] ?? date('Y'));
    $monthStart = intval($_POST['month_start'] ?? 1);
    $monthEnd   = intval($_POST['month_end'] ?? 12);
    $teenAge    = isset($_POST['teenage_age']) && trim($_POST['teenage_age']) !== '' ? intval($_POST['teenage_age']) : 19;
    $srcPartner = !empty($_POST['source_partner']);
    $srcLGU     = !empty($_POST['source_lgu']);
    $savePath   = trim($_POST['save_path'] ?? '');

    if (!$srcPartner && !$srcLGU) {
        $exportError = 'Please select at least one Source (Partner or LGU Register).';
    } elseif (empty($savePath)) {
        $exportError = 'Please enter a save path for the exported file.';
    } else {
        try {
            // Template
            $templatePath = __DIR__ . DIRECTORY_SEPARATOR . 'ExcelTemplate' . DIRECTORY_SEPARATOR . 'birthtemplate.xlsx';
            if (!file_exists($templatePath)) throw new Exception('Template not found: ExcelTemplate/birthtemplate.xlsx');

            // Destination filename
            $filename = $year . '_Birth_Statistics_Reports_' . date('Ymd_His') . '.xlsx';
            $destPath = is_dir($savePath)
                ? rtrim($savePath, '/\\') . DIRECTORY_SEPARATOR . $filename
                : preg_replace('/\.xlsx$/i', '', $savePath) . '.xlsx';

            // Copy template
            copy($templatePath, $destPath);

            // DB
            $dbManager = new MySQL_DatabaseManager();
            $conn = $dbManager->getMainConnection(); // expects mysqli
            if (!$conn) throw new Exception('Database connection failed.');

            // Source where
            $srcParts = [];
            if ($srcPartner) $srcParts[] = "(RegistryNum LIKE '!%' OR RegistryNum REGEXP '^[0-9]{4}-[0-9]+$')";
            if ($srcLGU)     $srcParts[] = "LEFT(RegistryNum,1) != '!'";
            $srcWhere = '(' . implode(' OR ', $srcParts) . ')';

            // Main SQL (fields based on your provided snippet)
            $sql = "
                SELECT RegistryNum, DocumentStatus,
                       CFirstName, CMiddleName, CLastName, CSexId, CBirthDate,
                       CBirthAddress, CBirthMunicipality, CBirthMunicipalityId,
                       CBirthProvince, CBirthProvinceId, CBirthCountry, CBirthCountryId, CBirthTypeId,
                       MFirstName, MMiddleName, MLastName, MCitizenship, MCitizenshipId,
                       MOccupation, MOccupationId, MAge,
                       MAddress, MMunicipality, MMunicipalityId, MProvince, MProvinceId, MCountry, MCountryId,
                       FFirstName, FMiddleName, FLastName, FCitizenship, FCitizenshipId,
                       FOccupation, FOccupationId, FAge,
                       FAddress, FMunicipality, FMunicipalityId, FProvince, FProvinceId, FCountry, FCountryId,
                       AttendantId, AttendantName, AttendantTitle,
                       PreparerName, PreparerTitle, PreparerDate,
                       DateReceived, DateRegistered
                FROM phcris.birthdocument
                WHERE $srcWhere
                  AND YEAR(CBirthDate) = ?
                  AND MONTH(CBirthDate) BETWEEN ? AND ?
                ORDER BY LEFT(RegistryNum,4),
                         CAST(SUBSTRING_INDEX(SUBSTRING_INDEX(RegistryNum,'-',-1),'-',1) AS UNSIGNED) ASC
            ";

            $stmt = $conn->prepare($sql);
            if ($stmt === false) throw new Exception('SQL prepare failed: ' . $conn->error);
            $stmt->bind_param('iii', $year, $monthStart, $monthEnd);
            $stmt->execute();
            $res = $stmt->get_result();
            $records = $res->fetch_all(MYSQLI_ASSOC);
            $stmt->close();

            if (empty($records)) {
                $exportError = 'No records found for the selected filters.';
                throw new Exception($exportError);
            }

            // Add computed week fields
            foreach ($records as &$r) {
                $bd = $r['CBirthDate'] ?? '';
                if ($bd !== '' && strtotime($bd) !== false) {
                    $w = weekFields($bd);
                    $r['Week Number']     = $w['wn'];
                    $r['Week Start Date'] = $w['ws'];
                    $r['Week End Date']   = $w['we'];
                    $r['Year']            = $w['year'];
                    $r['Month']           = $w['month'];
                    $r['Day']             = $w['day'];
                } else {
                    $r['Week Number'] = $r['Week Start Date'] = $r['Week End Date']
                                      = $r['Year'] = $r['Month'] = $r['Day'] = null;
                }
                // Ensure numeric ages
                $r['MAgeNumeric'] = safeInt($r['MAge'] ?? null, null);
                $r['FAgeNumeric'] = safeInt($r['FAge'] ?? null, null);
            }
            unset($r);

            // Municipality statistics (group by mother's municipality MMunicipality if present, else CBirthMunicipality)
            $munRef = loadMunicipalityRef();
            $munGroups = []; // code => [male, female, total]
            foreach ($records as $r) {
                $sex = strtoupper(trim((string)($r['CSexId'] ?? '')));
                // Choose municipality value: prefer MMunicipality, fallback to CBirthMunicipality
                $munVal = trim((string)($r['MMunicipality'] ?? $r['CBirthMunicipality'] ?? ''));
                $code = 'UNKNOWN';
                if ($munVal !== '' && strpos($munVal, '|') !== false) {
                    $code = trim(substr($munVal, strrpos($munVal, '|') + 1));
                    if ($code === '') $code = 'UNKNOWN';
                }
                if (!isset($munGroups[$code])) $munGroups[$code] = [0,0,0];
                if ($sex === 'MALE') $munGroups[$code][0]++;
                elseif ($sex === 'FEMALE') $munGroups[$code][1]++;
                $munGroups[$code][2]++;
            }

            // Sort with UNKNOWN last
            uksort($munGroups, fn($a,$b) => ($a==='UNKNOWN'?'ZZZZZ':$a) <=> ($b==='UNKNOWN'?'ZZZZZ':$b));

            $munStats = [];
            $sNo = 1;
            foreach ($munGroups as $code => [$male, $female, $total]) {
                if ($code === 'UNKNOWN') {
                    [$mun,$prov,$ctry] = ['Not Stated','Not Stated','Philippines'];
                } elseif (isset($munRef[$code])) {
                    $mun  = $munRef[$code]['municipality'];
                    $prov = $munRef[$code]['province'];
                    $ctry = $munRef[$code]['country'];
                } else {
                    [$mun,$prov,$ctry] = ["Unknown ($code)",'Unknown','Philippines'];
                }
                $munStats[] = ['no'=>$sNo++,'mun'=>$mun,'prov'=>$prov,'ctry'=>$ctry,
                               'male'=>$male,'female'=>$female,'total'=>$total];
            }

            // TeenAge: mothers with MAgeNumeric <= $teenAge grouped by mother's municipality (MMunicipality)
            $teenGroups = []; // code => count
            foreach ($records as $r) {
                $mage = safeInt($r['MAge'] ?? null, null);
                if ($mage === null) continue;
                if ($mage <= $teenAge) {
                    $munVal = trim((string)($r['MMunicipality'] ?? ''));
                    $code = 'UNKNOWN';
                    if ($munVal !== '' && strpos($munVal, '|') !== false) {
                        $code = trim(substr($munVal, strrpos($munVal, '|') + 1));
                        if ($code === '') $code = 'UNKNOWN';
                    }
                    if (!isset($teenGroups[$code])) $teenGroups[$code] = 0;
                    $teenGroups[$code]++;
                }
            }
            uksort($teenGroups, fn($a,$b) => ($a==='UNKNOWN'?'ZZZZZ':$a) <=> ($b==='UNKNOWN'?'ZZZZZ':$b));

            $teenStats = [];
            $tNo = 1;
            foreach ($teenGroups as $code => $count) {
                if ($code === 'UNKNOWN') {
                    [$mun,$prov,$ctry] = ['Not Stated','Not Stated','Philippines'];
                } elseif (isset($munRef[$code])) {
                    $mun  = $munRef[$code]['municipality'];
                    $prov = $munRef[$code]['province'];
                    $ctry = $munRef[$code]['country'];
                } else {
                    [$mun,$prov,$ctry] = ["Unknown ($code)",'Unknown','Philippines'];
                }
                $teenStats[] = ['no'=>$tNo++,'mun'=>$mun,'prov'=>$prov,'ctry'=>$ctry,'count'=>$count];
            }

            /* ---------- Write to Excel (PhpSpreadsheet) ---------- */
            $spreadsheet = IOFactory::load($destPath);

            // Data Source sheet
            $columns = [
                'RegistryNum','DocumentStatus','CFirstName','CMiddleName','CLastName',
                'CSexId','CBirthDate','CBirthAddress','CBirthMunicipality','CBirthMunicipalityId',
                'CBirthProvince','CBirthProvinceId','CBirthCountry','CBirthCountryId','CBirthTypeId',
                'MFirstName','MMiddleName','MLastName','MCitizenship','MCitizenshipId',
                'MOccupation','MOccupationId','MAge','MAddress','MMunicipality','MMunicipalityId',
                'MProvince','MProvinceId','MCountry','MCountryId',
                'FFirstName','FMiddleName','FLastName','FCitizenship','FCitizenshipId',
                'FOccupation','FOccupationId','FAge','FAddress','FMunicipality','FMunicipalityId',
                'FProvince','FProvinceId','FCountry','FCountryId',
                'AttendantId','AttendantName','AttendantTitle',
                'PreparerName','PreparerTitle','PreparerDate',
                'DateReceived','DateRegistered',
                'Week Number','Week Start Date','Week End Date',
                'Year','Month','Day',
            ];

            $dsSheet = $spreadsheet->getSheetByName('Data Source');
            if ($dsSheet === null) {
                $dsSheet = new Worksheet($spreadsheet, 'Data Source');
                $spreadsheet->addSheet($dsSheet);
            }
            // Write headers
            foreach ($columns as $ci => $col) {
                $cell = colLetter($ci + 1) . '1';
                $dsSheet->setCellValue($cell, $col);
                $dsSheet->getStyle($cell)->getFont()->setBold(true);
            }
            // Write data rows
            $rn = 2;
            foreach ($records as $rec) {
                foreach ($columns as $ci => $col) {
                    $val = $rec[$col] ?? '';
                    // Format birth date as Y-m-d if possible
                    if (($col === 'CBirthDate' || $col === 'PreparerDate' || $col === 'DateReceived' || $col === 'DateRegistered') && $val !== '' && strtotime($val) !== false) {
                        $val = (new DateTime($val))->format('Y-m-d');
                    }
                    $cell = colLetter($ci + 1) . $rn;
                    $dsSheet->setCellValue($cell, $val);
                }
                $rn++;
            }
            // Auto-size columns
            $lastCol = count($columns);
            for ($c = 1; $c <= $lastCol; $c++) {
                $dsSheet->getColumnDimension(colLetter($c))->setAutoSize(true);
            }

            // ByMunicipality sheet (data starts row 14)
            $munSheet = $spreadsheet->getSheetByName('ByMunicipality');
            if ($munSheet !== null && !empty($munStats)) {
                $startRow = 14;
                $r = $startRow;
                foreach ($munStats as $m) {
                    $munSheet->setCellValue("B{$r}", $m['no']);
                    $munSheet->setCellValue("C{$r}", $m['mun']);
                    $munSheet->setCellValue("D{$r}", $m['prov']);
                    $munSheet->setCellValue("E{$r}", $m['ctry']);
                    $munSheet->setCellValue("F{$r}", $m['male']);
                    $munSheet->getStyle("F{$r}")->getNumberFormat()->setFormatCode('#,##0');
                    $munSheet->setCellValue("G{$r}", $m['female']);
                    $munSheet->getStyle("G{$r}")->getNumberFormat()->setFormatCode('#,##0');
                    $munSheet->setCellValue("H{$r}", "=F{$r}+G{$r}");
                    $munSheet->getStyle("H{$r}")->getNumberFormat()->setFormatCode('#,##0');
                    $r++;
                }
                // Total row
                $totR = $r;
                $lastDR = $r - 1;
                $munSheet->mergeCells("B{$totR}:E{$totR}");
                $munSheet->setCellValue("B{$totR}", 'TOTAL');
                $munSheet->getStyle("B{$totR}")->getFont()->setBold(true);
                $munSheet->getStyle("B{$totR}")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
                $munSheet->setCellValue("F{$totR}", "=SUM(F{$startRow}:F{$lastDR})");
                $munSheet->setCellValue("G{$totR}", "=SUM(G{$startRow}:G{$lastDR})");
                $munSheet->setCellValue("H{$totR}", "=SUM(H{$startRow}:H{$lastDR})");
                foreach (['F','G','H'] as $col) {
                    $munSheet->getStyle("{$col}{$totR}")->getNumberFormat()->setFormatCode('#,##0');
                    $munSheet->getStyle("{$col}{$totR}")->getFont()->setBold(true);
                }
                // Borders
                $drng = "B{$startRow}:H{$totR}";
                $munSheet->getStyle($drng)->getBorders()->getAllBorders()->setBorderStyle(Border::BORDER_THIN);
                $munSheet->getStyle($drng)->getBorders()->getOutline()->setBorderStyle(Border::BORDER_MEDIUM);
            }

            // TeenAge sheet (data starts row 15)
            $teenSheet = $spreadsheet->getSheetByName('TeenAge');
            if ($teenSheet !== null && !empty($teenStats)) {
                $startRow = 15;
                $r = $startRow;
                foreach ($teenStats as $t) {
                    $teenSheet->setCellValue("B{$r}", $t['no']);
                    $teenSheet->setCellValue("C{$r}", $t['mun']);
                    $teenSheet->setCellValue("D{$r}", $t['prov']);
                    $teenSheet->setCellValue("E{$r}", $t['ctry']);
                    $teenSheet->setCellValue("F{$r}", $t['count']);
                    $teenSheet->getStyle("F{$r}")->getNumberFormat()->setFormatCode('#,##0');
                    $r++;
                }
                // Total row
                $totR = $r;
                $lastDR = $r - 1;
                $teenSheet->mergeCells("B{$totR}:E{$totR}");
                $teenSheet->setCellValue("B{$totR}", 'TOTAL');
                $teenSheet->getStyle("B{$totR}")->getFont()->setBold(true);
                $teenSheet->getStyle("B{$totR}")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
                $teenSheet->setCellValue("F{$totR}", "=SUM(F{$startRow}:F{$lastDR})");
                $teenSheet->getStyle("F{$totR}")->getNumberFormat()->setFormatCode('#,##0');
                $teenSheet->getStyle("F{$totR}")->getFont()->setBold(true);

                // Borders
                $trng = "B{$startRow}:F{$totR}";
                $teenSheet->getStyle($trng)->getBorders()->getAllBorders()->setBorderStyle(Border::BORDER_THIN);
                $teenSheet->getStyle($trng)->getBorders()->getOutline()->setBorderStyle(Border::BORDER_MEDIUM);
            }

            // Try a light calculation pass (Excel will recalc on open if needed)
            try {
                \PhpOffice\PhpSpreadsheet\Calculation\Calculation::getInstance($spreadsheet)->calculate();
            } catch (\Throwable $ex) {
                // ignore
            }

            // Save workbook
            $writer = IOFactory::createWriter($spreadsheet, 'Xlsx');
            $writer->save($destPath);

            // Cleanup
            $spreadsheet->disconnectWorksheets();
            unset($spreadsheet, $writer);

            $exportSuccess = true;
            $savedFile = $destPath;

        } catch (Exception $ex) {
            $exportError = 'Export failed: ' . $ex->getMessage();
        }
    }
}

/* ---------- HTML FORM VARIABLES (kept same as provided UI) ---------- */
$currentYear  = (int)date('Y');
$fYear        = intval($_POST['year']        ?? $currentYear);
$fM1          = intval($_POST['month_start'] ?? 1);
$fM2          = intval($_POST['month_end']   ?? 12);
$fTeen        = $_POST['teenage_age'] ?? '';
$fPartner     = !empty($_POST['source_partner']);
$fLGU         = !empty($_POST['source_lgu']);
$fSavePath    = $_POST['save_path'] ?? '';
$months       = ['January','February','March','April','May','June',
                 'July','August','September','October','November','December'];
?>
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Birth Statistics Export</title>
<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
<link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css" rel="stylesheet">
<style>
*{margin:0;padding:0;box-sizing:border-box}
body{background:#0a0e27;color:#fff;font-family:'Segoe UI',sans-serif;min-height:100vh;display:flex;align-items:center;justify-content:center;padding:20px 0}
.card{background:#1e293b;border:2px solid #667eea;border-radius:12px;width:540px;max-width:96vw;box-shadow:0 20px 60px rgba(0,0,0,.6);overflow:hidden}
.card-head{background:linear-gradient(135deg,#0d47a1,#1565c0);padding:16px 22px;display:flex;align-items:center;gap:12px}
.card-head h3{margin:0;font-size:1.1rem;font-weight:700;color:#fff}
.card-head i{font-size:1.4rem;color:#00d9ff}
.card-body{padding:22px}
.frow{display:grid;grid-template-columns:185px 14px 1fr;align-items:start;gap:8px;margin-bottom:13px}
.frow label{color:#94a3b8;font-weight:600;font-size:.875rem;padding-top:8px}
.frow span{color:#94a3b8;font-weight:700;padding-top:8px}
.form-control,.form-select{background:#0f172a;border:1px solid #334155;color:#fff;padding:7px 10px;border-radius:6px;font-size:.875rem;width:100%}
.form-control:focus,.form-select:focus{background:#0f172a;border-color:#667eea;color:#fff;box-shadow:0 0 0 3px rgba(102,126,234,.2);outline:none}
.form-control option,.form-select option{background:#0f172a;color:#fff}
.sbox{background:rgba(102,126,234,.08);border:1px solid #334155;border-radius:8px;padding:13px 15px;margin-bottom:13px}
.sbox-title{font-size:.73rem;font-weight:700;text-transform:uppercase;color:#00d9ff;margin-bottom:9px;letter-spacing:.05em}
.chk{display:flex;align-items:center;gap:10px;padding:5px 0}
.chk input[type=checkbox]{width:17px;height:17px;accent-color:#667eea;cursor:pointer;flex-shrink:0}
.chk label{color:#e2e8f0;font-size:.88rem;cursor:pointer;margin:0}
.chk small{color:#64748b;font-size:.73rem}
.divider{border:none;border-top:1px solid #334155;margin:14px 0}
.btn-row{display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-top:18px}
.btn-ok{background:linear-gradient(135deg,#667eea,#764ba2);color:#fff;border:none;padding:11px;border-radius:7px;font-weight:700;font-size:.9rem;cursor:pointer;transition:opacity .2s}
.btn-ok:hover{opacity:.88}
.btn-ok:disabled{opacity:.45;cursor:not-allowed}
.btn-cancel{background:#1e293b;color:#94a3b8;border:1px solid #334155;padding:11px;border-radius:7px;font-weight:700;font-size:.9rem;cursor:pointer}
.btn-cancel:hover{background:#334155;color:#fff}
.alert-err{background:rgba(239,68,68,.15);border:1px solid #ef4444;color:#fca5a5;border-radius:6px;padding:10px 14px;font-size:.875rem;margin-bottom:15px;display:flex;align-items:flex-start;gap:8px}
.alert-ok{background:rgba(34,197,94,.15);border:1px solid #22c55e;color:#86efac;border-radius:6px;padding:10px 14px;font-size:.875rem;margin-bottom:15px;display:flex;align-items:flex-start;gap:8px}
.hint{font-size:.73rem;color:#64748b;margin-top:4px}
.info-box{background:rgba(0,217,255,.05);border:1px solid #00d9ff33;border-radius:8px;padding:13px 15px;margin-bottom:13px}
.info-title{font-size:.73rem;font-weight:700;text-transform:uppercase;color:#00d9ff;margin-bottom:9px;letter-spacing:.05em}
.info-box p{font-size:.82rem;color:#94a3b8;line-height:1.9;margin:0}
.save-path-row{display:flex;gap:6px;align-items:center}
.save-path-row .form-control{flex:1}
</style>
</head>
<body>
<div class="card">
    <div class="card-head">
        <i class="fas fa-file-excel"></i>
        <h3>Birth Statistics Export</h3>
    </div>
    <div class="card-body">
        <?php if ($exportSuccess): ?>
        <div class="alert-ok">
            <i class="fas fa-check-circle mt-1"></i>
            <div>
                <strong>Export completed successfully!</strong><br>
                <span style="font-size:.8rem">File saved to: <code><?= htmlspecialchars($savedFile ?? '') ?></code></span>
            </div>
        </div>
        <?php elseif ($exportError): ?>
        <div class="alert-err">
            <i class="fas fa-exclamation-circle mt-1"></i>
            <span><?= nl2br(htmlspecialchars($exportError)) ?></span>
        </div>
        <?php endif; ?>
        <form id="frm" method="POST" action="birth_statistics_export.php">
        <input type="hidden" name="action" value="export_birth">
        <!-- Year -->
        <div class="frow">
            <label>Year</label><span>:</span>
            <select name="year" class="form-select" required>
                <?php for($y=$currentYear;$y>=1900;$y--): ?>
                <option value="<?=$y?>" <?=$y===$fYear?'selected':''?>><?=$y?></option>
                <?php endfor; ?>
            </select>
        </div>
        <!-- Month Start -->
        <div class="frow">
            <label>Month Start</label><span>:</span>
            <select name="month_start" class="form-select" required id="m1">
                <?php for($m=1;$m<=12;$m++): ?>
                <option value="<?=$m?>" <?=$m===$fM1?'selected':''?>><?=$m?> — <?=$months[$m-1]?></option>
                <?php endfor; ?>
            </select>
        </div>
        <!-- Month End -->
        <div class="frow">
            <label>Month End</label><span>:</span>
            <select name="month_end" class="form-select" required id="m2">
                <?php for($m=1;$m<=12;$m++): ?>
                <option value="<?=$m?>" <?=$m===$fM2?'selected':''?>><?=$m?> — <?=$months[$m-1]?></option>
                <?php endfor; ?>
            </select>
        </div>
        <!-- Teenage Pregnancy Age -->
        <div class="frow">
            <label>Teenage Pregnancy Age</label><span>:</span>
            <div>
                <input type="number" name="teenage_age" class="form-control"
                       placeholder="e.g. 19  (default: 19)" min="1" max="100"
                       value="<?=htmlspecialchars($fTeen)?>">
                <div class="hint">Filters TeenAge sheet: mothers age ≤ this value</div>
            </div>
        </div>
        <hr class="divider">
        <!-- Source -->
        <div class="sbox">
            <div class="sbox-title"><i class="fas fa-database me-1"></i>Source</div>
            <div class="chk">
                <input type="checkbox" name="source_partner" id="sp" value="1" <?=$fPartner?'checked':''?>>
                <label for="sp">Partner <small>(Registry starts with "!" or yyyy-nnnnn)</small></label>
            </div>
            <div class="chk">
                <input type="checkbox" name="source_lgu" id="sl" value="1" <?=$fLGU?'checked':''?>>
                <label for="sl">LGU Register <small>(Registry does NOT start with "!")</small></label>
            </div>
        </div>
        <!-- Save Path -->
        <div class="sbox">
            <div class="sbox-title"><i class="fas fa-folder-open me-1"></i>Save Location</div>
            <div>
                <input type="text" name="save_path" class="form-control"
                       placeholder="e.g. C:\Reports  or  C:\Reports\MyFile.xlsx"
                       value="<?=htmlspecialchars($fSavePath)?>" required>
                <div class="hint">
                    Enter a folder (filename auto-generated) or a full file path.<br>
                    The file will open automatically after export — same as the desktop app.
                </div>
            </div>
        </div>
        <!-- Info box -->
        <div class="info-box">
            <div class="info-title"><i class="fas fa-info-circle me-1"></i>Sheets populated</div>
            <p>
                <i class="fas fa-table me-1 text-success"></i>
                <strong style="color:#e2e8f0">Data Source</strong> — all records + week/date computed columns<br>
                <i class="fas fa-map-marker-alt me-1 text-warning"></i>
                <strong style="color:#e2e8f0">ByMunicipality</strong> — Male / Female / Total per municipality (row 14)<br>
                <i class="fas fa-baby me-1" style="color:#f472b6"></i>
                <strong style="color:#e2e8f0">TeenAge</strong> — teenage pregnancy grouped by municipality (row 15)
            </p>
        </div>
        <div class="btn-row">
            <button type="submit" class="btn-ok" id="btnGo">
                <i class="fas fa-file-export me-2"></i>Process &amp; Export
            </button>
            <button type="button" class="btn-cancel" onclick="history.back()">
                <i class="fas fa-times me-2"></i>Cancel
            </button>
        </div>
        </form>
        <?php if ($exportSuccess && !empty($savedFile)): ?>
        <script>
        window.onload = function() {
            setTimeout(function() {
                alert(' Export complete!\n\nFile saved to:\n<?= addslashes($savedFile) ?>\n\nPlease open the file from the location above.');
            }, 300);
        };
        </script>
        <?php endif; ?>
    </div>
</div>
<script>
document.getElementById('frm').addEventListener('submit', function(e) {
    const p  = document.getElementById('sp').checked;
    const l  = document.getElementById('sl').checked;
    const m1 = parseInt(document.getElementById('m1').value);
    const m2 = parseInt(document.getElementById('m2').value);
    if (!p && !l) {
        e.preventDefault();
        alert(' Please select at least one Source (Partner or LGU Register).');
        return;
    }
    if (m2 < m1) {
        e.preventDefault();
        alert(' Month End cannot be before Month Start.');
        return;
    }
    const btn = document.getElementById('btnGo');
    btn.disabled = true;
    btn.innerHTML = '<i class="fas fa-spinner fa-spin me-2"></i>Processing... please wait';
    setTimeout(() => {
        btn.disabled = false;
        btn.innerHTML = '<i class="fas fa-file-export me-2"></i>Process &amp; Export';
    }, 60000);
});
</script>
</body>
</html>