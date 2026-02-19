<?php
/**
 * Death Statistics Export - Single File with Modal Spinner
 * Clean processing without streaming output
 */
set_time_limit(0);
ini_set('max_execution_time', 0);
ini_set('memory_limit', '512M');

require_once 'config/config.php';
require_once 'classes/SecurityHelper.php';
require_once 'classes/MySQL_DatabaseManager.php';

SecurityHelper::requireLogin();

if (!defined('CP_UTF8')) define('CP_UTF8', 65001);

function loadMunRef(): array {
    $paths = ['C:/PhilCRIS/Resources/References/RMunicipality.ref', __DIR__ . '/Resources/References/RMunicipality.ref', __DIR__ . '/references/RMunicipality.ref'];
    $dict = [];
    foreach ($paths as $p) {
        if (!file_exists($p)) continue;
        foreach (file($p, FILE_IGNORE_NEW_LINES | FILE_SKIP_EMPTY_LINES) as $line) {
            $parts = explode('|', $line);
            if (count($parts) >= 4) {
                $code = trim($parts[3]);
                if ($code !== '' && !isset($dict[$code])) {
                    $dict[$code] = ['municipality' => trim($parts[0]), 'province' => trim($parts[1]), 'country' => trim($parts[2])];
                }
            }
        }
        return $dict;
    }
    return $dict;
}

function weekFields(string $dateStr): array {
    $dt = new DateTime($dateStr);
    $y = (int)$dt->format('Y');
    $m = (int)$dt->format('n');
    $d = (int)$dt->format('j');
    $wn = (int)(($d - 1) / 7) + 1;
    $first = new DateTime("$y-$m-01");
    $ws = clone $first; $ws->modify('+' . (($wn - 1) * 7) . ' days');
    $we = clone $ws; $we->modify('+6 days');
    $last = new DateTime($dt->format('Y-m-t'));
    if ($we > $last) $we = clone $last;
    return ['wn'=>$wn, 'ws'=>$ws->format('Y-m-d'), 'we'=>$we->format('Y-m-d'), 'year'=>$y, 'month'=>$m, 'day'=>$d];
}

function safeInt($v, $default = null) {
    return is_numeric($v) ? (int)$v : $default;
}

// Check for session messages
$exportError = $_SESSION['export_error'] ?? null;
$exportSuccess = $_SESSION['export_success'] ?? false;
$savedFile = $_SESSION['export_file'] ?? '';
unset($_SESSION['export_error'], $_SESSION['export_success'], $_SESSION['export_file']);

// Process export
if ($_SERVER['REQUEST_METHOD'] === 'POST' && ($_POST['action'] ?? '') === 'export_death') {
    
    $year = intval($_POST['year']);
    $monthStart = intval($_POST['month_start']);
    $monthEnd = intval($_POST['month_end']);
    $inclCause = !empty($_POST['include_cause']);
    $cause1 = trim($_POST['cause1'] ?? '');
    $cause2 = trim($_POST['cause2'] ?? '');
    $cause3 = trim($_POST['cause3'] ?? '');
    $srcPartner = !empty($_POST['source_partner']);
    $srcLGU = !empty($_POST['source_lgu']);
    $savePath = trim($_POST['save_path'] ?? '');

    // Validation
    if ($inclCause && $cause1 === '' && $cause2 === '' && $cause3 === '') {
        echo json_encode([
            'success' => false,
            'error' => 'At least one Cause of Death must be entered.'
        ]);
        exit;

    } elseif (!$srcPartner && !$srcLGU) {
        $_SESSION['export_error'] = 'Please select at least one Source.';
        header('Location: death_statistics_export.php');
        exit;
    } elseif (empty($savePath)) {
        $_SESSION['export_error'] = 'Please enter a save path.';
        header('Location: death_statistics_export.php');
        exit;
    }

    try {
        $templatePath = __DIR__ . DIRECTORY_SEPARATOR . 'ExcelTemplate' . DIRECTORY_SEPARATOR . 'deathtemplate.xlsx';
        if (!file_exists($templatePath)) throw new Exception('Template not found: ' . $templatePath);

        $filename = $year . '_Death_Statistics_Reports_' . date('Ymd_His') . '.xlsx';
        if (is_dir($savePath)) {
            $destPath = rtrim($savePath, '/\\') . DIRECTORY_SEPARATOR . $filename;
        } else {
            if (preg_match('/\.xlsx$/i', $savePath)) {
                $destPath = $savePath;
            } else {
                @mkdir($savePath, 0777, true);
                if (is_dir($savePath)) {
                    $destPath = rtrim($savePath, '/\\') . DIRECTORY_SEPARATOR . $filename;
                } else {
                    @mkdir(__DIR__.'/exports', 0777, true);
                    $destPath = __DIR__.'/exports/'.$filename;
                }
            }
        }
        @mkdir(dirname($destPath), 0777, true);

        if (!@copy($templatePath, $destPath)) throw new Exception('Failed to copy template to: ' . dirname($destPath));
        sleep(1);

        $dbManager = new MySQL_DatabaseManager();
        $conn = $dbManager->getMainConnection();
        if (!$conn) throw new Exception('Database connection failed');

        $srcParts = [];
        if ($srcPartner) $srcParts[] = "(RegistryNum LIKE '!%' OR RegistryNum REGEXP '^[0-9]{4}-[0-9]+\$')";
        if ($srcLGU) $srcParts[] = "LEFT(RegistryNum,1) != '!'";
        $srcWhere = '(' . implode(' OR ', $srcParts) . ')';

        $sql = "SELECT RegistryNum, DocumentStatus, CFirstName, CMiddleName, CLastName, CSexId, CDeathDate, CBirthDate, 
                CAgeYears, CAgeMonths, CAgeDays, CAgeHours, CAgeMinutes, CDeathAddress, CDeathMunicipality, CDeathMunicipalityId, 
                CDeathProvince, CDeathProvinceId, CDeathCountry, CDeathCountryId, CCivilStatusId, CReligion, CCitizenship, 
                CResidenceAddress, CResidenceMunicipality, CResidenceMunicipalityId, CResidenceProvince, CResidenceProvinceId, 
                CResidenceCountry, CResidenceCountryId, COccupation, FFirstName, FMiddleName, FLastName, MFirstName, MMiddleName, 
                MLastName, CCauseImmediate, CCauseImmediateId, CCauseImmediateInterval, CCauseAntecedent, CCauseAntecedentId, 
                CCauseAntecedentInterval, CCauseUnderlying, CCauseUnderlyingId, CCauseUnderlyingInterval, CCauseOther, CCauseOtherId, 
                AttendantId, AttendantName, AttendantTitle, AttendantAttendedFrom, PreparerName, PreparerTitle, PreparerDate, 
                DateReceived, DateRegistered 
                FROM phcris.deathdocument 
                WHERE $srcWhere AND YEAR(CDeathDate) = ? AND MONTH(CDeathDate) BETWEEN ? AND ? 
                ORDER BY LEFT(RegistryNum,4), CAST(SUBSTRING_INDEX(SUBSTRING_INDEX(RegistryNum,'-',-1),'-',1) AS UNSIGNED) ASC";

        $stmt = $conn->prepare($sql);
        if (!$stmt) throw new Exception('SQL prepare failed');
        
        $stmt->bind_param('iii', $year, $monthStart, $monthEnd);
        $stmt->execute();
        $res = $stmt->get_result();
        $records = $res->fetch_all(MYSQLI_ASSOC);

        if (empty($records)) throw new Exception('No records found for the selected filters.');

        // Compute week fields
        foreach ($records as &$r) {
            $dd = $r['CDeathDate'] ?? '';
            if ($dd !== '' && strtotime($dd) !== false) {
                $w = weekFields($dd);
                $r['Week Number'] = $w['wn'];
                $r['Week Start Date'] = $w['ws'];
                $r['Week End Date'] = $w['we'];
                $r['Year'] = $w['year'];
                $r['Month'] = $w['month'];
                $r['Day'] = $w['day'];
            } else {
                $r['Week Number'] = $r['Week Start Date'] = $r['Week End Date'] = $r['Year'] = $r['Month'] = $r['Day'] = null;
            }
            $r['AgeYearsNumeric'] = safeInt($r['CAgeYears'] ?? null, null);
        }
        unset($r);

        // Municipality statistics
        $munRef = loadMunRef();
        $munGroups = [];
        foreach ($records as $r) {
            if ($inclCause) {
                $hasCause = (trim((string)($r['CCauseImmediate'] ?? '')) !== '') || (trim((string)($r['CCauseAntecedent'] ?? '')) !== '') || 
                            (trim((string)($r['CCauseUnderlying'] ?? '')) !== '') || (trim((string)($r['CCauseOther'] ?? '')) !== '');
                if (!$hasCause) continue;
            }
            $munVal = (string)($r['CResidenceMunicipality'] ?? '');
            $code = 'UNKNOWN';
            if (strpos($munVal, '|') !== false) {
                $code = trim(substr($munVal, strrpos($munVal, '|') + 1));
                if ($code === '') $code = 'UNKNOWN';
            }
            $sex = strtoupper(trim((string)($r['CSexId'] ?? '')));
            if (!isset($munGroups[$code])) $munGroups[$code] = [0,0,0];
            if ($sex === 'MALE') $munGroups[$code][0]++;
            elseif ($sex === 'FEMALE') $munGroups[$code][1]++;
            $munGroups[$code][2]++;
        }
        
        uksort($munGroups, function($a,$b) { return ($a==='UNKNOWN'?'ZZZZZ':$a) <=> ($b==='UNKNOWN'?'ZZZZZ':$b); });
        
        $munStats = []; $sNo = 1;
        foreach ($munGroups as $code => list($male,$female,$total)) {
            if ($code === 'UNKNOWN') {
                $mun = 'Not Stated'; $prov = 'Not Stated'; $ctry = 'Philippines';
            } elseif (isset($munRef[$code])) { 
                $mun = $munRef[$code]['municipality']; $prov = $munRef[$code]['province']; $ctry = $munRef[$code]['country']; 
            } else {
                $mun = "Unknown ($code)"; $prov = 'Unknown'; $ctry = 'Philippines';
            }
            $munStats[] = ['no'=>$sNo++,'mun'=>$mun,'prov'=>$prov,'ctry'=>$ctry,'male'=>$male,'female'=>$female,'total'=>$total];
        }

        // Cause of death data
        $causeRows = [];
        if ($inclCause) {
            $cNo = 1;
            foreach ($records as $r) {
                $hasCause = (trim((string)($r['CCauseImmediate'] ?? '')) !== '') || (trim((string)($r['CCauseAntecedent'] ?? '')) !== '') || 
                            (trim((string)($r['CCauseUnderlying'] ?? '')) !== '') || (trim((string)($r['CCauseOther'] ?? '')) !== '');
                if (!$hasCause) continue;
                $dd = $r['CDeathDate'] ?? '';
                $ddf = ($dd !== '' && strtotime($dd) !== false) ? (new DateTime($dd))->format('Y-m-d') : '';
                $causeRows[] = ['no'=>$cNo++,'reg'=>$r['RegistryNum']??'','last'=>$r['CLastName']??'','first'=>$r['CFirstName']??'',
                               'mid'=>$r['CMiddleName']??'','dod'=>$ddf,'imm'=>$r['CCauseImmediate']??'','ant'=>$r['CCauseAntecedent']??'',
                               'und'=>$r['CCauseUnderlying']??'','undInt'=>$r['CCauseUnderlyingInterval']??'','other'=>$r['CCauseOther']??''];
            }
        }

        // Dead on arrival data
        $doaRows = []; $dNo = 1;
        foreach ($records as $r) {
            $af = strtoupper(trim((string)($r['AttendantAttendedFrom'] ?? '')));
            $isDOA = (strpos($af,'DEAD') === 0) || (strpos($af,'DOA') !== false) || (strpos($af,'ER DEATH') !== false);
            if (!$isDOA) continue;
            $dd = $r['CDeathDate'] ?? '';
            $dod = ($dd !== '' && strtotime($dd) !== false) ? (new DateTime($dd))->format('Y-m-d') : '';
            $doaRows[] = ['no'=>$dNo++,'reg'=>$r['RegistryNum']??'','last'=>$r['CLastName']??'',
                         'first'=>$r['CFirstName']??'','mid'=>$r['CMiddleName']??'','dod'=>$dod];
        }

        // Excel COM processing
        $excel = new COM("Excel.Application") or die("Unable to instantiate Excel");
        try {
            $excel->Visible = false;
            $excel->DisplayAlerts = false;
            $excel->ScreenUpdating = false;

            $workbook = $excel->Workbooks->Open(realpath($destPath));
            
            // Data Source sheet
            $dsSheet = $workbook->Sheets("Data Source");
            $cols = ['RegistryNum','DocumentStatus','CFirstName','CMiddleName','CLastName','CSexId','CDeathDate','CBirthDate',
                     'CAgeYears','CAgeMonths','CAgeDays','CAgeHours','CAgeMinutes','CDeathAddress','CDeathMunicipality',
                     'CDeathMunicipalityId','CDeathProvince','CDeathProvinceId','CDeathCountry','CDeathCountryId',
                     'CCivilStatusId','CReligion','CCitizenship','CResidenceAddress','CResidenceMunicipality',
                     'CResidenceMunicipalityId','CResidenceProvince','CResidenceProvinceId','CResidenceCountry',
                     'CResidenceCountryId','COccupation','FFirstName','FMiddleName','FLastName','MFirstName',
                     'MMiddleName','MLastName','CCauseImmediate','CCauseImmediateId','CCauseImmediateInterval',
                     'CCauseAntecedent','CCauseAntecedentId','CCauseAntecedentInterval','CCauseUnderlying',
                     'CCauseUnderlyingId','CCauseUnderlyingInterval','CCauseOther','CCauseOtherId','AttendantId',
                     'AttendantName','AttendantTitle','AttendantAttendedFrom','PreparerName','PreparerTitle',
                     'PreparerDate','DateReceived','DateRegistered','Week Number','Week Start Date','Week End Date',
                     'Year','Month','Day','AgeYearsNumeric'];

            $rn = 2;
            foreach ($records as $rec) {
                foreach ($cols as $ci => $colName) {
                    $val = $rec[$colName] ?? '';
                    $dsSheet->Cells($rn, $ci + 1)->Value = $val;
                }
                $rn++;
            }

            // ByMunicipality sheet
            $munSheet = $workbook->Sheets("ByMunicipality");
            if (!empty($munStats)) {
                $r = 14;
                foreach ($munStats as $m) {
                    $munSheet->Cells($r, 2)->Value = $m['no'];
                    $munSheet->Cells($r, 3)->Value = $m['mun'];
                    $munSheet->Cells($r, 4)->Value = $m['prov'];
                    $munSheet->Cells($r, 5)->Value = $m['ctry'];
                    $munSheet->Cells($r, 6)->Value = $m['male'];
                    $munSheet->Cells($r, 7)->Value = $m['female'];
                    $munSheet->Cells($r, 8)->Formula = "=F{$r}+G{$r}";
                    $r++;
                }
                $totR = $r;
                $lastDR = $r - 1;
                $munSheet->Range($munSheet->Cells($totR,2), $munSheet->Cells($totR,5))->Merge();
                $munSheet->Cells($totR, 2)->Value = 'TOTAL';
                $munSheet->Cells($totR, 6)->Formula = "=SUM(F14:F{$lastDR})";
                $munSheet->Cells($totR, 7)->Formula = "=SUM(G14:G{$lastDR})";
                $munSheet->Cells($totR, 8)->Formula = "=SUM(H14:H{$lastDR})";
            }

            // CauseOfDeath sheet
            if ($inclCause && !empty($causeRows)) {
                $causeSheet = $workbook->Sheets("CauseOfDeath");
                $cr = 7;
                foreach ($causeRows as $c) {
                    $causeSheet->Cells($cr, 1)->Value = $c['no'];
                    $causeSheet->Cells($cr, 2)->Value = $c['reg'];
                    $causeSheet->Cells($cr, 3)->Value = $c['last'];
                    $causeSheet->Cells($cr, 4)->Value = $c['first'];
                    $causeSheet->Cells($cr, 5)->Value = $c['mid'];
                    $causeSheet->Cells($cr, 6)->Value = $c['dod'];
                    $causeSheet->Cells($cr, 7)->Value = $c['imm'];
                    $causeSheet->Cells($cr, 8)->Value = $c['ant'];
                    $causeSheet->Cells($cr, 9)->Value = $c['und'];
                    $causeSheet->Cells($cr, 10)->Value = $c['undInt'];
                    $causeSheet->Cells($cr, 11)->Value = $c['other'];
                    $cr++;
                }
            }

            // DeadonArrival sheet
            if (!empty($doaRows)) {
                $doaSheet = $workbook->Sheets("DeadonArrival");
                $dr = 9;
                foreach ($doaRows as $d) {
                    $doaSheet->Cells($dr, 2)->Value = $d['no'];
                    $doaSheet->Cells($dr, 3)->Value = $d['reg'];
                    $doaSheet->Cells($dr, 4)->Value = $d['last'];
                    $doaSheet->Cells($dr, 5)->Value = $d['first'];
                    $doaSheet->Cells($dr, 6)->Value = $d['mid'];
                    $doaSheet->Cells($dr, 7)->Value = $d['dod'];
                    $dr++;
                }
            }

            $excel->Calculate();
            $workbook->Save();
            $workbook->Close();
            $excel->Quit();

          echo json_encode([
                'success' => true,
                'file' => $destPath
            ]);
            exit;



        } catch (Exception $ex) {
            throw new Exception('COM Error: ' . $ex->getMessage());
        } finally {
            if (isset($excel)) {
                $excel->Quit();
                unset($excel);
            }
        }

        header('Location: death_statistics_export.php');
        exit;

    } catch (Exception $ex) {
        $_SESSION['export_error'] = $ex->getMessage();
        header('Location: death_statistics_export.php');
        exit;
    }
}

// Form values
$currentYear = (int)date('Y');
$fYear = isset($_POST['year']) ? intval($_POST['year']) : $currentYear;
$fM1 = isset($_POST['month_start']) ? intval($_POST['month_start']) : 1;
$fM2 = isset($_POST['month_end']) ? intval($_POST['month_end']) : 12;
$fInclC = !empty($_POST['include_cause']);
$fC1 = $_POST['cause1'] ?? '';
$fC2 = $_POST['cause2'] ?? '';
$fC3 = $_POST['cause3'] ?? '';
$fPartner = !empty($_POST['source_partner']);
$fLGU = !empty($_POST['source_lgu']);
$fSave = $_POST['save_path'] ?? '';
$months = ['January','February','March','April','May','June','July','August','September','October','November','December'];
?>
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>Death Statistics Export</title>
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
.sbox.red{background:rgba(239,68,68,.08);border-color:#ef444466}
.sbox-title{font-size:.73rem;font-weight:700;text-transform:uppercase;color:#00d9ff;margin-bottom:9px;letter-spacing:.05em}
.sbox.red .sbox-title{color:#fca5a5}
.chk{display:flex;align-items:center;gap:10px;padding:5px 0}
.chk input[type=checkbox]{width:17px;height:17px;accent-color:#667eea;cursor:pointer;flex-shrink:0}
.chk label{color:#e2e8f0;font-size:.88rem;cursor:pointer;margin:0}
.chk small{color:#64748b;font-size:.73rem}
.cause-fields{display:none;margin-top:10px;border-top:1px solid #ef444433;padding-top:10px}
.cause-fields.show{display:block}
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
.modal-overlay{display:none;position:fixed;top:0;left:0;width:100%;height:100%;background:rgba(0,0,0,0.85);z-index:9999;align-items:center;justify-content:center}
.modal-overlay.show{display:flex}
.modal-content{background:#1e293b;border:2px solid #667eea;border-radius:12px;padding:40px;text-align:center;max-width:400px;width:90%}
.spinner{border:4px solid rgba(102,126,234,0.3);border-top:4px solid #667eea;border-radius:50%;width:60px;height:60px;animation:spin 1s linear infinite;margin:0 auto 20px}
@keyframes spin{0%{transform:rotate(0deg)}100%{transform:rotate(360deg)}}
.modal-content h3{font-size:1.2rem;margin-bottom:10px;color:#fff}
.modal-content p{color:#94a3b8;font-size:0.9rem;line-height:1.6}
</style>
</head>
<body>
<div class="modal-overlay" id="progressModal">
    <div class="modal-content">
        <div class="spinner"></div>
        <h3>Processing Export...</h3>
        <p>Please wait while we generate your report.<br>This may take 2-5 minutes for large datasets.<br><br><strong>Do not close this window.</strong></p>
    </div>
</div>

<div class="card">
    <div class="card-head">
        <i class="fas fa-file-excel"></i>
        <h3>Death Statistics Export</h3>
    </div>
    <div class="card-body">
        <?php if ($exportSuccess): ?>
        <div class="alert-ok">
            <i class="fas fa-check-circle mt-1"></i>
            <div>
                <strong>Export completed successfully!</strong><br>
                <span style="font-size:.8rem">Saved to: <code><?= htmlspecialchars($savedFile) ?></code></span></wordwrap>
            </div>
        </div>
        <?php elseif ($exportError): ?>
        <div class="alert-err">
            <i class="fas fa-exclamation-circle mt-1"></i>
            <span><?= htmlspecialchars($exportError) ?></span>
        </div>
        <?php endif; ?>
        
        <form id="frm" method="POST" action="death_statistics_export.php">
        <input type="hidden" name="action" value="export_death">
        
        <div class="frow">
            <label>Year</label><span>:</span>
            <select name="year" class="form-select" required>
                <?php for($y=$currentYear;$y>=1900;$y--): ?>
                <option value="<?=$y?>" <?=$y===$fYear?'selected':''?>><?=$y?></option>
                <?php endfor; ?>
            </select>
        </div>
        
        <div class="frow">
            <label>Month Start</label><span>:</span>
            <select name="month_start" class="form-select" required id="m1">
                <?php for($m=1;$m<=12;$m++): ?>
                <option value="<?=$m?>" <?=$m===$fM1?'selected':''?>><?=$m?> — <?=$months[$m-1]?></option>
                <?php endfor; ?>
            </select>
        </div>
        
        <div class="frow">
            <label>Month End</label><span>:</span>
            <select name="month_end" class="form-select" required id="m2">
                <?php for($m=1;$m<=12;$m++): ?>
                <option value="<?=$m?>" <?=$m===$fM2?'selected':''?>><?=$m?> — <?=$months[$m-1]?></option>
                <?php endfor; ?>
            </select>
        </div>
        
        <hr class="divider">
        
        <div class="sbox red">
            <div class="sbox-title"><i class="fas fa-heartbeat me-1"></i>Cause of Death</div>
            <div class="chk">
                <input type="checkbox" name="include_cause" id="inclCause" value="1" <?=$fInclC?'checked':''?>>
                <label for="inclCause"><strong>Include Cause of Death Filter</strong> <small>(at least one field required)</small></label>
            </div>
            <div class="cause-fields <?=$fInclC?'show':''?>" id="causeFields">
                <div class="frow" style="margin-top:8px">
                    <label>Search Cause 1</label><span>:</span>
                    <input type="text" name="cause1" class="form-control" placeholder="Searches all 4 cause fields..." value="<?=htmlspecialchars($fC1)?>">
                </div>
                <div class="frow">
                    <label>Search Cause 2</label><span>:</span>
                    <input type="text" name="cause2" class="form-control" placeholder="Searches all 4 cause fields..." value="<?=htmlspecialchars($fC2)?>">
                </div>
                <div class="frow">
                    <label>Search Cause 3</label><span>:</span>
                    <input type="text" name="cause3" class="form-control" placeholder="Searches all 4 cause fields..." value="<?=htmlspecialchars($fC3)?>">
                </div>
                <div class="hint">Each search term filters across CCauseImmediate, CCauseAntecedent, CCauseUnderlying, CCauseOther</div>
            </div>
        </div>
        
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
        
        <div class="sbox">
            <div class="sbox-title"><i class="fas fa-folder-open me-1"></i>Save Location</div>
            <input type="text" name="save_path" class="form-control" placeholder="e.g. C:\Reports or D:\" value="<?=htmlspecialchars($fSave)?>" required>
            <div class="hint">Enter a folder path (filename auto-generated) or full file path.</div>
        </div>
        
        <div class="info-box">
            <div class="info-title"><i class="fas fa-info-circle me-1"></i>Sheets Populated</div>
            <p>
                <i class="fas fa-table me-1" style="color:#52b788"></i><strong>Data Source</strong> — all records + computed columns<br>
                <i class="fas fa-map-marker-alt me-1" style="color:#f59e0b"></i><strong>ByMunicipality</strong> — Male/Female/Total (row 14↓)<br>
                <i class="fas fa-heartbeat me-1" style="color:#f472b6"></i><strong>CauseOfDeath</strong> — records with cause data (row 7↓)<br>
                <i class="fas fa-ambulance me-1" style="color:#ef4444"></i><strong>DeadonArrival</strong> — DOA records (row 9↓)
            </p>
        </div>
        
        <div class="btn-row">
            <button type="submit" class="btn-ok" id="btnGo">
                <i class="fas fa-file-export me-2"></i>Process & Export
            </button>
            <button type="button" class="btn-cancel" onclick="history.back()">
                <i class="fas fa-times me-2"></i>Cancel
            </button>
        </div>
        </form>
    </div>
</div>

<script>
document.getElementById('inclCause').addEventListener('change', function() {
    document.getElementById('causeFields').classList.toggle('show', this.checked);
});

document.getElementById('frm').addEventListener('submit', function(e) {
    e.preventDefault(); // prevent normal form submission

    const p = document.getElementById('sp').checked;
    const l = document.getElementById('sl').checked;
    const m1 = parseInt(document.getElementById('m1').value);
    const m2 = parseInt(document.getElementById('m2').value);
    const ic = document.getElementById('inclCause').checked;

    // Validation
    if (!p && !l) { alert('Please select at least one Source.'); return; }
    if (m2 < m1) { alert('Month End cannot be before Month Start.'); return; }
    if (ic) {
        const c1 = document.querySelector('[name=cause1]').value.trim();
        const c2 = document.querySelector('[name=cause2]').value.trim();
        const c3 = document.querySelector('[name=cause3]').value.trim();
        if (!c1 && !c2 && !c3) { alert('At least one Cause of Death field must be entered.'); return; }
    }

    // Show spinner modal
    const spinnerModal = document.getElementById('progressModal');
    spinnerModal.classList.add('show');

    const btnGo = document.getElementById('btnGo');
    btnGo.disabled = true;
    btnGo.innerHTML = '<i class="fas fa-spinner fa-spin me-2"></i>Processing...';

    // Send form via AJAX
    fetch('death_statistics_export.php', {
        method: 'POST',
        body: new FormData(this)
    })
    .then(res => res.json())
    .then(data => {
        spinnerModal.classList.remove('show');
        btnGo.disabled = false;
        btnGo.innerHTML = '<i class="fas fa-file-export me-2"></i>Process & Export';

        if (data.success) {
            // Show success modal with OK button
            showSuccessModal("Death Statistics Report Completed", data.file);
        } else {
            alert('Export failed: ' + (data.error || 'Unknown error'));
        }
    })
    .catch(err => {
        spinnerModal.classList.remove('show');
        btnGo.disabled = false;
        btnGo.innerHTML = '<i class="fas fa-file-export me-2"></i>Process & Export';
        alert('Export failed: ' + err.message);
    });
});

// Function to create success modal
function showSuccessModal(message, filePath = '') {
    const modal = document.createElement('div');
    modal.className = 'modal-overlay show';
    modal.innerHTML = `
        <div class="modal-content">
            <h3>${message}</h3>
            ${filePath ? `<p>Saved to: <code>${filePath}</code></p>` : ''}
            <button class="btn-ok" id="okBtn">OK</button>
        </div>
    `;
    document.body.appendChild(modal);

    document.getElementById('okBtn').addEventListener('click', function() {
        modal.remove();
        // Redirect to death_transmission.php after OK
        window.location.href = 'death_transmission.php';
    });
}
</script>
</body>
</html>