

# Import email settings from config file
[xml]$ConfigFile = Get-Content ".\Settings.xml"

add-type -Path $ConfigFile.Settings.Configuration.SQLiteLibLoc

# Init vars
$outcsv = $ConfigFile.Settings.Configuration.outputCSVFileLoc

$sqlitedbsrc = $ConfigFile.Settings.Configuration.sqlitedbsrc
$sqlitedbcopy = $ConfigFile.Settings.Configuration.sqlitedbcopy

# Get copy of GSAK database to avoid locking issues
Write-Information -MessageData "Copying database..."
Copy-Item -Path $sqlitedbsrc -Destination $sqlitedbcopy

$connectionString = "Data Source=" + $sqlitedbcopy + "\sqlite.db3"
function DatabaseConn ($connectionString) {

    return $connection
}

function QuerySQLite ($query) {

    $con = New-Object -TypeName System.Data.SQLite.SQLiteConnection
    $con.ConnectionString = $connectionString

    $con.Open()
    $sql = $con.CreateCommand()
    $sql.CommandText = $query
    $adapter = New-Object -TypeName System.Data.SQLite.SQLiteDataAdapter $sql
    $data = New-Object System.Data.DataSet
    [void]$adapter.Fill($data)

    $sql.Dispose()
    $con.close()

    return $data
}

function SelectSQLite ($query, $connectionString) {

    $con = New-Object -TypeName System.Data.SQLite.SQLiteConnection
    $con.ConnectionString = $connectionString

    $con.Open()
    $sql = $con.CreateCommand()
    $sql.CommandText = $query
    $adapter = New-Object -TypeName System.Data.SQLite.SQLiteDataAdapter $sql
    $data = New-Object System.Data.DataSet
    [void]$adapter.Fill($data)

    $sql.Dispose()
    $con.close()

    return $data
}

# Dataprep

$sqlDataPrep = @'
-- Lag tabell poeng
drop table if exists poeng;
create table poeng(lBy varchar(50), publ_selv int, c_publ_selv varchar(200), ftf int, c_ftf varchar(200), delt_ftf int, c_delt_ftf varchar(200), funn_publ_dato int, c_funn_publ_dato varchar(200), funn_desember int, c_funn_desember varchar(200), total int, rank int);

--Lag tabell ftf
drop table if exists ftf; 
create table ftf (lBy varchar(50), code varchar(50), points int);

--Legg inn alle logger som inneholder {*FTF*} i ftf tabellen
insert into ftf select l.lBy, l.lParent, 0 from Logs l inner join logmemo lm on l.lLogId = lm.lLogId where lText like "%{*FTF*}%";

--Legg inn alle FTF'er som ikke er brukt rett kode på.. >:-O
--#2
--insert into ftf (lby, code, points) values ('O-K Haukland', 'GC6WRPD', 0);
--insert into ftf (lby, code, points) values ('annesto', 'GC6WRPD', 0);
--insert into ftf (lby, code, points) values ('dogteam', 'GC6WRPD', 0);
--#4
--insert into ftf (lby, code, points) values ('dogteam', 'GC6WYJF', 0);
--#17
--insert into ftf (lby, code, points) values ('Team Skartun', 'GC6X841', 0);
--#18
--insert into ftf (lby, code, points) values ('skogmal', 'GC6XVEX', 0);
--#23
--insert into ftf (lby, code, points) values ('sømna', 'GC6XVD3', 0);



-- Gi poeng til FTF og Co-FTF
update ftf set points = 2 where code in (select code from ftf group by code having count(code) > 1);
update ftf set points = 3 where code in (select code from ftf group by code having count(code) = 1);

--Legg inn alle som har logget i poengtabellen
insert into poeng select lBy, null, null, null, null, null, null, null, null, null, null, 0, 0 from logs where lType IN ("Found it", "Attended") and lBy NOT IN ("Makuta Teridax", "Poltrona Polaris", "Hexa Nomos", "NonaNorwegianAdiutor", "Octa Ceres", "cervisvenator") group by lBy order by lBy;

--Legg inn alle som har lagt ut cacher i poengtabellen
insert into poeng select placedby, null, null, null, null, null, null, null, null, null, null, 0, 0 from caches where placedby not in (select lBy from poeng);

--Sett korrekt placeddate og lag nytt smartname på cachene
update caches set placeddate = "2017-12-01", User4 = "[#1]" where code = "GC7E58G";
update caches set placeddate = "2017-12-02", User4 = "[#2]" where code = "GC7ER9N";
update caches set placeddate = "2017-12-03", User4 = "[#3]" where code = "GC7EKM1";
update caches set placeddate = "2017-12-04", User4 = "[#4]" where code = "GC7FD4G";
update caches set placeddate = "2017-12-05", User4 = "[#5]" where code = "GC7FD5P";
update caches set placeddate = "2017-12-06", User4 = "[#6]" where code = "GC7EPYZ";
update caches set placeddate = "2017-12-07", User4 = "[#7]" where code = "";
update caches set placeddate = "2017-12-08", User4 = "[#8]" where code = "GC7F114";
update caches set placeddate = "2017-12-09", User4 = "[#9]" where code = "GC7F6YH";
update caches set placeddate = "2017-12-10", User4 = "[#10]" where code = "GC7ERHH";
update caches set placeddate = "2017-12-11", User4 = "[#11]" where code = "GC7EZX3";
update caches set placeddate = "2017-12-12", User4 = "[#12]" where code = "";
update caches set placeddate = "2017-12-13", User4 = "[#13]" where code = "GC7EE2E";
update caches set placeddate = "2017-12-14", User4 = "[#14]" where code = "GC7F2DB";
update caches set placeddate = "2017-12-15", User4 = "[#15]" where code = "";
update caches set placeddate = "2017-12-16", User4 = "[#16]" where code = "";
update caches set placeddate = "2017-12-17", User4 = "[#17]" where code = "";
update caches set placeddate = "2017-12-18", User4 = "[#18]" where code = "";
update caches set placeddate = "2017-12-19", User4 = "[#19]" where code = "";
update caches set placeddate = "2017-12-20", User4 = "[#20]" where code = "";
update caches set placeddate = "2017-12-21", User4 = "[#21]" where code = "";
update caches set placeddate = "2017-12-22", User4 = "[#22]" where code = "";
update caches set placeddate = "2017-12-23", User4 = "[#23]" where code = "";
update caches set placeddate = "2017-12-24", User4 = "[#24]" where code = "";

--Andre manuelle triks
--Noen logger må flyttes
--#1
--update logs set ldate = "2016-12-01" where lparent = "GC6WN01" and lby IN ("dogteam", "cara2006", "O-K Haukland") and lType="Found it";
--#3
--update logs set ldate = "2016-12-03" where lparent = "GC6WVNR" and lby IN ("dogteam", "cara2006") and lType="Found it";
--#4
--update logs set ldate = "2016-12-04" where lparent = "GC6WYJF" and lby IN ("cara2006") and lType="Found it";

--#5
--update logs set ldate = "2016-12-05" where lparent = "GC6WQ2M" and lby IN ("dogteam", "cara2006") and lType="Found it";
--#6
--update logs set ldate = "2016-12-06" where lparent = "GC6WY6P" and lby IN ("dogteam") and lType="Found it";
--#7
--update logs set ldate = "2016-12-07" where lparent = "GC6WY6Z" and lby IN ("dogteam") and lType="Found it";
--#9
--update logs set ldate = "2016-12-09" where lparent = "GC6X2AN" and lby IN ("dogteam", "minni09", "mirg75", "silyam") and lType="Found it";
--#10
--update logs set ldate = "2016-12-10" where lparent = "GC6WW07" and lby IN ("silyam", "dogteam") and lType="Found it";

--#13 Luciaeventet var alle på samme dag
--update logs set ldate = "2016-12-13" where lparent = "GC6X53M" and lType="Attended";
--#14
--update logs set ldate = "2016-12-14" where lparent = "GC6X33Q" and lby IN ("O-K Haukland") and lType="Found it";
--#15
--update logs set ldate = "2016-12-15" where lparent = "GC6X260" and lby IN ("dogteam", "cara2006") and lType="Found it";
--#16
--update logs set ldate = "2016-12-15" where lparent = "GC6X7RN" and lby IN ("cara2006") and lType="Found it";

'@

QuerySQLite -query $sqlDataPrep | Out-Null


# Her begynner vi med utgangspunkt i alle deltagere = alle som har logget hittil i Desember
$sqlAlleDeltagere = "select lBy from poeng;"
$resAlleDeltagere = QuerySQLite -query $sqlAlleDeltagere

foreach ($deltager in $resAlleDeltagere.tables.rows) {

# Finn cacher Deltager har logget i Desember, legg til User4 for disse i poeng, og regn ut antall poeng
    $sqlDeltagerHarLogget = "select c.User4 from logs l inner join caches c on l.lParent = c.code where l.lType='Found it' and l.lBy = '" + $deltager.lBy + "' and lDate between '2017-12-01' AND '2017-12-31' and lDate != c.placeddate order by l.lDate;"
    $resDeltagerHarLogget = QuerySQLite -query $sqlDeltagerHarLogget

    foreach ($c in $resDeltagerHarLogget.tables.rows) {
        $caches += $c.User4 + " "
        $points++
    }
    $queryresult = "update poeng set funn_desember = " + $points + ", c_funn_desember = '" + $caches + " (" + $points + ")', total = total+"+ $points +" where lBy = '" + $deltager.lBy + "';"
    if($points -gt 0) {
        QuerySQLite -query $queryresult | Out-Null
    }
    $caches = ""
    Clear-Variable points
    
# Finn cacher Deltager har logget på publiseringsdato, legg til User4 og antall poeng
    $sqlLoggetPublDato= 'select c.User4 from logs l inner join caches c on l.lParent = c.code where l.lType in ("Attended", "Found it") and l.lBy = "' + $deltager.lBy + '" and l.lDate = c.placeddate order by l.lDate;'
    $resLoggetPublDato = QuerySQLite -query $sqlLoggetPublDato

    foreach ($c in $resLoggetPublDato.tables.rows) {
        $caches += $c.User4 + " "
        $points = $points +2
    }
    $queryresult = "update poeng set funn_publ_dato=" + $points +", c_funn_publ_dato='" + $caches + " (" + $points + ")', total=total+" + $points +" where lBy='" + $deltager.lBy + "';"
    if ($points -gt 0) {
        QuerySQLite -query $queryresult | Out-Null
    }
    $caches = ""
    Clear-Variable points   

# Finn cacher Deltager har publisert selv, legg til User4 og antall poeng
    $sqlPublSelv= "select c.User4 from poeng p inner join caches c on p.lBy = c.placedBy where c.placedBy = '" + $deltager.lBy + "';"
    $resPublSelv = QuerySQLite -query $sqlPublSelv

    foreach ($c in $resPublSelv.tables.rows) {
        $caches += $c.User4 + " "
        $points = $points + 3
    }

    $queryresult = "update poeng set publ_selv="+ $points +", c_publ_selv='" + $caches + " (" + $points + ")', total=total+" + $points + " where lBy='" + $deltager.lBy + "';"
    if($points -gt 0) {
        QuerySQLite -query $queryresult | Out-Null
    }
    $caches = ""
    Clear-Variable points

# Finn cacher Deltager har logget FTF på, og legg til * og redusert antall poeng på co-FTF

    $sqlCoFTF = "select * from ftf f inner join caches c on f.code = c.code where f.points = 2 and f.lBy = '" + $deltager.lBy + "';"
    $resCoFTF = QuerySQLite -query $sqlCoFTF

    foreach ($c in $resCoFTF.tables.rows) {
        $caches += $c.User4 + "* "
        $points += $c.points
    }

    $sqlCoFTF = "select * from ftf f inner join caches c on f.code = c.code where f.points = 3 and f.lBy = '" + $deltager.lBy + "';"
    $resCoFTF = QuerySQLite -query $sqlCoFTF

    foreach ($c in $resCoFTF.tables.rows) {
        $caches += $c.User4 + " "
        $points += $c.points;
    }

    $queryresult = "update poeng set ftf = " + $points + ", c_ftf = '" + $caches + " (" + $points + ")', total = total + " + $points + " where lBy='" + $deltager.lBy + "';"
    if($points -gt 0) {
        QuerySQLite -query $queryresult | Out-Null
    }
    $caches = ""
    if ($points) {
        Clear-Variable points  
    }

}

# TODO Beregn ranking og legg inn i poengbasen

$sqlPoeng = "select total from poeng order by total desc;"
$resPoeng = QuerySQLite -query $sqlPoeng
$tempPoeng = 0
$thisRank = 1

foreach ($c in $resPoeng.tables.rows) {
    
    if ($tempPoeng -gt $c.total) {
        $thisRank++
    }

    $sqlUpdateRank = "update poeng set rank = " + $thisRank + " where total = " + $c.total + ";"
    $resUpdateRank = QuerySQLite -query $sqlUpdateRank - Out-Null
    $tempPoeng = $c.total
}

$poeng = QuerySQLite -query "select rank as [Plass], lBy as [Nick], c_publ_selv as [Publisert selv], c_ftf as [FTF], c_funn_publ_dato as [Funn på publ dato], c_funn_desember as[Funn i Desember], total as [Total] from poeng order by total desc;"
$poeng.tables.rows | Out-GridView

if (Test-Path $outcsv) {Remove-Item $outcsv}
$poeng.tables.rows | Export-csv -LiteralPath $outcsv -NoTypeInformation -NoClobber -Encoding Unicode
$totalLogs = QuerySQLite -query "select count(*) AS [Ant logger] from logs where ltype in ('Found it', 'Attended')"
$totalLogs.tables.rows
$totalTeams = QuerySQLite -query "select count(*) AS [Ant team] from poeng"
$totalTeams.tables
# Lage HTML side av dataene
#$poeng.tables.rows

