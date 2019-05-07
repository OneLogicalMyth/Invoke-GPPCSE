Function New-GPPCSEReport {
param($Data)

@"
<!DOCTYPE HTML>
<html>
<head>
	<style>
	body {
			font-family: "Helvetica Neue", Arial, sans-serif;
		}
	div {
			border: 2px solid;
			border-radius: 5px;
			width: 80%;
			margin:0 auto;
			padding: 10px;
		}
	h1 {
			width: 80%;
			margin:0 auto;
			padding: 10px;
		}
	table {
	  font-size: 14px;
	  border-collapse: collapse;
	}

	td, th {
	  padding: 10px;
	  text-align: left;
	  margin: 0;
	}

	tbody tr:nth-child(2n){
	  background-color: #eee;
	}

	th {
	  position: sticky;
	  top: 0;
	  background-color: #333;
	  color: white;
	}
</style>
</head>
<body>
	<h1>Invoke-GPPCSE Report</h1>
"@
foreach($Item in $Data)
{
@"
	<div>
		<h3>GPO Information</h3>
		<ul>
			<li>GPP Type: <b>$($Item.GPPType)</b></li>
			<li>GPP Username: <b>$($Item.GPPUsername)</b></li>
			<li>GPP New Username: <b>$($Item.GPPNewUsername)</b></li>
			<li>GPP Password: <b>$($Item.GPPPassword)</b></li>
			<li>GPO Path: <b>$($Item.GPOPath)</b></li>
			<li>GPO Name: <b>$($Item.GPOName)</b></li>
			<li>GPO GUID: <b>$($Item.GPOGUID)</b></li>
			<li>GPO Status: <b>$($Item.GPOStatus)</b></li>
			<li>GPO Linked OUs:
				<ul>
                    $($Item.GPOLinkedOUs | %{ "<li><b>$($_)</b></li>" })
				</ul>
			 </li>
			<li>GPO Associated OUs:
				<ul>
					$($Item.GPOAssociatedOUs | %{ "<li><b>$($_)</b></li>" })
				</ul>
			 </li>
		</ul>
      <h3>Related Active Computers</h3>
       $(if($Item.RelatedActiveComputers){ $Item.RelatedActiveComputers | ConvertTo-Html -Fragment }else{ '<i>None</i>' })

	</div>
<p>&nbsp;</p>
"@

}
@"
</body>
</html>
"@

}
