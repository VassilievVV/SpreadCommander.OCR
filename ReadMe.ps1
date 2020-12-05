using namespace Tesseract;

$schost           = Get-SCHost;
$assemblyFileName = $schost.MapPath('~\bin\Tesseract.dll');
$tessDataDir      = $schost.MapPath('~\bin\tessdata');
$imageFileName    = $schost.MapPath('~\Images\phototest.tif');

Add-Type -Path:$assemblyFileName | Out-Null;

$schost.Silent = $true;

Clear-Book;
Clear-Spreadsheet;

Invoke-SCScript '~\InitBookStyles.ps1';

Set-BookSectionHeader '<b>Spread Commander</b> - <i>Examples: OCR</i>' -Html;
Set-BookSectionFooter 'Page {PAGE} of {NUMPAGES}' -ExpandFields;

Write-Text -ParagraphStyle:'Header1' 'OCR';

Write-Html -ParagraphStyle:'Description' @'
<p align=justify><b>PowerShell</b> script allows to reference .Net (Core) assemblies.
This example shows how to use .Net Core <b>Tesseract</b> engine
to OCR text. After executing script look also tab <b>Spreadsheet</b>
to review detail information.</p>
'@;

Write-Html -ParagraphStyle:'Description' @'
Download last <b>Tesseract</b> libraries from <a href="https://www.nuget.org/packages/Tesseract/">Nuget,
https://www.nuget.org/packages/Tesseract/</a> and unpack in folder <b>bin</b> of then <b>SpreadCommander</p> project.
'@;

Write-Text -ParagraphStyle:'Header2' 'Sample image';

Write-Image $imageFileName;

Write-Text -ParagraphStyle:'Header2' 'Recognized text';

$engine = [TesseractEngine]::new($tessDataDir, 'eng', [EngineMode]::Default);
$tblWords = [Data.DataTable]::new('Words');
try
{
    [void]$tblWords.Columns.Add("Word", [string]);
    [void]$tblWords.Columns.Add("Confidence", [float]);
    [void]$tblWords.Columns.Add("FontName", [string]);
    [void]$tblWords.Columns.Add("PointSize", [int]);
    [void]$tblWords.Columns.Add("IsSerif", [bool]);
    [void]$tblWords.Columns.Add("IsFixedPitch", [bool]);
    [void]$tblWords.Columns.Add("IsBold", [bool]);
    [void]$tblWords.Columns.Add("IsItalic", [bool]);
    [void]$tblWords.Columns.Add("IsUnderlined", [bool]);
    [void]$tblWords.Columns.Add("IsSmallCaps", [bool]);
    [void]$tblWords.Columns.Add("IsFromDictionary", [bool]);
    [void]$tblWords.Columns.Add("IsNumeric", [bool]);
    [void]$tblWords.Columns.Add("IsSubscript", [bool]);
    [void]$tblWords.Columns.Add("IsSuperscript", [bool]);
    [void]$tblWords.Columns.Add("Left", [int]);
    [void]$tblWords.Columns.Add("Top", [int]);
    [void]$tblWords.Columns.Add("Width", [int]);
    [void]$tblWords.Columns.Add("Height", [int]);

	$img = [Pix]::LoadFromFile($imageFileName);
	try
	{
		$page = [Page]$engine.Process($img);
		try
		{
			$page.GetText() | Write-Text;		
			Write-Text "Mean confidence: $($page.GetMeanConfidence())";
			
			$iter = [ResultIterator]$page.GetIterator();
			try
			{
				$iter.Begin();
				do
				{
					do
					{
						do
						{
							do
							{
								$bounds = [Rect]::new(0, 0, 0, 0);
								[void]$iter.TryGetBoundingBox([PageIteratorLevel]::Word, [ref]$bounds);

                                [void]$tblWords.Rows.Add(
                                    $iter.GetText([PageIteratorLevel]::Word),
                                    $iter.GetConfidence([PageIteratorLevel]::Word),
                                    $iter.GetWordFontAttributes().FontInfo.Name,
                                    $iter.GetWordFontAttributes().PointSize,
                                    $iter.GetWordFontAttributes().FontInfo.IsSerif,
                                    $iter.GetWordFontAttributes().FontInfo.IsFixedPitch,
                                    $iter.GetWordFontAttributes().FontInfo.IsBold,
                                    $iter.GetWordFontAttributes().FontInfo.IsItalic,
                                    $iter.GetWordFontAttributes().IsUnderlined,
                                    $iter.GetWordFontAttributes().IsSmallCaps,
                                    $iter.GetWordIsFromDictionary(),
                                    $iter.GetWordIsNumeric(),
                                    $iter.GetSymbolIsSubscript(),
                                    $iter.GetSymbolIsSuperscript(),
                                    $bounds.X1,
                                    $bounds.Y1,
                                    $bounds.Width,
                                    $bounds.Height);
							} while ($iter.Next([PageIteratorLevel]::TextLine, [PageIteratorLevel]::Word));
						} while ($iter.Next([PageIteratorLevel]::Para, [PageIteratorLevel]::TextLine));
					} while ($iter.Next([PageIteratorLevel]::Block, [PageIteratorLevel]::Para));
				} while ($iter.Next([PageIteratorLevel]::Block));
			}
			finally
			{
				$iter.Dispose();
			}
		}
		finally
		{
			$page.Dispose();
		}
	}
	finally
	{
		$img.Dispose();
	}
	
	$tblWords | Out-SpreadTable -TableName:'Words' -SheetName:'Words' `
		-TableStyle:Medium1 -Replace;
}
finally
{
	$tblWords.Dispose();
	$engine.Dispose();
}

Add-BookSection -ContinuePageNumbering -LinkHeaderToPrevious -LinkFooterToPrevious;
Write-Text -ParagraphStyle:'Header2' 'Table of Contents';
Add-BookTOC;

Save-Book '~\ReadMe.docx' -Replace;
Save-Spreadsheet '~\ReadMe.xlsx' -Replace;