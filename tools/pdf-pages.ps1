# Rasterise the pages of a PDF to PNG, with nothing installed: Windows'
# own PDF renderer (Windows.Data.Pdf, the one Edge and the Photos app use).
#
#   powershell -File tools/pdf-pages.ps1 -Pdf "a.pdf" -Out "D:\compare\a" [-Pages 3] [-Width 1000]
#
# One PNG per page, page-1.png upward, each scaled to the width asked for.
# Used to put Microsoft Word's export of a document beside ours, page by page:
# the same document, the same page, two renderers.

param(
  [Parameter(Mandatory = $true)] [string] $Pdf,
  [Parameter(Mandatory = $true)] [string] $Out,
  [int] $Pages = 3,
  [int] $Width = 1000
)

Add-Type -AssemblyName System.Runtime.WindowsRuntime
$null = [Windows.Data.Pdf.PdfDocument, Windows.Data.Pdf, ContentType = WindowsRuntime]
$null = [Windows.Storage.StorageFile, Windows.Storage, ContentType = WindowsRuntime]
$null = [Windows.Storage.Streams.RandomAccessStream, Windows.Storage.Streams, ContentType = WindowsRuntime]

# WinRT async → .NET Task, the way PowerShell 5.1 has to do it.
$asTask = [System.WindowsRuntimeSystemExtensions].GetMethods() |
  Where-Object { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' } |
  Select-Object -First 1
function Await($op, [Type] $resultType) {
  $task = $asTask.MakeGenericMethod($resultType).Invoke($null, @($op))
  $task.Wait(-1) | Out-Null
  return $task.Result
}
$asTaskVoid = [System.WindowsRuntimeSystemExtensions].GetMethods() |
  Where-Object { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncAction' } |
  Select-Object -First 1
function AwaitAction($op) {
  $task = $asTaskVoid.Invoke($null, @($op))
  $task.Wait(-1) | Out-Null
}

New-Item -ItemType Directory -Force -Path $Out | Out-Null
$full = (Resolve-Path -LiteralPath $Pdf).Path
$file = Await ([Windows.Storage.StorageFile]::GetFileFromPathAsync($full)) ([Windows.Storage.StorageFile])
$doc = Await ([Windows.Data.Pdf.PdfDocument]::LoadFromFileAsync($file)) ([Windows.Data.Pdf.PdfDocument])
$count = [Math]::Min($Pages, $doc.PageCount)
Write-Output ("{0} pages in {1}; rendering {2}" -f $doc.PageCount, [System.IO.Path]::GetFileName($full), $count)

for ($i = 0; $i -lt $count; $i++) {
  $page = $doc.GetPage($i)
  $options = New-Object Windows.Data.Pdf.PdfPageRenderOptions
  $scale = $Width / $page.Size.Width
  $options.DestinationWidth = [uint32] $Width
  $options.DestinationHeight = [uint32] [Math]::Round($page.Size.Height * $scale)
  $stream = New-Object Windows.Storage.Streams.InMemoryRandomAccessStream
  AwaitAction ($page.RenderToStreamAsync($stream, $options))
  $stream.Seek(0)
  $reader = New-Object Windows.Storage.Streams.DataReader($stream.GetInputStreamAt(0))
  $size = [uint32] $stream.Size
  Await ($reader.LoadAsync($size)) ([uint32]) | Out-Null
  $bytes = New-Object byte[] $size
  $reader.ReadBytes($bytes)
  $png = Join-Path $Out ("page-{0}.png" -f ($i + 1))
  [System.IO.File]::WriteAllBytes($png, $bytes)
  Write-Output ("  page {0}: {1}x{2} -> {3}" -f ($i + 1), $options.DestinationWidth, $options.DestinationHeight, $png)
  $page.Dispose()
  $stream.Dispose()
}
