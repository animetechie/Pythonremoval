# Get a list of all installed products
$products = Get-WmiObject -Class Win32_Product

# Filter the list of products to get only the Python installations
$pythonProducts = $products | Where-Object {$_.Name -like "*Python*"}

# Uninstall each Python installation
foreach ($product in $pythonProducts) {
  $product.Uninstall()
}

# Download the latest version of Python
$downloadUrl = "https://www.python.org/ftp/python/3.10.2/python-3.10.2-amd64.exe"
$installerPath = "$env:TEMP\python-3.10.2-amd64.exe"
Invoke-WebRequest -Uri $downloadUrl -OutFile $installerPath

# Extract the downloaded installer to a temporary directory
$extractPath = "$env:TEMP\python-3.10.2"
Expand-Archive -Path $installerPath -DestinationPath $extractPath

# Install Python silently
Start-Process -FilePath "$extractPath\python-3.10.2-amd64.exe" -ArgumentList '/quiet' -Wait

# Add Python to the PATH environment variable
$path = [Environment]::GetEnvironmentVariable("Path", "Machine")
$pythonPath = "$extractPath\python-3.10.2-amd64"
[Environment]::SetEnvironmentVariable("Path", "$path;$pythonPath", "Machine")
