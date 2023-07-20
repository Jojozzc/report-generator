$path = "./文档生成器"
$templatePath = $path + '/template'
If(!(test-path -PathType container $path))
{
      New-Item -ItemType Directory -Path $path
}
If(!(test-path -PathType container $templatePath))
{
      New-Item -ItemType Directory -Path $templatePath
}
pyinstaller -F --noconsole app.py
Copy-Item "./dist/app.exe" -Destination $path
Copy-Item -Path "./template" -Destination $templatePath -Recurse


