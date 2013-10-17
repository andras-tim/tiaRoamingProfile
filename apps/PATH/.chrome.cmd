@call createLinkedEmptyDirs chrome "%LOCALAPPDATA%\Google\CrashReports" "%LOCALAPPDATA%\Google\Chrome\User Data\Default\Cache" "%LOCALAPPDATA%\Google\Chrome\User Data\Default\GPUCache" "%LOCALAPPDATA%\Google\Chrome\User Data\Default\Media Cache"
@runFromPrgFiles "Google\Chrome\Application\chrome.exe" %*
