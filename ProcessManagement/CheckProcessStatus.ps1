# This script checks the status of Title Page generation.
# If the word application is not responding then it kills the word processes and restarts the filewatcher applciation.
# Some debug output is foreseen if the parameter "DEBUG" is passed as an argument.
# In fututre the application name and the paths can be passed as arguments.

# Generate log as status.txt in case of restart.
$debug_argument = $args[0]

$logfile_path = "C:\XXXX\status.txt"

$filewatcher_path = "C:\Program Files (x86)\XXXXX\"
$filewatcher_executable = "WFAXXXXX"
$word_executable = "WINWORD"

$process_count = (Get-Process | ?{$_.ProcessName -eq $word_executable} | ?{$_.Responding -like "False"}).count
if($process_count -gt 0)
{

   try 
   {
      # Generate log as status.txt in case of restart.
      $log_line = Get-Date -Format "dd/MM/yyyy HH:mm:ss"
      $log_line += "`t TitlePageStatus"
      $log_line += "`t Word is not responding. Killing all frozen word applications and restarting FileWatcher. Title Page generation will resume after restart."
      $log_line | Out-File $logfile_path -Append

      # Stopping dotnet application before killing word to avoid deletion of unprocessed input file.
      # Stopping only those word processes that are not responding.
      Get-Process | ?{$_.ProcessName -eq $filewatcher_executable} | Stop-Process
      Get-Process | ?{$_.ProcessName -eq $word_executable} | ?{$_.Responding -like "False"}| Stop-Process
      $filewatcher_path += $filewatcher_executable + ".exe"
      & $filewatcher_path 
   }
   
   catch
   {
      $log_line = Get-Date -Format "dd/MM/yyyy HH:mm:ss"
      $log_line += "`t TitlePageStatus"
      $log_line += "`t Failed to Kill frozen applications or restart FileWatcher. Manual intervention needed."
      $log_line += "`t Unable to resolve Title Generation."
      $log_line | Out-File $logfile_path -Append
   }
      
   
}
elseif($debug_argument -eq "DEBUG")
{
   # Debug output.
   $log_line = Get-Date -Format "dd/MM/yyyy HH:mm:ss"
   $log_line += "`t XXXXStatus"
   $log_line += "`t No issues found."
   $log_line | Out-File $logfile_path -Append
}