# Loops through worltrees and finds the content of gitdir file.
# Might be useful if the local repositiry folder has been moved.

$worktree_location = "last\.git\worktrees\"

$directories = dir | % { if ($_.PsIsContainer) { $_.FullName + "\" } }
foreach ($directory in $directories)
{
	$d = -join($directory, $worktree_location)
	$git_worktrees = dir $d | % { if ($_.PsIsContainer) { $_.FullName + "\" } }
	foreach ($git_worktree in $git_worktrees)
	{
		$workdir_file = -join($git_worktree, "gitdir")
		get-content $workdir_file
	}
}
