on run
  set launcherPath to "/Users/hfraser/Documents/New project/Run Preppy.command"

  tell application "Terminal"
    activate
    do script ("zsh " & quoted form of launcherPath)
  end tell
end run
