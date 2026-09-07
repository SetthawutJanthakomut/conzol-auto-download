ConZoL Auto Download - daily run on Windows
===========================================

What this does
--------------
Task Scheduler opens ConZoL in your "setthawut" Chrome profile every morning.
The userscript in that page then does its own once-a-day watch-list run, so the
new documents are on disk before you sit down.

It cannot work with Chrome closed - the script lives inside the ConZoL page -
which is why the task opens Chrome rather than downloading anything itself.

If the ConZoL session is still alive, the whole run happens with no clicks
at all. If it expired overnight, ConZoL shows its login page and the panel
puts a small bar in the corner: press "Sign in and start today's run" once
and everything continues by itself.

The password is never handled by this tool. Chrome fills it, and Chrome keeps
autofilled passwords hidden from page scripts until the page is clicked - so
a click has to happen, and clicking blindly would post a blank password and
could lock the account. One deliberate click a morning is the safe way round.

Setup
-----
1. Keep these files together in one folder.
2. In the ConZoL panel, tab "Daily":
      - put watchlist.txt in your destination folder
      - tick "Once a day, when the ConZoL page opens"
3. Double-click Install-Task.bat        -> creates the 08:00 daily task
   Double-click Run-Now.bat             -> try it straight away
   Double-click Uninstall-Task.bat      -> remove the task

No administrator rights are needed. The task runs as you and only while you are
logged on, because it has to open a window.

Changing things
---------------
Time         edit $RunAt in Install-Task.ps1, then run Install-Task.bat again
             (or change it in Task Scheduler)
Profile      edit $ProfileName in ConZoL-Daily.ps1
What is new  ConZoL-Daily.log, written next to these files, one line per run

If the profile name is not found the log says so and Chrome opens with its
Default profile instead.
