Tenjin-Sync
===========

Tenjin Sync is a quick and dirty XML worker to allow Tenjin to grab OneNote based homework. Tenjin Sync must first be compiled with the appropriate, roommate specific values incorperated into it (near the top of the main file). After compiling, deploy it to a dedicated user account (i.e an AD based user in Windows Server). Create a notebook in OneNote named 'Tenjin' with a section named 'Homework'. Add two pages, each with the first name of the respective roommate. On each page, add a task list for managing homework. Share the notebook between both roommates. Ensure that OneNote is installed on the server and logged into the notebook owner.

Despite the convoluted process and the poorly build sync 'agent', your OneNote homework will successfully sync into Tenjin's web UI. Maybe at somepoint, I'll rebuild Tenjin Sync to behave like a 'proper' piece of software. But for now, there are bigger things to do....
