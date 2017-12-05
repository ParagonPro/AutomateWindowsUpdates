# AutomateWindowsUpdates
Script to run on a server to automate windows updates and document the updates that are installed

1. The scripts tells the server to check for any pending updates from our WSUS server
  a. If there aren't any then it closes out and does nothing else
2. It then collects all the details of the updates pending and sends a clean designed HTML email to an address we have setup so we have a document of what updates were installed on which server and when
3. Updates are then installed on the server
4. Once they are installed it checks if the server is pending a reboot
  a. If a reboot is pending it will email the department and let them know a reboot is pending and they can reboot it when they have time
  b. Otherwise, it closes out and updates are all done

To get this to work on your network you will need to edit lines 41 through 44 with the information that matches your setup

The legal stuff - I am not responsible for anything that may cause issues on your network, security issues, loss of data, or anything else that may occur from using this script. With that said, I have been using this script for 2 months now and have had no issues, but I am still not responsible for any issues that occur from this. Scripts are always a use at your own risk.
