# AutoPresenterView
Simple batch script work around to start a PPT in presenter view within windows. 

# SETUP: 
Drag/Copy both 'AutoPresenterView.bat' and 'AutoPresenter.pptm' into any directory you'll be using.

Within Powerpoint File > Options > Trust Center > Trust Center Settings > Macro Settings > Enable all macros (This is a security concern so only do this if you're confident in your security)

OR

The first time you trigger the .bat you'll need to enable macros within 'AutoPresenterView.pptm'. If you close completely out of powerpoint you'll need to do this step again. If you're triggering from multiple directories you'll need to do this for each directory.  

# How this works

Basically 'AutoPresenter.pptm' needs to stay open, that's where the macros live. The script searches the given directory for .ppt* with the leading search term, if two files start with the same term it'll open the most recent.

# How to Trigger

In your command line or VICREO Shell command input the following command. Where %PATHTOFILES% is the full path to your folder (ie C:\Users\Username\Desktop\Showfiles) and %SEARCH% is the begining of the file name (ie 'Fin' will open 'Final Deck.pptx'). If you're using VICREO Listener 2.1.0 and Companion 2.1.3, you'll need to 'reverse' the slashes or double them up to get it to work. (ie C:\\Users\\Username\\Desktop OR C:/Users/Username/Desktop)

cd %PATHTOFILES% && start AutoPresenterView.bat %SEARCH%
