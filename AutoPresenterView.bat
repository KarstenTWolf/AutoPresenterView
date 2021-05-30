for /f "delims=" %%a in ('dir /A-d /b /OD %1*.ppt*') do set "name=%%a"
start /B "" powerpnt /m AutoPresenterView.pptm HideMe
timeout 1 && start /B "" powerpnt /m "%name%" presenterview
exit