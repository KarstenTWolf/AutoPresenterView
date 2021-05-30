# This File is for documentation, it has no function

# HideMe

Sub HideMe()
    Application.WindowState = ppWindowMinimized
End Sub


# PresenterView

Sub PresenterView()
    ActivePresentation.SlideShowSettings.ShowType = ppShowTypeSpeaker
    ActivePresentation.SlideShowSettings.ShowPresenterView = True
    ActivePresentation.SlideShowSettings.Run.View.AcceleratorsEnabled = False
    Application.WindowState = ppWindowMinimized
End Sub
