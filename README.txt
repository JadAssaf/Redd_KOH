Part 1 - MicroVisioneer Setup:
1. Install latest mvSlide version (2025.2.9179.31543 or newer)
- If no license is available, use the trial code MVRP-30TR for the activation of the software
2. In mvSlide, you first set up the objective presets as you're used to from manualWSI (see manual). 
https://get.mvslide.com/Microvisioneer_mvSlide_Manual_2024.pdf
3. Then, you need to switch the software into a mode where it asks for the slide name before you start the scan. In this mode, the software stores the scans in a preselected folder. 
- In "Adjust Scan Settings", click the settings icon (bottom left) which is advanced settings
- In "Automation & Power User Options" tab, turn on "Activate Scan Storage" and select "Manual (or from Barcode reader)"
- Click Save
4. Close the software.

Part 2 - KOH Smear Automation Setup:
1. Unzip the folder "Redd_KOH"
2. Place it under "C:\Users\Public\"
3. Press the Windows Key + R Key to open the run dialog window and paste `%programdata%\microvisioneer\mvslide` and press enter to open the settings folder of mvSlide. 
4. Open settings.mvsettings in notepad
5. Scroll down and
- Edit <RunExternalCommandAfterSavingScan> to `true`
- Edit <ExternalShellCommand> to `"C:\Users\Public\Redd_KOH\run_inference.bat"` (include the double quotations)
- Final lines should look like this (beware of quotations):

  <RunExternalCommandAfterSavingScan>true</RunExternalCommandAfterSavingScan>
  <RunExternalCommandAfterSavingScanFileExtension>.svs</RunExternalCommandAfterSavingScanFileExtension>
  <ExternalShellCommand>"C:\Users\Public\Redd_KOH\run_inference.bat"</ExternalShellCommand>
  <ExternalShellCommandArguments>"%1"</ExternalShellCommandArguments>
</MvSettings>

Part 3 - Using the Automation:
- The AI model will run automatically once you save an image on mvSlide software. Make sure you are using the correct magnification with the correct preset settings on the mvslide software. 20x and above are required for this algorithm to work.
- It will first ask you if you wish to run the analysis - press Yes.
- It will prompt a loading screen. Loading may take few minutes per slide, depending on your computer specifications and the slide size.
- A prediction will be made (fungus positive or not) with a probability score. 1.0 means 100% certainty by the AI model that the slide is fungus. The model can make mistakes. Please interpret its results accordingly.
- A heatmap image will appear which is an overlay on top of the slide. The greener the area, the more likely it is to contain a fungus filament.

