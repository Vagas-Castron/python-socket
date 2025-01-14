[Setup]
AppName=ReportApp
AppVersion=v1.0
DefaultDirName={autopf}\ReportApp
DefaultGroupName=ReportApp
OutputDir=.
OutputBaseFilename=ITReportApp
AllowNoIcons=False

[Files]
; Include the main program file
Source: "C:\Users\Vagas\Desktop\reporting project\dist\report.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\Vagas\Desktop\reporting project\it_agent.xlsx"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\Vagas\Desktop\reporting project\config.json"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\Vagas\Desktop\reporting project\sign.ico"; DestDir: "{app}"; Flags: ignoreversion
//Source: "C:\Users\Vagas\Documents\it report form\app_folder_creation.py"; DestDir: "{app}"; Flags: ignoreversion
//Source: "C:/Users/Vagas/AppData/Local/Programs/Python/Python37/python.exe"; DestDir: "{app}"; Flags: ignoreversion

[Dirs]
; Create the folder in AppData\Local
Name: "{localappdata}\ReportApp"; Flags: uninsalwaysuninstall


;[Run]
; Run the Python script after installation
;Filename: "{app}\python.exe"; Parameters: """{app}\app_folder_creation.py"""; WorkingDir: "{app}"; Flags: waituntilterminated

[Icons]
Name: "{commondesktop}\ReportApp"; Filename: "{app}\report.exe"; WorkingDir: "{app}"; IconFilename: "{app}\sign.ico"; IconIndex: 0

[Code]
var
  CustomPage: TWizardPage;
  ServerRadio: TRadioButton;
  ClientRadio: TRadioButton;
  InstallMode: String;

procedure InitializeWizard();
begin
  // Create a custom page after the directory selection page
  CustomPage := CreateCustomPage(wpSelectDir, 'Installation Type', 'Choose how to install ReportApp');

  // Set page description and subtitle


  // Create and configure the "Server" radio button
  ServerRadio := TNewRadioButton.Create(WizardForm);
  ServerRadio.Parent := CustomPage.Surface;
  ServerRadio.Top := 16;
  ServerRadio.Left := 16;
  ServerRadio.Width := 400;
  ServerRadio.Caption := 'Install as Server';
  ServerRadio.Checked := True; // Default selection

  // Create and configure the "Client" radio button
  ClientRadio := TNewRadioButton.Create(WizardForm);
  ClientRadio.Parent := CustomPage.Surface;
  ClientRadio.Top := 40;
  ClientRadio.Left := 16;
  ClientRadio.Width := 400;
  ClientRadio.Caption := 'Install as Client';
end;

procedure CurStepChanged(CurStep: TSetupStep);
var
  ConfigFile: String;
  ConfigJSON: String;
begin
  // If at the post-installation step, save the selection to a config file
  if CurStep = ssPostInstall then
  begin
    // Determine selected mode
    if ServerRadio.Checked then
      InstallMode := 'server'
    else
      InstallMode := 'client';

    // Save the selection in a JSON file
    ConfigFile := ExpandConstant('{app}\config.json');
    ConfigJSON := '{ "mode": "' + InstallMode + '" }';
    SaveStringToFile(ConfigFile, ConfigJSON, False);
  end;
end;
