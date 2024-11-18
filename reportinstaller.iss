[Setup]
AppName=ReportApp
AppVersion=1.0
DefaultDirName={autopf}\ReportApp
DefaultGroupName=ReportApp
OutputDir=.
OutputBaseFilename=MyAppSetup
AllowNoIcons=False

[Files]
; Include the main program file
Source: "C:\Users\Vagas\Documents\it report form\dist\report.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\Vagas\Documents\it report form\it_agent.xlsx"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\Vagas\Documents\it report form\config.json"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\Vagas\Documents\it report form\sign.ico"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{desktop}\ReportApp"; Filename: "{app}\ReportApp.exe"; WorkingDir: "{app}"; IconFilename: "{app}\ReportApp.ico"; IconIndex: 0

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
