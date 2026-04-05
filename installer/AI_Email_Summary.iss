#define AppName "AI Email Summary"
#define AppVersion "0.2.0"
#define AppPublisher "AI Email Summary"
#define AppExeName "AI_Email_Summary.exe"

[Setup]
AppId={{C2558B3D-8A15-4B1A-AE94-6A4B07E69B69}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher={#AppPublisher}
DefaultDirName={autopf}\AI Email Summary
DefaultGroupName={#AppName}
DisableProgramGroupPage=yes
OutputDir=..\dist-installer
OutputBaseFilename=AI_Email_Summary_Setup
Compression=lzma
SolidCompression=yes
WizardStyle=modern
ArchitecturesInstallIn64BitMode=x64compatible

[Files]
Source: "..\dist\AI_Email_Summary.exe"; DestDir: "{app}"; DestName: "{#AppExeName}"; Flags: ignoreversion

[Icons]
Name: "{autoprograms}\{#AppName}"; Filename: "{app}\{#AppExeName}"
Name: "{autodesktop}\{#AppName}"; Filename: "{app}\{#AppExeName}"

[Run]
Filename: "{app}\{#AppExeName}"; Description: "Launch AI Email Summary"; Flags: nowait postinstall skipifsilent

[Code]
var
  ProviderPage: TInputOptionWizardPage;
  KeysPage: TWizardPage;
  OpenAIEdit: TNewEdit;
  AnthropicEdit: TNewEdit;
  OpenAIModelEdit: TNewEdit;
  AnthropicModelEdit: TNewEdit;

procedure AddLabeledEdit(Parent: TWizardPage; CaptionText: String; TopPos: Integer; var Edit: TNewEdit; DefaultValue: String);
var
  LabelCtl: TNewStaticText;
begin
  LabelCtl := TNewStaticText.Create(Parent);
  LabelCtl.Parent := Parent.Surface;
  LabelCtl.Caption := CaptionText;
  LabelCtl.Left := ScaleX(0);
  LabelCtl.Top := TopPos;
  LabelCtl.Width := Parent.SurfaceWidth;

  Edit := TNewEdit.Create(Parent);
  Edit.Parent := Parent.Surface;
  Edit.Left := ScaleX(0);
  Edit.Top := TopPos + ScaleY(16);
  Edit.Width := Parent.SurfaceWidth;
  Edit.Text := DefaultValue;
end;

procedure InitializeWizard;
begin
  ProviderPage := CreateInputOptionPage(
    wpSelectDir,
    'Default AI Provider',
    'Choose the provider the app should default to',
    'Users can still switch providers later inside the app.',
    True,
    False
  );
  ProviderPage.Add('Anthropic Sonnet');
  ProviderPage.Add('OpenAI');
  ProviderPage.Values[0] := True;

  KeysPage := CreateCustomPage(
    ProviderPage.ID,
    'API Keys',
    'Enter one or both API keys. The installer will write them into the .env file.'
  );

  AddLabeledEdit(KeysPage, 'OpenAI API key', ScaleY(0), OpenAIEdit, '');
  AddLabeledEdit(KeysPage, 'OpenAI model', ScaleY(52), OpenAIModelEdit, 'gpt-4.1');
  AddLabeledEdit(KeysPage, 'Anthropic API key', ScaleY(104), AnthropicEdit, '');
  AddLabeledEdit(KeysPage, 'Anthropic model', ScaleY(156), AnthropicModelEdit, 'claude-sonnet-4-20250514');
end;

function NextButtonClick(CurPageID: Integer): Boolean;
begin
  Result := True;
  if CurPageID = KeysPage.ID then
  begin
    if (Trim(OpenAIEdit.Text) = '') and (Trim(AnthropicEdit.Text) = '') then
    begin
      MsgBox('Enter at least one API key before continuing.', mbError, MB_OK);
      Result := False;
    end;
  end;
end;

procedure CurStepChanged(CurStep: TSetupStep);
var
  EnvText: AnsiString;
  DefaultProvider: String;
begin
  if CurStep = ssPostInstall then
  begin
    if ProviderPage.Values[0] then
      DefaultProvider := 'anthropic'
    else
      DefaultProvider := 'openai';

    EnvText :=
      'DEFAULT_AI_PROVIDER=' + DefaultProvider + #13#10 +
      'OPENAI_API_KEY=' + OpenAIEdit.Text + #13#10 +
      'OPENAI_MODEL=' + OpenAIModelEdit.Text + #13#10 +
      'ANTHROPIC_API_KEY=' + AnthropicEdit.Text + #13#10 +
      'ANTHROPIC_MODEL=' + AnthropicModelEdit.Text + #13#10;

    SaveStringToFile(ExpandConstant('{app}\.env'), EnvText, False);
  end;
end;
