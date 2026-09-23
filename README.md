# Yangi alifbo: Uzbek new alphabet for Microsoft Word

A Word add-in that writes Uzbek text in the new Latin alphabet (Ö, Ğ, Ş, Ç). It converts both old Latin text
(Oʻ, Gʻ, Sh, Ch) and Uzbek Cyrillic. Before anything changes it highlights every spot it will convert, so you can
check them and skip the ones you want to keep.

```
Oʻzbekiston Respublikasi poytaxti Toshkent shahri  →  Özbekiston Respublikasi poytaxti Toşkent şahri
Ўзбекистон Республикаси пойтахти Тошкент шаҳри      →  Özbekiston Respublikasi poytaxti Toşkent şahri
```

## What it converts

### Old Latin

| Old         | New |
|-------------|-----|
| Oʻ oʻ       | Ö ö |
| Gʻ gʻ       | Ğ ğ |
| Sh sh SH    | Ş ş |
| Ch ch CH    | Ç ç |

Every character people type for the ʻ counts: ʻ (the official one), `'`, `` ` ``, ’, ‘, ʼ and ´.
The tutuq belgisi in a word like Isʼhoq (s, ʼ, h) is not a letter pair and stays.

### Uzbek Cyrillic

The rules follow the [lotin-kirill](https://github.com/diyorbek/lotin-kirill) transliterator, with Ш, Ч, Ў and Ғ
written as the new letters.

| Cyrillic      | New alphabet | Example |
|---------------|--------------|---------|
| Ш Ч Ў Ғ       | Ş Ç Ö Ğ      | Тошкент → Toşkent, ғишт → ğişt |
| Ҳ Х Қ Й Ж     | H X Q Y J    | шаҳар → şahar |
| Ё Ю Я         | Yo Yu Ya     | Ёш → Yoş, ЁШ → YOŞ |
| Е             | Ye at the start of a word, after a vowel and after Ъ or Ь; otherwise e | ер → yer, Мирзиёев → Mirziyoyev, кеча → keça |
| Ц             | ts after a vowel, otherwise s | лицей → litsey, цирк → sirk |
| Ъ             | ʼ, dropped after Ў or Ғ and before Е | маъно → maʼno, мўъжиза → möjiza, съезд → syezd |
| Ь             | dropped; ьо becomes yo | медальон → medalyon |
| Э             | e            | эшик → eşik |
| СҲ            | sʼh, so it is not read as ş | Исҳоқ → Isʼhoq |

A few words the rules would get wrong are built in: бюджет → budjet, октябрь → oktabr, сентябрь → sentabr.
They also match longer forms such as бюджетга and октябрда.

Words with letters Uzbek does not use are left alone. This covers the Russian Щ and Ы and Ukrainian or Kyrgyz letters.

## Installing

Download and run [setup.exe](https://discoveruz.github.io/NewAlphabetPlugin/setup.exe). It adds the Visual Studio Tools
for Office runtime if that is missing, then installs the add-in for your Windows user. Windows asks whether to trust the
publisher, because the add-in is signed with a certificate made for the project rather than a bought one.

Every time Word starts, the add-in looks at the same address for a newer version and installs it by itself. To remove
it, use Windows Settings → Apps → Installed apps → NewAlphabetPlugin.

## Using it

Word gets a **Yangi alifbo** tab next to Home:

| Group            | Control         | What it does |
|------------------|-----------------|--------------|
| Almaştiriş       | **Körib çiqiş** | Highlights everything that will change: in the selection, or in the whole document when nothing is selected |
|                  | **Qöllaş**      | Converts what is still highlighted and removes the highlights |
|                  | **Bekor qiliş** | Removes the highlights and leaves the text as it was |
| Qaysi yozuvdan   | **Eski lotin**, **Kirill** | Which scripts Körib çiqiş looks for |
| Belgilaş rangi   | colour list     | The colour of the highlights (Word's 15 highlight colours; Feruza by default) |
| Dastur haqida    | **Maʼlumot**    | Shows the version, who made the add-in and how to contact them |

1. Open the document. To convert only part of it, select that part.
2. Press **Körib çiqiş**. The status bar says how many letters and words were found.
3. Check the highlights. To keep something as it is, such as an English word, a Russian quote or a name, remove its
   highlight: select it, then choose Home → Text Highlight Color → No Color.
4. Press **Qöllaş**. One Ctrl+Z undoes the whole conversion.

The ticked scripts and the colour are remembered the next time Word opens.

## Good to know

- **Where it looks:** body text, tables, headers and footers, footnotes and endnotes, and text boxes. It does not
  look in comments.
- **Links:** web addresses and e-mails stay as typed, for example https://shop.uz, info@maktab.uz and сайт.уз.
  A word ending joined to an address is converted: Kun.uz'га → Kun.uz'ga.
- **Formatting:** each new letter keeps the font, size, colour, bold and so on of the letter it replaces.
  Your own highlights stay as well.
- **Track Changes:** the preview highlights are never tracked. The replacements follow your Track Changes setting,
  and each Cyrillic word shows up as one change.
- **Russian text:** Russian words spelled only with letters Uzbek also has would be converted like Uzbek. Select only
  the Uzbek text, remove the highlight from the Russian parts before Qöllaş, or untick **Kirill**.
- **Protected documents:** the add-in asks you to remove the protection first. It does not work in Protected View.
- **Closing a document** takes a pending preview's highlights off, so they are never saved by mistake.
- **Known limitation:** if you yourself highlighted just an old letter pair, such as only "sh", in the preview
  colour, the add-in takes it for a leftover of an earlier preview and removes that highlight.

## Requirements

- Windows with Microsoft Word 2013 or later. It is tested with Word 2016 (32-bit).
- .NET Framework 4.8 and the Visual Studio Tools for Office runtime. The ClickOnce setup installs the runtime if it
  is missing.

To build it you also need Visual Studio 2022 with the **Office/SharePoint development** workload. To run the tests
from the command line you need the .NET SDK.

## Build and run

1. Open `NewAlphabetPlugin.sln` in Visual Studio.
2. The first time only: open the **NewAlphabetPlugin** project's Properties → **Signing**, tick
   **Sign the ClickOnce manifests** and press **Create Test Certificate**. The certificate is not kept in the repository.
3. **Build → Build Solution**. Building also registers the add-in in Word for your Windows user.
4. Press **F5** to start Word with the add-in, or just open Word.

To stop Word loading the development build, use **Build → Clean Solution**.
The version that **Maʼlumot** shows is the `AssemblyVersion` in `NewAlphabetPlugin/Properties/AssemblyInfo.cs`.
A release made from the Actions tab uses the version typed in instead; see
[Continuous integration](#continuous-integration).
To install the add-in on other computers, use **Build → Publish NewAlphabetPlugin**, which makes a ClickOnce setup.
With a test certificate, Windows asks the user whether to trust the publisher. A setup made this way is installed from
the folder it sits in and does not look for newer versions; a released one does.

## Tests

```
dotnet test NewAlphabetPlugin.Tests
```

- `AlphabetRulesTests` and `CyrillicRulesTests` test the rules without Word. They include lotin-kirill's sample news
  texts: the Cyrillic version, converted, must match the old Latin version, converted.
- `AlphabetConverterWordTests` runs the add-in's Word code against a hidden copy of Word, each test in a new
  document. On a computer without Word these tests are reported as skipped.

## Continuous integration

`.github/workflows/build.yml` runs on GitHub's `windows-2022` runner for pushes to `main`, `dev` and `ci/cd`, for
pull requests into `main` and `dev`, and when it is run from the Actions tab:

- **Test** runs `dotnet test`. The runner has no Word, so the Word tests are skipped.
- **Build the add-in** publishes the ClickOnce setup and uploads it as a build artifact named
  `NewAlphabetPlugin-<version>`. The version is the `AssemblyVersion` (or the version typed in for a release) plus
  the run number, for example `0.1.0.42`.
- **Release** runs only for a release (see below), once Test and Build pass. It attaches the zipped setup to a new
  GitHub release and then starts Publish updates.

`.github/workflows/pages.yml` (**Publish updates**) puts the newest release's setup on the GitHub Pages site,
<https://discoveruz.github.io/NewAlphabetPlugin/>. People install the add-in from there, and every Word start looks
there for a newer version. Turn the site on once, in **Settings → Pages → Source: GitHub Actions**; it is free for a
public repository. A release stops early and says so if the site is off. You can also run Publish updates from the
Actions tab, which puts the newest release on the site again.

The build signs the ClickOnce manifests with the add-in's release certificate, taken from these repository secrets:

| Secret | Value |
|--------|-------|
| `CLICKONCE_PFX_BASE64` | The `.pfx` file, Base64-encoded |
| `CLICKONCE_PFX_PASSWORD` | Its password |

Visual Studio's test certificate is only for your own builds: it is named after your computer and user name, and it
lasts one year.
Make the release certificate once, in PowerShell. Replace `Your Name` first, because it can't change later:

```
$options = @{
    Subject           = 'CN=Your Name'
    Type              = 'CodeSigningCert'
    NotAfter          = (Get-Date).AddYears(10)
    CertStoreLocation = 'Cert:\CurrentUser\My'
    Provider          = 'Microsoft Enhanced RSA and AES Cryptographic Provider'  # ClickOnce needs a CryptoAPI key
    KeySpec           = 'Signature'
    KeyExportPolicy   = 'Exportable'
}
$cert = New-SelfSignedCertificate @options
Export-PfxCertificate -Cert "Cert:\CurrentUser\My\$($cert.Thumbprint)" -FilePath "$HOME\NewAlphabetPlugin-release.pfx" -Password (Read-Host 'Choose a password' -AsSecureString)
```

Keep `NewAlphabetPlugin-release.pfx` and its password safe, outside the repository. Then set both secrets from the
repository folder; the second command asks for the password:

```
[Convert]::ToBase64String([IO.File]::ReadAllBytes("$HOME\NewAlphabetPlugin-release.pfx")) | gh secret set CLICKONCE_PFX_BASE64
gh secret set CLICKONCE_PFX_PASSWORD
```

Without the secret, branch builds are signed with a throwaway certificate and a release fails. ClickOnce only updates
an installed add-in from a version signed with the same certificate, so every release must use the same one.

There are two ways to release:

- **From the Actions tab:** open **Actions → Build → Run workflow**, keep `main`, type the version (for example
  `0.2.0`) and run it. The build uses that version and, if Test and Build pass, creates the `v0.2.0` tag and the
  release. `AssemblyInfo.cs` in the repository stays as it is.
- **With a tag:** raise `AssemblyVersion` in `NewAlphabetPlugin/Properties/AssemblyInfo.cs` in a pull request, merge
  it, and push a tag with the same version, for example `v0.2.0` for `0.2.0.0`. A tag that does not match fails the
  build.

Either way the setup ends up on the Pages site, and installed add-ins take it the next time Word starts.

## Project layout

| Path | Contents |
|------|----------|
| `NewAlphabetPlugin/AlphabetRules.cs` | Old Latin letters and link detection (no Word code) |
| `NewAlphabetPlugin/CyrillicRules.cs` | Cyrillic to new alphabet rules (no Word code) |
| `NewAlphabetPlugin/AlphabetConverter.cs` | Finding, highlighting and replacing in Word documents |
| `NewAlphabetPlugin/AlphabetRibbon.cs`, `AlphabetRibbon.Designer.cs` | The Yangi alifbo ribbon tab |
| `NewAlphabetPlugin/HighlightColors.cs` | The highlight colour list |
| `NewAlphabetPlugin/UserSettings.cs` | The remembered ribbon choices |
| `NewAlphabetPlugin/ThisAddIn.cs` | Add-in startup and Word events |
| `NewAlphabetPlugin.Tests/` | Tests; `Fixtures/` holds lotin-kirill's sample texts |
| `.github/workflows/build.yml` | The GitHub Actions build, test and release workflow |
| `.github/workflows/pages.yml` | Puts the newest release on the GitHub Pages site, where add-ins look for updates |

## Credits

The Cyrillic rules follow [lotin-kirill](https://github.com/diyorbek/lotin-kirill), and the sample texts in
`NewAlphabetPlugin.Tests/Fixtures` come from it. lotin-kirill is © 2024 Diyorbek Sadullaev, MIT licence.
