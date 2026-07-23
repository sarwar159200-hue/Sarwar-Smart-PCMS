# A–Z: Build and Share Sarwar Smart PCMS Without Installing Visual Studio

1. Create a free GitHub account.
2. Create a new repository named `Sarwar-Smart-PCMS`.
3. Extract this ZIP.
4. Upload all extracted folders and files to the repository.
5. Open the repository's **Actions** tab.
6. Enable GitHub Actions when prompted.
7. Open **Build Sarwar Smart PCMS x86**.
8. Click **Run workflow**.
9. Wait for the green check mark.
10. Open the completed workflow and download the artifact named `Sarwar-Smart-PCMS-x86`.
11. Extract the artifact. It contains the generated XLL and installer when the build succeeds.
12. Run `Sarwar-Smart-PCMS-Setup-x86.exe`.
13. Close Excel before installation.
14. Reopen Excel. The **Sarwar Smart PCMS** tab should appear automatically.
15. Share the same setup file with colleagues who use 32-bit Excel.

## Publishing a version
1. In GitHub, create a release tag such as `v1.0.0`.
2. The workflow builds the installer and attaches it to the GitHub Release.
3. Share the Release page or the installer with colleagues.

## Security notes
- Do not disable Microsoft Defender.
- Do not add broad trusted folders unnecessarily.
- Digitally sign the XLL and installer before public commercial distribution.
- Test on a spare Windows account and with a sample workbook first.
