:package: **QQT Script Manager** — Auto-updater for community DiabloQQT Scripts

A lightweight Windows GUI tool that keeps your QQT scripts up to date straight from GitHub — no manual downloading required.

**Place script in the same location your diabloqqt exe is. The powershell script and your scripts folder should be in the same directory.**

**What it does**
> • Checks GitHub for the latest version of each script
> • Downloads and installs/updates scripts with one click
> • Tracks installed versions so it only downloads what changed
> • Preserves your personal settings files on update (only changed files are replaced)
> • Supports **collection packs** (e.g. War-Pig-Zewx, D4QQT) that bundle multiple scripts in one repo

> :gear: **First-time setup — allow the script to run**
> > Open PowerShell as an Administrator and run the following command before launching the tool:
```
function test() {
  Set-ExecutionPolicy unrestricted
}
```
> > You only need to do this once. When prompted, press **Y** to confirm.
> > Alternatively, You can run powershell with the unrestricted powershell Execution Policy if you don't want to/cant change it globally.
> 

**How to use it**
> 1. Run `QQTScriptManager.ps1` in PowerShell
> 2. The GUI will open — installed scripts show a ✔ and their update status
> 3. **Double-click** a script or right-click → Install/Update to install it
> 4. Use **Update All Installed** to update everything at once
> 5. Use **Force Refresh** to recheck GitHub for new versions

**Other features**
> • Add any public GitHub repo with **Add Repository**
> • Set a **GitHub Token** to avoid API rate limits (Settings → GitHub Token)
> • Right-click → **Uninstall** to remove a script and its files


:warning: *Requires PowerShell 5.1+ on Windows. No external modules needed.*
