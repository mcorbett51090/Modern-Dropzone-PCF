# Modern Dropzone PCF — SharePoint Only Mode
## Custom Build: 2026.4.8

---

## What Was Changed

| File | Change |
|---|---|
| `Dropzone/ControlManifest.Input.xml` | Added `sharePointOnlyMode` property (TwoOptions, default `true`); version → `2026.4.8` |
| `Dropzone/index.ts` | Version log string → `2026.4.8` |
| `Dropzone/Landing/Landing.tsx` | SP-only mode logic: forced SP state, hidden toggle, hidden dropdown, hidden callout, Notes drops disabled |
| `Solution/ModernDropzone/src/Other/Solution.xml` | Solution version → `2026.4.8` |
| `Solution/ModernDropzone/ModernDropzone.cdsproj` | No logic change; cleaned up for managed output |

### Behaviour when `sharePointOnlyMode = Yes` (default)
- Control starts in SharePoint mode immediately — no toggle, no Notes
- The SharePoint/Notes toggle is **not rendered**
- The document-location dropdown is **not rendered**
- The "Remember Location" settings gear/callout is **not rendered**
- `allowNoteDrops` is overridden to `false` regardless of form property
- `allowSharePointDrops` is driven only by whether the control is disabled
- New/unsaved records still show the "record not yet created" message (unchanged)

### Behaviour when `sharePointOnlyMode = No`
- Full original behaviour restored — toggle, dropdown, and callout all reappear

---

## Prerequisites

Install these once on your build machine if not already present:

```
node --version   # 18.x or 20.x LTS recommended
npm --version    # 10.x
dotnet --version # .NET 4.6.2 SDK (for msbuild solution step)
```

Install the Power Platform CLI:
```
npm install -g pac
```

---

## Step 1 — Copy Custom Files Into Your Fork

Copy these four files from this package into the matching paths of your
cloned `GorgonUK/Modern-Dropzone-PCF` repository:

```
Dropzone/ControlManifest.Input.xml
Dropzone/index.ts
Dropzone/Landing/Landing.tsx
Solution/ModernDropzone/src/Other/Solution.xml
Solution/ModernDropzone/ModernDropzone.cdsproj
```

> Do NOT overwrite any other files (CSS, .resx, DataverseActions.ts, utils.ts, etc.)
> Only the five files listed above need replacing.

---

## Step 2 — Install npm Dependencies

From the repo root (where `package.json` lives):

```bash
npm install
```

---

## Step 3 — Build the PCF Control

```bash
npm run build
```

Expected output: no errors; a `bundle.js` (or equivalent) is emitted under
`out/controls/Dropzone/`.

If you see TypeScript errors about missing generated types, run:

```bash
npm run refreshTypes
npm run build
```

---

## Step 4 — Build the Managed Solution (.zip)

The solution project lives at `Solution/ModernDropzone/ModernDropzone.cdsproj`.

### Option A — msbuild (Windows / Visual Studio)

Open a Developer Command Prompt and run from the repo root:

```cmd
cd Solution\ModernDropzone
msbuild /t:build /restore
```

The managed solution zip is written to:
```
Solution\ModernDropzone\bin\Debug\ModernDropzone.zip
```

### Option B — Power Platform CLI (pac)

If you prefer `pac` over msbuild:

```bash
# 1. Push the PCF bundle first
pac pcf push --publisher-prefix guk

# 2. Then export the solution (after it's been imported and customised once)
pac solution export --path ./ModernDropzone_managed.zip --managed true
```

---

## Step 5 — Import Into Power Platform

### Via Power Platform Admin Center / make.powerapps.com

1. Navigate to **Solutions** → **Import solution**
2. Browse to `ModernDropzone.zip`
3. Click **Next** → **Import**
4. Wait for the import to complete (usually < 2 minutes)

### Via Power Platform CLI

```bash
pac auth create --url https://yourorg.crm.dynamics.com
pac solution import --path ModernDropzone.zip --async
```

---

## Step 6 — Configure the Component on the Form

1. Open the target form in the Power Apps form editor
2. If upgrading from a previous version, the component is already on the form —
   just **publish** and the new `sharePointOnlyMode` property appears with
   default `Yes`.  No further configuration needed.
3. If adding fresh:
   - **Get more components** → search "Dropzone" → Add
   - Drag the component onto the form
   - In the component properties pane:
     - **SharePoint Only Mode?** → `Yes` (default — no change needed)
     - **Enable SharePoint Integration?** → `Yes`
     - All other properties can remain at defaults
4. **Save and Publish** the form

---

## Step 7 — Verify in the App

1. Open a record for an entity that has SharePoint Document Management enabled
2. The Dropzone should show **only the drag-and-drop area** — no toggle, no
   document-location dropdown, no gear icon
3. Drop a file — it should upload directly to the entity's SharePoint
   Document Location
4. Files should appear in the dropzone and in the SharePoint Documents subgrid

---

## Reverting to Classic Mode

On any form where you need the toggle back:

1. Open the form editor
2. Select the Dropzone component → Properties
3. Set **SharePoint Only Mode?** to `No`
4. Save and Publish

---

## Troubleshooting

| Symptom | Likely cause | Fix |
|---|---|---|
| Build error: "Cannot find module '../DataverseActions'" | Custom files copied but original files missing | Ensure you only replaced the 5 listed files; all others must remain |
| Build error: "Property 'sharePointOnlyMode' does not exist on type 'IInputs'" | Generated manifest types stale | Run `npm run refreshTypes` then `npm run build` |
| Toggle still visible after import | Browser cache or old component version on form | Hard-refresh (Ctrl+Shift+R); republish the form |
| Files not appearing after drop | SharePoint integration not enabled on the entity | Follow [MS Docs — Enable SharePoint document management](https://learn.microsoft.com/en-us/power-platform/admin/enable-sharepoint-document-management-specific-entities) |
| "Record hasn't been created yet" message | Uploading to a new unsaved record | Save the record first — this is expected PCF behaviour |

---

## Version History

| Version | Date | Notes |
|---|---|---|
| 2026.4.8 | 2026-04-08 | Added `sharePointOnlyMode`; hides toggle/dropdown/callout; disables Note drops |
| 0.0.77 / 2.9.5 | (upstream) | Original upstream release |
