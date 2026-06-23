# Merck Corporate Design PPTX Tool — Setup

## Prerequisites

- **Python 3.11** — the pipeline is tested on 3.11; other 3.x versions may work but are not validated
- **Merck Foundry AIP access** — required only for the `generate` command (markdown / PowerPoint input); the `build` command works without any credentials

---

## Installation

### 1. Get the code

```bash
git clone <repo-url>
cd "Merck PPTX Template"
```

Or download and unzip the repository to a local folder.

### 2. Activate your Python environment

If you are using the shared `ds_env` conda environment:

```bash
conda activate ds_env
```

To create a fresh environment instead:

```bash
conda create -n merck_pptx python=3.11
conda activate merck_pptx
```

### 3. Install dependencies

```bash
pip install -r requirements.txt
```

### 4. Verify the installation

```bash
python -c "from merck_pptx import generate_deck, build_from_plan; print('OK')"
python -m merck_pptx --help
```

Expected output:

```
OK
usage: python -m merck_pptx [-h] {generate,build,discover-templates,list-templates,register-template} ...
```

### 5. (Optional) Run the setup script

Run `setup_merck_pptx.py` to auto-detect the empower add-in, scan all available templates, and print instructions for adding templates manually:

```bash
python setup_merck_pptx.py
```

This is not required if you only use `build` with `division: merck` (the bundled templates always work). It is recommended for Mac/Linux users and anyone who needs non-`merck` division templates.

---

## Foundry AIP credentials

The `generate` command calls Claude via Merck's internal AI platform. The `build` command (plan JSON to `.pptx`) does not require any credentials.

### Environment variables (checked in priority order)

| Primary name | Fallback name | Required for `generate` | Description |
|---|---|---|---|
| `AIP_BASE_URL` | `ANTHROPIC_BASE_URL` | Yes | AI platform endpoint URL |
| `AIP_TOKEN` | `ANTHROPIC_AUTH_TOKEN` | Yes | Your personal API token |
| `AIP_MODEL` | — | No | Model override (default: `claude-sonnet-4-6`) |

The pipeline checks the primary name first, then falls back to the alternative name. On Windows, it also searches `HKCU\Environment` in the registry, so variables set via PowerShell's `[Environment]::SetEnvironmentVariable` are found automatically without restarting the terminal.

### Setting the variables

> **Security note:** Never paste a token as a literal in a command — it will be saved to your shell history. Use the masked-input forms below instead.

**Windows — PowerShell (persistent, current user):**

```powershell
# AIP_BASE_URL is not sensitive — paste the URL directly
[Environment]::SetEnvironmentVariable("AIP_BASE_URL", "https://your-instance.example.com/api/v1", "User")

# AIP_TOKEN — entered via a masked prompt so it is never written to history
$_t = Read-Host "Paste your AIP_TOKEN" -AsSecureString
[Environment]::SetEnvironmentVariable(
    "AIP_TOKEN",
    [Runtime.InteropServices.Marshal]::PtrToStringAuto(
        [Runtime.InteropServices.Marshal]::SecureStringToBSTR($_t)),
    "User")
Remove-Variable _t
```

The pipeline reads these from the registry on next run — no terminal restart needed.

**Windows — current session only:**

```powershell
$env:AIP_BASE_URL = "https://your-instance.example.com/api/v1"
$env:AIP_TOKEN    = [Runtime.InteropServices.Marshal]::PtrToStringAuto(
    [Runtime.InteropServices.Marshal]::SecureStringToBSTR(
        (Read-Host "Paste your AIP_TOKEN" -AsSecureString)))
```

**bash / zsh (current session):**

```bash
export AIP_BASE_URL="https://your-instance.example.com/api/v1"
read -rs AIP_TOKEN && export AIP_TOKEN   # hidden input, not saved to history
```

To persist across sessions, add the `AIP_BASE_URL` export to `~/.bashrc` or `~/.zshrc`. Run the `read -rs` line each time you open a terminal. Do not write the token literal into your shell profile.

### Verify credentials

```bash
python -c "
import os, sys
# Check both primary and fallback names
url   = os.environ.get('AIP_BASE_URL')   or os.environ.get('ANTHROPIC_BASE_URL')
token = os.environ.get('AIP_TOKEN')      or os.environ.get('ANTHROPIC_AUTH_TOKEN')
if sys.platform == 'win32' and (not url or not token):
    try:
        import winreg
        k = winreg.OpenKey(winreg.HKEY_CURRENT_USER, 'Environment')
        if not url:
            try: url, _ = winreg.QueryValueEx(k, 'AIP_BASE_URL')
            except FileNotFoundError:
                try: url, _ = winreg.QueryValueEx(k, 'ANTHROPIC_BASE_URL')
                except FileNotFoundError: pass
        if not token:
            try: token, _ = winreg.QueryValueEx(k, 'AIP_TOKEN')
            except FileNotFoundError:
                try: token, _ = winreg.QueryValueEx(k, 'ANTHROPIC_AUTH_TOKEN')
                except FileNotFoundError: pass
    except Exception: pass
print('Endpoint:', 'SET' if url   else 'MISSING')
print('Token:   ', 'SET' if token else 'MISSING')
"
```

---

## Non-interactive / scripted usage

The `generate` command can run without any prompts in three ways:

```bash
# 1. Provide answers as a JSON file
python -m merck_pptx generate source.md output.pptx --meta meta.json

# 2. Use built-in defaults (EU, Internal, Mixed audience, merck_corporate, organic)
python -m merck_pptx generate source.md output.pptx --defaults

# 3. Pipe stdin or run in CI — defaults activate automatically when stdin is not a TTY
echo "" | python -m merck_pptx generate source.md output.pptx
```

The `build` command is always non-interactive.

### Saving the generated plan

Use `--save-plan` to write the LLM-generated slide plan to a JSON file alongside the deck. This lets you inspect the plan, make manual edits, and rebuild without another LLM call:

```bash
python -m merck_pptx generate source.md output.pptx \
    --meta meta.json \
    --save-plan plan.json

# Edit plan.json, then rebuild deterministically:
python -m merck_pptx build plan.json output_revised.pptx
```

---

## Requirements

```
python-pptx>=0.6.21
anthropic>=0.25.0
```

`python-pptx` pulls in `lxml` and `Pillow` automatically as transitive dependencies — no separate installation needed.

---

## Updating

Pull the latest version and re-install dependencies:

```bash
git pull
pip install -r requirements.txt --upgrade
```

---

## Template files

The pipeline resolves templates in this priority order:

1. **empower BinaryFiles** (primary, Windows only) — exact per-theme PPTX files with correct design shapes baked in. Requires the empower add-in to be installed.
2. **Manual templates** (`~/.merck_pptx/templates/`) — user-placed files for Mac/Linux users or selective Windows installs. See [Manual templates](#manual-templates-macos--linux-and-selective-download) below.
3. **Division static templates** (`merck_pptx/templates/`) — `.pptx` files committed to the repository, one per division/region. Only `merck` division files are bundled.
4. **Region default** (bundled) — `EU_Merck_Themed.pptx` (EU) or `USA_Merck_Themed_Base_v1.pptx` (USA). **Available only for `division: merck`.**

> **Important:** For divisions other than `merck` (e.g. `emd_serono`, `millipore_sigma`), the pipeline raises a `TemplateNotFoundError` if no template is found in steps 1–2. It does **not** silently fall back to the generic Merck template, because that would produce slides with the wrong logo and branding.
>
> Run `python -m merck_pptx list-templates` at any time to see what is available on your machine.

### empower BinaryFile templates (Windows, recommended)

If your organisation uses empower and you have BinaryFiles installed, register them for exact per-theme rendering:

```bash
# Discover all available BinaryFile templates (grouped by color theme)
python -m merck_pptx discover-templates

# Register a specific UID for a division + color theme
python -m merck_pptx register-template <uid> merck plastic
python -m merck_pptx register-template <uid> emd_serono organic
```

`<uid>` is the BinaryFile identifier shown by `discover-templates` (the filename without `.pptx`). Registration is stored in `merck_pptx/binary_registry.json`.

### Manual templates (macOS / Linux and selective download)

Mac and Linux users (and Windows users without empower) can download individual template files and place them in a local folder. The pipeline checks this folder automatically.

**Directory:** `~/.merck_pptx/templates/`

**Naming convention:** `{division}_{color_theme}.pptx`

| Examples | |
|---|---|
| `merck_organic.pptx` | Generic Merck, organic (cream/red) theme |
| `merck_plastic.pptx` | Generic Merck, plastic (green/pink) theme |
| `emd_serono_organic.pptx` | EMD Serono, organic theme |
| `millipore_sigma_functional.pptx` | MilliporeSigma, functional (green/teal) theme |

**Valid division keys:**

| Key | Brand |
|---|---|
| `merck` | Merck KGaA (EU/global) — *bundled, always available* |
| `merck_asia` | Merck 默克 (Asia/China branding) |
| `emd_serono` | EMD Serono (USA/Canada Healthcare) |
| `millipore_sigma` | MilliporeSigma (USA/Canada Life Science) |
| `emd_electronics` | EMD Electronics (USA/Canada) |
| `usa` | USA tri-brand (cross-business) |

**Valid color_theme keys:** `plastic`  `functional`  `organic`  `synthetic`  `technical`  `electronics`

**How to get the template files:**

1. Ask a Windows colleague who has empower installed to open PowerPoint.
2. Go to: **empower tab → Corporate Design Templates → Master Templates → Merck** (or the relevant division).
3. Right-click the desired template → **Export to file** → save as `.pptx`.
4. Rename the file to `{division}_{color_theme}.pptx` and place it in `~/.merck_pptx/templates/`.
5. Verify it is recognised:
   ```bash
   python -m merck_pptx list-templates
   ```

> **You do not need to download all templates** — only the ones you will actually use. `division: merck` always works without any download.

### Custom manual templates directory

To use a different folder, set `manual_templates.dir` in your `config.yaml`:

```yaml
manual_templates:
  dir: "/path/to/my/templates"
```

### Registering additional empower BinaryFile templates

If `discover-templates` shows UIDs that are not yet registered:

```bash
python -m merck_pptx discover-templates      # shows registered vs. unregistered
python -m merck_pptx register-template <uid> <division> <color_theme>
```

---

## Troubleshooting

**`TemplateNotFoundError: Template not available: division='X', color_theme='Y', region='Z'`**
The pipeline found no template for the requested division + color theme combination. Fix options:
- **Windows:** Install the empower add-in, then re-run `python setup_merck_pptx.py`.
- **Mac/Linux (or no empower):** Download the template file from a Windows machine (see [Manual templates](#manual-templates-macos--linux-and-selective-download)) and place it at `~/.merck_pptx/templates/{division}_{color_theme}.pptx`.
- **Quick fix:** Change `meta.division` to `merck` in your plan — the bundled EU/USA Merck template is always available for all 6 color themes.
- Run `python -m merck_pptx list-templates` to see what is currently available on your machine.

**`ModuleNotFoundError: No module named 'merck_pptx'`**
Run all commands from the repository root directory (the folder that contains `merck_pptx/`). The package is not globally installed — Python must be able to find it on the path.

```bash
cd "Merck PPTX Template"
python -m merck_pptx --help
```

**`EnvironmentError: Missing required environment variable(s): AIP_BASE_URL, AIP_TOKEN`**
The AI platform credentials are not set. The pipeline checks both `AIP_BASE_URL`/`AIP_TOKEN` and `ANTHROPIC_BASE_URL`/`ANTHROPIC_AUTH_TOKEN`. On Windows, it also reads `HKCU\Environment` from the registry. See [Foundry AIP credentials](#foundry-aip-credentials) above.

**`BuildError: page N (layout_name): ...`**
A layout builder failed for slide N. The error message includes the original exception — read it for the specific cause. Common causes: a required content field is missing or has the wrong type. Check the layout's entry in [`merck_pptx/slide_plan_schema.md`](merck_pptx/slide_plan_schema.md).

**`ValidationError: ...`**
The slide plan contains a hard schema violation. Common causes:
- `action_title` on the cover slide exceeds 60 characters
- `section_number` is duplicated across slides
- An unknown `layout` key is used

If the plan was generated by the LLM, retry — the model occasionally produces slightly non-conforming plans, and the pipeline retries automatically up to 3 times.

**PowerPoint opens with a "needs repair" dialog**
This typically means a shape has a non-integer coordinate value. This can happen if a chart data value (e.g. `0.5`) was used directly as an EMU measurement in a custom layout. Report the issue with the plan JSON that produced it.

**Takeaway warnings in output**
Warnings like `Slide page=N: takeaway is X chars (approaching 120-char limit)` indicate the LLM generated a takeaway close to the rendering limit. Text above ~90 characters may wrap onto a second line that gets clipped by the fixed-height takeaway band. Shorten the takeaway in the plan JSON and rebuild with `python -m merck_pptx build`.
