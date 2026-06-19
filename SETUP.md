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
usage: python -m merck_pptx [-h] {generate,build} ...
```

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

# 2. Use built-in defaults (EU, Internal, Executive leadership, merck_executive)
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

The pipeline ships with two templates (EU and USA defaults). Division-specific templates must be added manually.

### Division × region file names

| `division` | EU template file | USA template file |
|---|---|---|
| `merck` *(default)* | `EU_Merck_Themed.pptx` ✅ included | `USA_Merck_Themed_Base_v1.pptx` ✅ included |
| `emd_electronics` | `EU_EMDElectronics_Themed.pptx` | `USA_EMDElectronics_Themed.pptx` |
| `emd_serono` | `EU_EMDSerono_Themed.pptx` | `USA_EMDSerono_Themed.pptx` |
| `millipore_sigma` | `EU_MilliporeSigma_Themed.pptx` | `USA_MilliporeSigma_Themed.pptx` |
| `merck_asia` | `EU_MerckAsia_Themed.pptx` | `USA_MerckAsia_Themed.pptx` |

If a division-specific file is missing the pipeline falls back to the region default — no error is raised.

### Adding a division template

1. In empower, go to **Corporate Design Templates → Master Templates → {Division}** and export the master template as a `.pptx` file.
2. Name the file according to the table above.
3. Place it in `merck_pptx/templates/`.

The pipeline picks it up automatically on the next run.

---

## Troubleshooting

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
