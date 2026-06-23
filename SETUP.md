# Merck Corporate Design PPTX Tool — Setup

## Prerequisites

- **Python 3.9+** (tested on 3.11)
- **Merck Foundry AIP access** — required only for `generate`; `build` works without credentials

---

## Installation

```bash
git clone <repo-url>
cd "Merck PPTX Template"
conda activate ds_env          # or your preferred Python 3.9+ environment
python setup_merck_pptx.py
```

The setup script installs dependencies, auto-detects your empower installation, writes a machine-local `config.yaml`, and reports which templates are available.

---

## Foundry AIP credentials

Required for `generate` (markdown / PPTX input). Not needed for `build`.

Set these as persistent user environment variables:

| Variable | Purpose |
|---|---|
| `AIP_BASE_URL` | Foundry endpoint, e.g. `https://merck.palantirfoundry.com/api/v1` |
| `AIP_TOKEN` | Your personal Foundry API token |
| `AIP_MODEL` | Optional model override (default: `claude-sonnet-4-6`) |

**Windows — PowerShell:**

```powershell
[Environment]::SetEnvironmentVariable("AIP_BASE_URL", "https://your-instance.example.com/api/v1", "User")

$_t = Read-Host "Paste AIP_TOKEN" -AsSecureString
[Environment]::SetEnvironmentVariable("AIP_TOKEN",
    [Runtime.InteropServices.Marshal]::PtrToStringAuto(
        [Runtime.InteropServices.Marshal]::SecureStringToBSTR($_t)), "User")
Remove-Variable _t
```

**bash / zsh:**

```bash
export AIP_BASE_URL="https://your-instance.example.com/api/v1"
read -rs AIP_TOKEN && export AIP_TOKEN
```

The pipeline also accepts `ANTHROPIC_BASE_URL` / `ANTHROPIC_AUTH_TOKEN` as fallback names, and reads Windows registry (`HKCU\Environment`) automatically — no terminal restart needed.

---

## Non-interactive / scripted usage

```bash
# Provide meta answers as a JSON file
python -m merck_pptx generate source.md output.pptx --meta meta.json

# Use built-in defaults (EU, Internal, mixed audience, merck_corporate, organic)
python -m merck_pptx generate source.md output.pptx --defaults

# Save the generated plan for inspection / manual editing / deterministic rebuild
python -m merck_pptx generate source.md output.pptx --meta meta.json --save-plan plan.json
python -m merck_pptx build plan.json output_revised.pptx
```

---

## Templates

Run `python -m merck_pptx list-templates` to see what is available on your machine.

### division: merck (always available)

The generic Merck EU and USA templates are bundled — no download needed, all 6 color themes work out of the box.

### Other divisions (emd_serono, millipore_sigma, …)

The pipeline raises a `TemplateNotFoundError` rather than silently falling back to the wrong logo/branding. You need a template file for each division you use.

**Windows with empower installed** — templates are auto-detected. The setup script registers them and `discover-templates` / `register-template` let you add more:

```bash
python -m merck_pptx discover-templates            # list all available BinaryFile UIDs
python -m merck_pptx register-template <uid> emd_serono organic
```

**macOS / Linux (or Windows without empower)** — place template files in `~/.merck_pptx/templates/`, named `{division}_{color_theme}.pptx`:

```
~/.merck_pptx/templates/
  merck_organic.pptx
  emd_serono_plastic.pptx
  millipore_sigma_functional.pptx
```

To get the files: ask a Windows colleague with empower to open PowerPoint → **empower tab → Corporate Design Templates → Master Templates → Merck** (or the relevant division) → right-click a template → **Export to file** → save as `.pptx`. Rename to the convention above and copy it over.

You only need the templates you will actually use. `division: merck` never requires a download.

**Valid division keys:** `merck` · `merck_asia` · `emd_serono` · `millipore_sigma` · `emd_electronics` · `usa`

**Valid color_theme keys:** `plastic` · `functional` · `organic` · `synthetic` · `technical` · `electronics`

To use a custom directory instead of `~/.merck_pptx/templates/`, set `manual_templates.dir` in `config.yaml`.

---

## Updating

```bash
git pull
pip install -r requirements.txt --upgrade
python setup_merck_pptx.py
```

---

## Troubleshooting

**`TemplateNotFoundError`** — Run `python -m merck_pptx list-templates`. Either install empower, add a manual template, or switch to `division: merck`.

**`ModuleNotFoundError: No module named 'merck_pptx'`** — Run commands from the repo root directory (the folder containing `merck_pptx/`).

**`EnvironmentError: Missing required environment variable(s): AIP_BASE_URL, AIP_TOKEN`** — Set credentials as shown above.

**`BuildError: page N (layout_name): ...`** — A required content field is missing or has the wrong type. Check [`merck_pptx/slide_plan_schema.md`](merck_pptx/slide_plan_schema.md) for the layout's schema.

**`ValidationError`** — The plan has a hard schema violation (e.g. duplicate `section_number`, unknown layout key). If LLM-generated, retry — the pipeline retries automatically up to 3 times.

**PowerPoint "needs repair" dialog** — Report with the plan JSON that produced it.

**Takeaway warnings** — Takeaways over ~90 characters may clip. Shorten in the plan JSON and rebuild with `python -m merck_pptx build`.
