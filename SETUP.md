# Setup

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

The `generate` command calls Claude via Merck's Foundry AIP environment. Two environment variables must be set before using it. The `build` command (plan JSON → `.pptx`) does not require them.

| Variable | Required for `generate` | Description |
|---|---|---|
| `AIP_BASE_URL` | Yes | Foundry AIP endpoint URL (ask your AIP administrator) |
| `AIP_TOKEN` | Yes | Your personal Foundry API token |
| `AIP_MODEL` | No | Model override — defaults to `claude-sonnet-4-6` |

### Setting the variables

**Windows — PowerShell (persistent, current user):**

```powershell
[Environment]::SetEnvironmentVariable("AIP_BASE_URL", "https://your-instance.palantirfoundry.com/api/v1", "User")
[Environment]::SetEnvironmentVariable("AIP_TOKEN", "your-token-here", "User")
```

Restart your terminal after running these commands.

**Windows — current session only:**

```powershell
$env:AIP_BASE_URL = "https://your-instance.palantirfoundry.com/api/v1"
$env:AIP_TOKEN    = "your-token-here"
```

**bash / zsh:**

```bash
export AIP_BASE_URL="https://your-instance.palantirfoundry.com/api/v1"
export AIP_TOKEN="your-token-here"
```

Add these lines to `~/.bashrc` or `~/.zshrc` to make them permanent.

### Verify credentials

```bash
python -c "
import os
url   = os.environ.get('AIP_BASE_URL')
token = os.environ.get('AIP_TOKEN')
print('AIP_BASE_URL:', 'SET' if url   else 'MISSING')
print('AIP_TOKEN:   ', 'SET' if token else 'MISSING')
"
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

## Troubleshooting

**`ModuleNotFoundError: No module named 'merck_pptx'`**
Run all commands from the repository root directory (the folder that contains `merck_pptx/`). The package is not globally installed — Python must be able to find it on the path.

```bash
cd "Merck PPTX Template"
python -m merck_pptx --help
```

**`EnvironmentError: Missing required environment variable(s): AIP_BASE_URL, AIP_TOKEN`**
The Foundry credentials are not set. Follow the [Foundry AIP credentials](#foundry-aip-credentials) section above.

**`ValidationError: Cover slide is missing a subtitle`**
The `action_title` on your cover slide must include a subtitle, separated by a semicolon:
```json
"action_title": "Your Deck Title; Your subtitle line here"
```

**`BuildError: page N (layout_name): ...`**
A layout builder failed for slide N. The error message includes the original exception — read it for the specific cause. Common causes: a required content field is missing or has the wrong type. Check the layout's entry in [`merck_pptx/slide_plan_schema.md`](merck_pptx/slide_plan_schema.md).

**PowerPoint opens with a "needs repair" dialog**
This typically means a shape has a non-integer coordinate value. This can happen if a chart data value (e.g. `0.5`) was used directly as an EMU measurement in a custom layout. Report the issue with the plan JSON that produced it.
