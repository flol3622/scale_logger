## Astral-sh UV

### Getting started

```bash
# On macOS and Linux.
curl -LsSf https://astral.sh/uv/install.sh | sh

# On Windows.
powershell -c "irm https://astral.sh/uv/install.ps1 | iex"

# With pip.
pip install uv
```

### Create virtual environment

```bash
uv venv  # Create a virtual environment at .venv.
```

### Activate virtual environment

```bash
# On macOS and Linux.
source .venv/bin/activate

# On Windows.
.venv\Scripts\activate
```

# Install packages

```bash
uv pip install flask                # Install Flask.
uv pip install -r requirements.txt  # Install from a requirements.txt file.
uv pip install -e .                 # Install the current project in editable mode.
uv pip install "package @ ."        # Install the current project from disk.
uv pip install "flask[dotenv]"      # Install Flask with "dotenv" extra.
```

### Generate requirements.txt

```bash
uv pip freeze | uv pip compile - -o requirements.txt  # Lock the current environment.
```

### Update to requirements.txt

```bash
uv pip sync requirements.txt
```