
# code2docx

`code2docx` is a command-line utility for instructors and content creators who need to quickly prepare code examples and screenshots for teaching materials.
The tool performs a complete workflow:

1. **Processes code files** and removes hidden sections based on markers you define.
2. **Processes screenshots** with similar hide rules.
3. **Assembles all resulting screenshots into a single PDF**, arranged **two images per row**, optimized for fast copy/paste into Word or Google Docs.

This eliminates the repetitive manual tasks normally required when preparing programming lessons or lab sheets for students.

---

## Features

* Hide selected parts of code using simple start/end markers.
* Apply the same hiding logic to screenshot images.
* Export all screenshots into a clean, consistent PDF layout (2 images per row).
* Run from any directory and target any project path.
* Lightweight and easy to integrate into your workflow.

---

## Project Structure

```
code2docx/
├── pyproject.toml
├── code2docx/
│   ├── __init__.py
│   └── cli.py
```

---

## Installation

### Editable install (recommended for development)

```bash
pip install -e .
```

### Normal install

```bash
pip install .
```

### Install directly from a Git repository

```bash
pip install git+https://github.com/<USERNAME>/code2docx.git
```

---

## Step 4: Use from anywhere

After installation, the `code2docx` CLI becomes globally available.

```bash
# Move to any project and run:
cd /path/to/your/project
code2docx
```

Or specify a path explicitly:

```bash
code2docx /path/to/project
```

---

## Usage Examples

Basic usage:

```bash
code2docx \
    --code src/*.py \
    --screenshots screenshots/*.png \
    --output final.pdf
```

Using custom hide markers and keeping temporary intermediate files:

```bash
code2docx \
    --code code/*.js \
    --screenshots shots/*.jpg \
    --hide-start "// hide-start" \
    --hide-end "// hide-end" \
    --keep-temp \
    -o lesson1.pdf
```

---

## Hide Markers

You can hide specific code sections using markers such as:

```python
print("visible")
# hide-start
print("this part will be removed")
# hide-end
print("visible again")
```

Hiding rules for screenshots follow the same start/end markers based on your tool configuration.

---

## Output

The generated PDF:

* Includes all processed screenshots.
* Places them **two images per row** for compact, readable layout.
* Is formatted for rapid insertion into Word or Google Docs.

---

## License

MIT License (or your preferred license).

---

## Contributing

Pull requests and issues are welcome.

---
