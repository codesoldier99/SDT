# Course Website (MkDocs Material)

## Local setup

```bash
pip install -r /home/runner/work/SDT/SDT/website/requirements.txt
mkdocs serve -f /home/runner/work/SDT/SDT/website/mkdocs.yml
```

## Build

```bash
mkdocs build -f /home/runner/work/SDT/SDT/website/mkdocs.yml
```

## Versioning (mike)

```bash
mike deploy --config-file /home/runner/work/SDT/SDT/website/mkdocs.yml --push 2026-spring latest
mike set-default --config-file /home/runner/work/SDT/SDT/website/mkdocs.yml --push latest
```
